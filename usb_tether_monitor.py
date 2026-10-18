#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
USB 网络共享 / 投屏流量监控（桌面悬浮窗）
- 自动识别 Windows 下 USB 网络共享网卡 (描述含 RNDIS / USB 的网卡)
- 桌面悬浮窗实时显示下载/上传速度、本次累计（按方向）、本次总量与监测时长
- 悬浮窗为 UpdateLayeredWindow 逐像素 Alpha 渲染：80% 深色淡底可全程拖动，文字带暗影
- 本次总量 = 自监控启动以来该网卡的累计收发总量（psutil 实时统计，准确）
用法:
    python usb_tether_monitor.py                # 桌面悬浮窗（默认模式）
    python usb_tether_monitor.py --iface "以太网 2"   # 指定网卡
依赖: psutil, pywin32, Pillow (悬浮窗用系统 Python3.12 自带 tkinter)
"""
import argparse
import ctypes
import os
import sys
import threading
import time

import psutil

# 可选依赖：网卡检测（pywin32）；悬浮窗渲染（Pillow）
try:
    import win32com.client
    HAVE_WIN32 = True
except Exception:
    HAVE_WIN32 = False

try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter
    HAVE_PIL = True
except Exception:
    HAVE_PIL = False

try:
    import tkinter as tk
    HAVE_GUI = True
except Exception:
    HAVE_GUI = False

# ============ 悬浮窗渲染：UpdateLayeredWindow 逐像素 Alpha ============
# 为什么不用 Tk 的 -transparentcolor（色键）：实测色键像素会被系统命中测试
# 排除（点击穿透到桌面），导致悬浮窗背景区域无法拖动；且文字直接贴在壁纸上，
# 浅色背景下不清晰。改用 ULW 逐像素 alpha：
#   · 整卡以极低 alpha（>0）铺底 → 视觉近透明，但整卡都可接收鼠标事件；
#   · 文字带柔和暗影 + 真逐像素抗锯齿 → 任何背景下清晰。
FLOAT_W, FLOAT_H = 188, 96    # 悬浮窗逻辑尺寸（96 DPI 像素）

try:
    from ctypes import wintypes as _wintypes

    class _BIHEAD(ctypes.Structure):
        _fields_ = [("biSize", _wintypes.DWORD), ("biWidth", _wintypes.LONG),
                    ("biHeight", _wintypes.LONG), ("biPlanes", _wintypes.WORD),
                    ("biBitCount", _wintypes.WORD),
                    ("biCompression", _wintypes.DWORD),
                    ("biSizeImage", _wintypes.DWORD),
                    ("biXPelsPerMeter", _wintypes.LONG),
                    ("biYPelsPerMeter", _wintypes.LONG),
                    ("biClrUsed", _wintypes.DWORD),
                    ("biClrImportant", _wintypes.DWORD)]

    class _BIMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", _BIHEAD), ("bmiColors", _wintypes.DWORD * 3)]

    class _BLENDFUNCTION(ctypes.Structure):
        _fields_ = [("BlendOp", _wintypes.BYTE), ("BlendFlags", _wintypes.BYTE),
                    ("SourceConstantAlpha", _wintypes.BYTE),
                    ("AlphaFormat", _wintypes.BYTE)]

    class _ULWSIZE(ctypes.Structure):
        _fields_ = [("cx", _wintypes.LONG), ("cy", _wintypes.LONG)]

    class _ULWPOINT(ctypes.Structure):
        _fields_ = [("x", _wintypes.LONG), ("y", _wintypes.LONG)]

    class _ULWRECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                    ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    _WS_EX_LAYERED = 0x00080000
    _SWP_NOSIZE = 0x0001
    _SWP_NOZORDER = 0x0004
    _SWP_NOACTIVATE = 0x0010
    HAVE_ULW = True
except Exception:
    HAVE_ULW = False


def ulw_toplevel(root):
    """取 Tk 根窗口的真实顶层 HWND（winfo_id 是 TkChild）。"""
    try:
        child = root.winfo_id()
        return ctypes.windll.user32.GetParent(child) or child
    except Exception:
        return None


def ulw_enable(hwnd):
    """给窗口加 WS_EX_LAYERED 样式。"""
    if not HAVE_ULW or not hwnd:
        return False
    try:
        user32 = ctypes.windll.user32
        ex = user32.GetWindowLongW(hwnd, -20)  # GWL_EXSTYLE
        user32.SetWindowLongW(hwnd, -20, ex | _WS_EX_LAYERED)
        return True
    except Exception:
        return False


def ulw_rect(hwnd):
    """窗口屏幕坐标 (l, t, r, b)，失败返回 None。"""
    if not HAVE_ULW or not hwnd:
        return None
    try:
        r = _ULWRECT()
        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(r))
        return (r.left, r.top, r.right, r.bottom)
    except Exception:
        return None


def ulw_paint(hwnd, img, x, y):
    """把 PIL RGBA 图像以逐像素 alpha 贴到分层窗口。"""
    if not HAVE_ULW or not hwnd or img is None:
        return False
    w, h = img.size
    try:
        bmi = _BIMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(_BIHEAD)
        bmi.bmiHeader.biWidth = w
        bmi.bmiHeader.biHeight = -h            # 自顶向下
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0        # BI_RGB
        bits = ctypes.c_void_p()
        user32 = ctypes.windll.user32
        gdi32 = ctypes.windll.gdi32
        hdc_screen = user32.GetDC(0)
        hdc_mem = gdi32.CreateCompatibleDC(None)
        hbmp = gdi32.CreateDIBSection(hdc_mem, ctypes.byref(bmi), 0,
                                      ctypes.byref(bits), None, 0)
        if not hbmp:
            gdi32.DeleteDC(hdc_mem)
            user32.ReleaseDC(0, hdc_screen)
            return False
        gdi32.SelectObject(hdc_mem, hbmp)
        raw = img.tobytes("raw", "BGRA")
        ctypes.memmove(bits, raw, len(raw))
        size = _ULWSIZE(w, h)
        src = _ULWPOINT(0, 0)
        dst = _ULWPOINT(int(x), int(y))
        bf = _BLENDFUNCTION(1, 0, 255, 1)      # AC_SRC_ALPHA
        ok = user32.UpdateLayeredWindow(hwnd, hdc_screen, ctypes.byref(dst),
                                        ctypes.byref(size), hdc_mem,
                                        ctypes.byref(src), 0, ctypes.byref(bf),
                                        2)      # ULW_ALPHA
        gdi32.DeleteObject(hbmp)
        gdi32.DeleteDC(hdc_mem)
        user32.ReleaseDC(0, hdc_screen)
        return bool(ok)
    except Exception:
        return False


def ulw_move(hwnd, x, y):
    """移动分层窗口（不改大小、不动 z 顺序）。"""
    if not HAVE_ULW or not hwnd:
        return
    try:
        ctypes.windll.user32.SetWindowPos(
            hwnd, 0, int(x), int(y), 0, 0,
            _SWP_NOSIZE | _SWP_NOZORDER | _SWP_NOACTIVATE)
    except Exception:
        pass


_FONT_DIR = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "Fonts")
_FONT_CACHE = {}


def _float_font(pt, bold=False, scale=1.0, ss=2):
    """按磅值取微软雅黑（ss 为超采样倍率；带缓存，避免每秒重复加载）。"""
    key = (round(pt, 1), bool(bold), round(scale, 3), int(ss))
    f = _FONT_CACHE.get(key)
    if f is not None:
        return f
    path = os.path.join(_FONT_DIR, "msyhbd.ttc" if bold else "msyh.ttc")
    px = max(6, int(round(pt * 96 / 72 * scale * ss)))
    try:
        f = ImageFont.truetype(path, px)
    except Exception:
        f = ImageFont.load_default()
    _FONT_CACHE[key] = f
    return f


def _compact_uptime(sec):
    """时长紧凑格式：3832 -> 1:03:52，225 -> 3:45。"""
    try:
        s = int(sec)
    except Exception:
        return "0:00"
    h, s = divmod(max(0, s), 3600)
    m, s = divmod(s, 60)
    return "%d:%02d:%02d" % (h, m, s) if h else "%d:%02d" % (m, s)


def _largest_pt(text, max_w, bold, scale, ss, base=13.0, floor=9.0):
    """返回能让 text 宽度不超过 max_w（逻辑像素）的最大字号（≤base）。"""
    limit = max_w * scale * ss
    pt = base
    while pt > floor:
        f = _float_font(pt, bold, scale, ss)
        try:
            if f.getlength(text) <= limit:
                break
        except Exception:
            break
        pt -= 0.5
    return pt


def render_float_card(snapshot, today, scale=1.0, alpha=204):
    """渲染悬浮窗卡片为 RGBA 图（2x 超采样后降采样，边缘平滑）。

    snapshot: Monitor.snapshot()；today: (rx, tx, ok) 本次总量（自监控启动以来的
    网卡累计收发，psutil 实时统计，准确）。
    alpha: 卡片底色不透明度 0~255（默认 204=80%；255 为不透明）。
    - 整卡铺一层带 alpha 的深色（>0 才能命中鼠标）
    - 所有文字带柔和暗影，浅色壁纸下也清晰
    - 本次总量只显示合计（不区分上传/下载）
    """
    ss = 2
    PW = int(round(FLOAT_W * scale))
    PH = int(round(FLOAT_H * scale))
    W, H = PW * ss, PH * ss
    k = scale * ss
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))

    # 卡片底 + 细边框（边框用暗色，避免浅色“白边”感）
    alpha = max(1, min(255, int(alpha)))
    tint = (13, 17, 23, alpha)
    edge = (70, 78, 92, 200 if alpha >= 255 else 130)
    ImageDraw.Draw(img).rounded_rectangle(
        [1 * k, 1 * k, W - 2 * k, H - 2 * k], radius=8 * k, fill=tint,
        outline=edge, width=max(1, int(round(k))))

    f_sm = _float_font(8, False, scale, ss)
    # 速度数字：以能同时放进左右半栏的最大字号渲染（长数值自动缩号）
    half = (FLOAT_W - 20) / 2 - 2
    rx_txt = snapshot.get("rx_rate_h", "0 B/s")
    tx_txt = snapshot.get("tx_rate_h", "0 B/s")
    f_lg = _float_font(min(_largest_pt(rx_txt, half, True, scale, ss),
                           _largest_pt(tx_txt, half, True, scale, ss)),
                       True, scale, ss)
    GRAY = (139, 148, 158, 255)
    DIM = (110, 118, 129, 255)
    BLUE = (88, 166, 255, 255)
    GREEN = (63, 185, 80, 255)
    PURPLE = (188, 140, 255, 255)
    AMBER = (210, 153, 34, 255)

    # 文本布局：[(mode, x, y, text, font, fill)]，mode: L 左对齐 / R 右对齐
    ops = [
        ("L", 24, 7, "USB 共享监控", f_sm, GRAY),
        ("R", FLOAT_W - 10, 7,
         "监测 " + _compact_uptime(snapshot.get("uptime", 0)), f_sm, DIM),
        ("L", 10, 26, "下载 ↓", f_sm, GRAY),
        ("L", 10, 38, rx_txt, f_lg, BLUE),
        ("R", FLOAT_W - 10, 26, "↑ 上传", f_sm, GRAY),
        ("R", FLOAT_W - 10, 38, tx_txt, f_lg, GREEN),
        ("L", 10, 64, "本次 ↓%s  ↑%s" % (snapshot.get("rx_total_h", "0 B"),
                                         snapshot.get("tx_total_h", "0 B")),
         f_sm, GRAY),
    ]
    trx, ttx, tok = today
    left_txt = ("本次总量 %s" % human_size(trx + ttx)) if tok else "本次总量 0 B"
    # 预计 GB/H：本次总量 ÷ 监测时长 × 3600s ÷ 1GB（右对齐放在本次总量右侧）
    right_txt = "预计 %.2f GB/H" % snapshot.get("gph", 0.0)
    # 底部一行左右分列：字号自适应收缩，保证总量数值再长也不会与预计重叠
    gap = 6.0
    avail = (FLOAT_W - 20 - gap) * scale * ss
    bot_pt = 8.0
    while bot_pt > 6.0:
        f_try = _float_font(bot_pt, False, scale, ss)
        if (f_try.getlength(left_txt) + f_try.getlength(right_txt)) <= avail:
            break
        bot_pt -= 0.5
    f_bot = _float_font(bot_pt, False, scale, ss)
    ops.append(("L", 10, 78, left_txt, f_bot, PURPLE if tok else DIM))
    ops.append(("R", FLOAT_W - 10, 78, right_txt, f_bot, AMBER))

    def draw_ops(layer, fill=None):
        for mode, x, y, text, font, f in ops:
            px = x * k if mode == "L" else x * k - layer.textlength(text,
                                                                    font=font)
            layer.text((px, y * k), text, font=font, fill=fill or f)

    # 状态点
    dr = ImageDraw.Draw(img)
    cx, cy, r = 15 * k, 14.5 * k, 3 * k
    dr.ellipse([cx - r, cy - r, cx + r, cy + r],
               fill=GREEN if snapshot.get("connected") else AMBER)

    # 文字暗影（模糊后垫在文字下）
    shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw_ops(ImageDraw.Draw(shadow), (0, 0, 0, 200))
    img.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(1.7 * k)))
    draw_ops(ImageDraw.Draw(img))

    return img.resize((PW, PH), Image.LANCZOS)


def detect_usb_tether_iface():
    """通过 WMI 找到 USB 网络共享网卡 (描述含 RNDIS / USB 的网卡)。
    不依赖 powershell，纯 Python 实现。"""
    try:
        import win32com.client
        wmi = win32com.client.GetObject("winmgmts:")
        nics = wmi.ExecQuery(
            "SELECT Name, Description, NetConnectionID, NetConnectionStatus "
            "FROM Win32_NetworkAdapter"
        )
        matches = []
        for n in nics:
            desc = (n.Description or "").lower()
            name = n.NetConnectionID or n.Name
            if ("rndis" in desc) or ("usb" in desc and "ethernet" in desc):
                connected = (n.NetConnectionStatus == 2)
                matches.append((name, n.Description, connected))
        if not matches:
            return None, None
        up = [m for m in matches if m[2]]
        pick = (up or matches)[0]
        return pick[0], pick[1]
    except Exception as e:
        print(f"[WARN] WMI 检测失败: {e}", flush=True)
        return None, None


def human_rate(bps):
    """字节/秒 -> 友好字符串。"""
    if bps < 1024:
        return f"{bps:.0f} B/s"
    elif bps < 1024 ** 2:
        return f"{bps / 1024:.1f} KB/s"
    elif bps < 1024 ** 3:
        return f"{bps / 1024 ** 2:.2f} MB/s"
    else:
        return f"{bps / 1024 ** 3:.2f} GB/s"


def human_size(byte):
    if byte < 1024:
        return f"{byte:.0f} B"
    elif byte < 1024 ** 2:
        return f"{byte / 1024:.1f} KB"
    elif byte < 1024 ** 3:
        return f"{byte / 1024 ** 2:.2f} MB"
    else:
        return f"{byte / 1024 ** 3:.2f} GB"


class Monitor:
    def __init__(self, iface=None):
        self.iface = iface
        self.desc = ""
        if iface is None:
            self.iface, self.desc = detect_usb_tether_iface()
        self.lock = threading.Lock()
        self.rx_total = 0.0   # 本次会话累计接收
        self.tx_total = 0.0   # 本次会话累计发送
        self.rx_rate = 0.0
        self.tx_rate = 0.0
        self.connected = False
        self.start_time = time.time()  # 监测启动时刻
        self._prev = None
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._sample, daemon=True)
        self._thread.start()

    def _read(self):
        if not self.iface:
            self.iface, self.desc = detect_usb_tether_iface()
        if not self.iface:
            return None
        io = psutil.net_io_counters(pernic=True)
        c = io.get(self.iface)
        if c is None:
            # 网卡名可能变化，尝试重新检测
            self.iface, self.desc = detect_usb_tether_iface()
            c = io.get(self.iface) if self.iface else None
        return c

    def _sample(self):
        while not self._stop.is_set():
            c = self._read()
            now = time.time()
            if c is not None:
                cur = (c.bytes_recv, c.bytes_sent)
                self.connected = True
                if self._prev is not None:
                    dt = now - self._prev[2]
                    if dt > 0:
                        dr = max(0.0, (cur[0] - self._prev[0]) / dt)
                        tr = max(0.0, (cur[1] - self._prev[1]) / dt)
                        self.rx_total += dr * dt
                        self.tx_total += tr * dt
                        with self.lock:
                            self.rx_rate = dr
                            self.tx_rate = tr
                self._prev = (cur[0], cur[1], now)
            else:
                self.connected = False
                self._prev = None
            time.sleep(1.0)

    def snapshot(self):
        with self.lock:
            return {
                "iface": self.iface,
                "desc": self.desc,
                "connected": self.connected,
                "rx_rate": self.rx_rate,
                "tx_rate": self.tx_rate,
                "rx_total": self.rx_total,
                "tx_total": self.tx_total,
                "rx_rate_h": human_rate(self.rx_rate),
                "tx_rate_h": human_rate(self.tx_rate),
                "rx_total_h": human_size(self.rx_total),
                "tx_total_h": human_size(self.tx_total),
                "uptime": (up := (time.time() - self.start_time)),
                # 预计 GB/H：本次总量 ÷ 监测时长 × 3600s ÷ 1GB
                "gph": ((self.rx_total + self.tx_total) / up * 3600 / (1024 ** 3))
                       if up > 1e-6 else 0.0,
                "ts": time.time(),
            }

    def reset(self):
        """重新监控：清零本次会话累计收发与速率，重置监测起始时刻与差分基准。"""
        with self.lock:
            self.rx_total = 0.0
            self.tx_total = 0.0
            self.rx_rate = 0.0
            self.tx_rate = 0.0
        self.start_time = time.time()
        self._prev = None   # 下一采样重新建立基准，避免清零后瞬间跳变

MONITOR = None


def run_float():
    """运行桌面悬浮窗：常驻、置顶、可拖动，实时显示速度与本次用量。
    渲染采用 UpdateLayeredWindow 逐像素 Alpha：
    - 整卡 alpha>0（视觉近透明），鼠标可全程拖动，不会点击穿透；
    - 文字带柔和暗影，浅色壁纸下也清晰。
    纯本地模式：不启动 Web 服务、不占用网络端口。"""
    get_stats = lambda: MONITOR.snapshot()

    if not (HAVE_ULW and HAVE_PIL):
        print("[ERR] 悬浮窗需要 Pillow（pip install Pillow）。", flush=True)
        return

    # 单实例锁：避免重复启动产生多个悬浮窗（占用锁端口 8779）
    import socket as _sock
    _lock = _sock.socket(_sock.AF_INET, _sock.SOCK_STREAM)
    try:
        _lock.bind(("127.0.0.1", 8779))
    except OSError:
        # 已有悬浮窗实例在运行，本次不重复启动
        print("[WARN] 已有悬浮窗在运行，本次不重复启动。", flush=True)
        sys.exit(0)

    root = tk.Tk()
    root.overrideredirect(True)          # 无标题栏
    root.attributes("-topmost", True)    # 置顶
    root.configure(bg="#000000")
    try:
        root.wm_attributes("-toolwindow", True)
    except Exception:
        pass

    # 尺寸（按系统 DPI 缩放，高 DPI 屏上不小）
    try:
        dpi = ctypes.windll.user32.GetDpiForSystem() or 96
    except Exception:
        dpi = 96
    scale = max(1.0, dpi / 96.0)
    PW = int(round(FLOAT_W * scale))
    PH = int(round(FLOAT_H * scale))
    root.geometry(f"{PW}x{PH}")
    root.update_idletasks()
    root.geometry(f"+{root.winfo_screenwidth() - PW - 20}+40")  # 右上角
    root.update()

    hwnd = ulw_toplevel(root)
    if not ulw_enable(hwnd):
        print("[ERR] 无法启用分层窗口，悬浮窗不可用。", flush=True)
        return

    pos = {"x": 0, "y": 0}
    rect = ulw_rect(hwnd)
    if rect:
        pos["x"], pos["y"] = rect[0], rect[1]
    last_img = {"img": None}

    def paint():
        """渲染一张卡片并按当前位置贴到分层窗口。"""
        try:
            d = get_stats()
            # 本次总量：自监控启动以来网卡的累计收发（psutil 实时统计，准确）
            today = (d["rx_total"], d["tx_total"], True)
            img = render_float_card(d, today, scale)   # 固定 80% 淡底
            last_img["img"] = img
            ulw_paint(hwnd, img, pos["x"], pos["y"])
        except Exception:
            pass

    # 拖动（整卡 alpha>0，任意位置都可抓取）
    def start_drag(e):
        root._dx, root._dy = e.x, e.y

    def on_drag(e):
        if root._dx is None:
            return
        nx = pos["x"] + e.x - root._dx
        ny = pos["y"] + e.y - root._dy
        pos["x"], pos["y"] = int(nx), int(ny)
        ulw_move(hwnd, nx, ny)
        if last_img["img"] is not None:   # 用缓存图重新贴到新位置
            ulw_paint(hwnd, last_img["img"], nx, ny)

    root._dx = root._dy = None
    root.bind("<Button-1>", start_drag)
    root.bind("<B1-Motion>", on_drag)

    # 右键菜单（置顶切换、重新监控、退出）
    menu = tk.Menu(root, tearoff=0, bg="#161b22", fg="#c9d1d9",
                  activebackground="#1f6feb", activeforeground="#ffffff")
    menu.add_command(label="置顶 / 置底", command=lambda: root.attributes(
        "-topmost", not root.attributes("-topmost")))
    menu.add_command(label="重新监控", command=lambda: MONITOR.reset())
    menu.add_command(label="退出", command=lambda: (root.destroy(), sys.exit(0)))
    root.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

    def update():
        paint()
        root.after(1000, update)

    root.after(200, update)
    root.mainloop()  # 阻塞直到窗口关闭


def main():
    global MONITOR
    p = argparse.ArgumentParser(description="USB 网络共享流量悬浮窗")
    p.add_argument("--iface", help="指定网卡名称，如 '以太网 2'")
    args = p.parse_args()

    # 唯一运行模式：桌面悬浮窗
    if not (HAVE_GUI and HAVE_ULW and HAVE_PIL):
        print("[ERR] 悬浮窗需要 tkinter + Pillow（系统 Python 3.12 + pip install Pillow）。",
              flush=True)
        return
    MONITOR = Monitor(iface=args.iface)
    iface, desc = MONITOR.iface, MONITOR.desc
    if iface:
        print(f"[OK] 已识别 USB 网络共享网卡: {iface}  ({desc})", flush=True)
    else:
        print("[WARN] 未检测到 USB 网络共享网卡（开启手机 USB 网络共享后将自动识别）",
              flush=True)
    run_float()


if __name__ == "__main__":
    main()
