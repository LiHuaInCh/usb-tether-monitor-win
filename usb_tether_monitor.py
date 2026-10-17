#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
USB 网络共享 / 投屏流量监控（桌面悬浮窗）
- 自动识别 Windows 下 USB 网络共享网卡 (描述含 RNDIS / USB 的网卡)
- 桌面悬浮窗实时显示下载/上传速度、本次累计、今日总用量（读系统「数据使用量」）与监测时长
- 悬浮窗为 UpdateLayeredWindow 逐像素 Alpha 渲染：80% 深色淡底可全程拖动，文字带暗影
- 历史窗口 (右键悬浮窗「查看历史 / 详情」或 --history)：
  · 各软件每日明细矩阵（读取 Windows 自带「数据使用量」记录，按当前网卡归档，首行总用量）
用法:
    python usb_tether_monitor.py                # 桌面悬浮窗（默认模式）
    python usb_tether_monitor.py --history      # 独立打开历史用量窗口
    python usb_tether_monitor.py --iface "以太网 2"   # 指定网卡
    python usb_tether_monitor.py --install-startup   # 创建开机自启快捷方式
    python usb_tether_monitor.py --install-shortcut  # 创建桌面快捷方式
    python usb_tether_monitor.py --uninstall-startup # 删除开机自启快捷方式
依赖: psutil, pywin32, Pillow (悬浮窗用系统 Python3.12 自带 tkinter)；
      按软件用量读取需 winrt-Windows.Networking.Connectivity / .Foundation[.Collections]
"""
import argparse
import ctypes
import datetime
import os
import sys
import threading
import time
import webbrowser

import psutil

# 可选依赖：开机自启（pywin32）；悬浮窗渲染（Pillow）
try:
    import win32com.client
    from win32com.shell import shell, shellcon
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

try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter
    HAVE_PIL = True
except Exception:
    HAVE_PIL = False

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

    snapshot: Monitor.snapshot()；today: (rx, tx, ok) 今日总用量（混合口径：
    系统「数据使用量」在参考时刻的快照 + psutil 实时增量，实时准确）。
    alpha: 卡片底色不透明度 0~255（默认 204=80%；255 为不透明）。
    - 整卡铺一层带 alpha 的深色（>0 才能命中鼠标）
    - 所有文字带柔和暗影，浅色壁纸下也清晰
    - 今日总用量只显示合计（不区分上传/下载）
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
    if tok:
        ops.append(("L", 10, 78, "今日 %s" % human_size(trx + ttx),
                    f_sm, PURPLE))
    else:
        ops.append(("L", 10, 78, "今日 暂无记录", f_sm, DIM))

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

APP_USAGE_DAYS = 7  # 各软件每日明细展示最近几天

# 系统「数据使用量」读取（按软件归因的用量，即设置里那本账）
try:
    from winrt.windows.networking.connectivity import (
        NetworkInformation as _NetworkInformation,
        NetworkUsageStates as _NetworkUsageStates,
        TriStates as _TriStates)
    HAVE_APPUSAGE = True
except Exception:
    HAVE_APPUSAGE = False


def _find_connection_profile(iface_name):
    """按网卡名找到系统记录用量的 ConnectionProfile（按适配器归档）。"""
    if not HAVE_APPUSAGE or not iface_name:
        return None
    try:
        profiles = list(_NetworkInformation.get_connection_profiles())
    except Exception:
        return None
    for p in profiles:
        if (p.profile_name or "") == iface_name:
            return p
    cands = [p for p in profiles
             if iface_name in (p.profile_name or "")]
    if cands:
        def _lvl(p):
            try:
                return 0 if p.get_network_connectivity_level() else 1
            except Exception:
                return 1
        cands.sort(key=_lvl)
        return cands[0]
    return None


def _display_name(name, aid):
    """把系统归因信息整理成可读的软件名。"""
    n = (name or "").strip()
    a = (aid or "").strip()
    if n.lower() == "system" or not a or a.lower().startswith("system"):
        return "系统"
    if n:
        return n
    if "\\" in a or "/" in a:  # 进程路径 -> 取 exe 名（去扩展名）
        base = a.replace("/", "\\").split("\\")[-1]
        if base.lower().endswith(".exe"):
            base = base[:-4]
        return base or "未知"
    # 包家族名，如 MicrosoftWindows.Client.WebExperience_cw5n1h2txyewy
    return a.split(".")[-1].split("_")[0] or a


def _query_app_usage(profile, start, end):
    """查询 [start, end] 内各软件收发字节 -> {名称: [rx, tx]}。
    注意：WinRT 返回的序列不可重复迭代，这里一次性遍历并固化。"""
    states = _NetworkUsageStates(roaming=_TriStates.DO_NOT_CARE,
                                 shared=_TriStates.DO_NOT_CARE)
    op = profile.get_attributed_network_usage_async(start, end, states)
    agg = {}
    for u in op.get():
        key = _display_name(u.attribution_name, u.attribution_id)
        cur = agg.setdefault(key, [0, 0])
        cur[0] += int(u.bytes_received)
        cur[1] += int(u.bytes_sent)
    return agg


# 今日用量基线状态：记录「今日 00:00」的网卡计数，跨监控重启/零点翻篇复用。
_TODAY_STATE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "usb_tether_today.json")


def _load_today_state():
    try:
        import json
        with open(_TODAY_STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _save_today_state(state):
    try:
        import json
        with open(_TODAY_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f)
    except Exception:
        pass


def app_usage_total(iface_name, days=30):
    """各软件最近 N 天累计用量 {名称: [rx, tx]}（含今天，起自相应日期 00:00）；
    找不到网卡记录返回 None。"""
    if not HAVE_APPUSAGE or not iface_name:
        return None
    profile = _find_connection_profile(iface_name)
    if profile is None:
        return None
    now = datetime.datetime.now()
    today0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
    start = today0 - datetime.timedelta(days=max(1, days) - 1)
    return _query_app_usage(profile, start, now)


def app_usage_by_day(iface_name, days=7):
    """各软件最近 N 天每日明细 ({名称: {日期: [rx, tx]}}, [日期...])；
    找不到网卡记录返回 (None, [])。"""
    if not HAVE_APPUSAGE or not iface_name:
        return None, []
    profile = _find_connection_profile(iface_name)
    if profile is None:
        return None, []
    now = datetime.datetime.now()
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    data = {}
    dates = []
    for i in range(days - 1, -1, -1):
        d0 = today - datetime.timedelta(days=i)
        label = d0.strftime("%m-%d")
        dates.append(label)
        end = now if i == 0 else d0 + datetime.timedelta(days=1)
        for name, (rx, tx) in _query_app_usage(profile, d0, end).items():
            data.setdefault(name, {})[label] = [rx, tx]
    return data, dates


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
                "uptime": time.time() - self.start_time,
                "ts": time.time(),
            }

MONITOR = None


def find_pythonw():
    """返回用于快捷方式的 pythonw（无控制台窗口）。

    优先用环境变量 USB_TETHER_PYTHONW 指定（比如自带 tkinter 的系统
    Python），否则退回当前解释器同目录的 pythonw，再否则用当前解释器。
    """
    candidates = [
        os.environ.get("USB_TETHER_PYTHONW"),
        os.path.join(os.path.dirname(sys.executable), "pythonw.exe"),
        sys.executable,
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return sys.executable


def make_shortcut(folder, name, description):
    """在指定文件夹创建指向 pythonw + 脚本（默认悬浮窗模式）的快捷方式。"""
    if not HAVE_WIN32:
        print("[ERR] 缺少 pywin32，无法创建快捷方式。")
        return None
    link = os.path.join(folder, name)
    script = os.path.abspath(__file__)
    pythonw = find_pythonw()
    ws = win32com.client.Dispatch("WScript.Shell")
    sc = ws.CreateShortcut(link)
    sc.TargetPath = pythonw
    sc.Arguments = f'"{script}"'
    sc.WorkingDirectory = os.path.dirname(script)
    sc.WindowStyle = 7  # 最小化运行，不弹窗
    sc.Description = description
    sc.Save()
    return link


def install_startup():
    """在用户「启动」文件夹创建快捷方式，实现开机自启（无黑窗口）。"""
    if not HAVE_WIN32:
        print("[ERR] 缺少 pywin32，无法创建开机自启快捷方式。")
        return False
    startup = shell.SHGetFolderPath(0, shellcon.CSIDL_STARTUP, None, 0)
    link = make_shortcut(startup, "USB_Tether_Monitor.lnk",
                         "USB 网络共享实时流量监控")
    if link:
        print(f"[OK] 已创建开机自启: {link}")
    return bool(link)


def install_shortcut():
    """在用户桌面创建快捷方式，方便手动双击启动悬浮窗。"""
    if not HAVE_WIN32:
        print("[ERR] 缺少 pywin32，无法创建桌面快捷方式。")
        return
    desktop = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOPDIRECTORY, None, 0)
    link = make_shortcut(desktop, "USB流量悬浮窗.lnk",
                         "USB 网络共享实时流量悬浮窗")
    if link:
        print(f"[OK] 已创建桌面快捷方式: {link}")


def uninstall_startup():
    """删除开机自启快捷方式。"""
    if not HAVE_WIN32:
        print("[ERR] 缺少 pywin32，无法卸载开机自启。")
        return
    startup = shell.SHGetFolderPath(0, shellcon.CSIDL_STARTUP, None, 0)
    link = os.path.join(startup, "USB_Tether_Monitor.lnk")
    if os.path.exists(link):
        os.remove(link)
        print(f"[OK] 已删除开机自启: {link}")
    else:
        print("[INFO] 未发现开机自启快捷方式。")


def run_float():
    """运行桌面悬浮窗：常驻、置顶、可拖动，实时显示速度。
    渲染采用 UpdateLayeredWindow 逐像素 Alpha：
    - 整卡 alpha>0（视觉近透明），鼠标可全程拖动，不会点击穿透；
    - 文字带柔和暗影，浅色壁纸下也清晰。
    纯本地模式：不启动 Web 服务、不占用网络端口。"""
    import tkinter as tk
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

    # 今日总用量：直接基于 psutil 网卡计数器计算「自 00:00 以来的真实流量」，
    #   实时、准确。系统「数据使用量」会延迟几十秒批量入账、短时窗漏记，且实测
    #   比真实流量少约 20%（还会把上传方向记反/多记），单用它做今日总量会偏低且滞后。
    #   本机每天 08:20 左右开机 → 计数器在当日 00:00 时为 0 → 今日总量 = 当前计数 - 0。
    #   跨零点 / 中途重启：用持久化基线（必要时以系统记录反推）保证翻篇正确。
    today_usage = {"mode": "system", "base_rx": 0.0, "base_tx": 0.0, "ok": False}
    _usage_lock = threading.Lock()
    tick = {"n": 0}

    def _psutil_now():
        """当前网卡收发字节（psutil，实时、准确）；取不到返回 None。"""
        try:
            io = psutil.net_io_counters(pernic=True).get(
                getattr(MONITOR, "iface", None))
            return (io.bytes_recv, io.bytes_sent) if io else None
        except Exception:
            return None

    def _establish_baseline(system_rx, system_tx):
        """确定「今日 00:00」的网卡计数基线 (rx, tx)，写入持久化状态并返回。"""
        today_str = datetime.date.today().isoformat()
        cur = _psutil_now()
        if cur is None:
            return None
        state = _load_today_state()
        same_day = state and state.get("date") == today_str
        rebooted = bool(same_day and (cur[0] < state.get("last_rx", 0)
                                      or cur[1] < state.get("last_tx", 0)))
        # 同一天且未重启（当前计数 >= 上次）：沿用已持久化的基线（监控重启不丢）
        if same_day and not rebooted:
            base = (state["baseline_rx"], state["baseline_tx"])
            _save_today_state({"date": today_str,
                               "baseline_rx": base[0], "baseline_tx": base[1],
                               "last_rx": cur[0], "last_tx": cur[1]})
            return base
        # 新的一天，或中途重启：重建基线
        booted_today = (datetime.date.fromtimestamp(psutil.boot_time())
                       == datetime.date.today())
        if booted_today and not rebooted:
            base = (0.0, 0.0)            # 计数器在当日 00:00 时为 0
        else:
            # 开机早于今天，或中途重启：用系统记录反推 00:00 基线（最佳近似）
            base = (max(0.0, cur[0] - system_rx),
                    max(0.0, cur[1] - system_tx))
        _save_today_state({"date": today_str,
                           "baseline_rx": base[0], "baseline_tx": base[1],
                           "last_rx": cur[0], "last_tx": cur[1]})
        return base

    def refresh_today_usage():
        def work():
            try:
                import pythoncom
                pythoncom.CoInitializeEx(0)  # MTA，供 WinRT 阻塞调用
            except Exception:
                pass
            data = None
            try:
                data = app_usage_total(getattr(MONITOR, "iface", None), 1)
            except Exception:
                data = None
            srx = sum(v[0] for v in data.values()) if data else 0.0
            stx = sum(v[1] for v in data.values()) if data else 0.0
            base = _establish_baseline(srx, stx)
            cur = _psutil_now()
            with _usage_lock:
                if base is not None and cur is not None:
                    # 正常：用 psutil 计数器 - 00:00 基线（实时、准确）
                    today_usage.update(mode="psutil", base_rx=base[0],
                                       base_tx=base[1], ok=True)
                else:
                    # 退化：psutil 取不到时直接用系统记录（可能滞后/偏低）
                    today_usage.update(mode="system", base_rx=srx,
                                       base_tx=stx, ok=bool(data))
        threading.Thread(target=work, daemon=True).start()

    def paint():
        """渲染一张卡片并按当前位置贴到分层窗口。"""
        try:
            d = get_stats()
            cur = _psutil_now()
            with _usage_lock:
                tu = dict(today_usage)
            if tu["mode"] == "psutil" and cur is not None:
                # 用最新实时计数 - 00:00 基线，消除 30s 刷新间隔带来的滞后
                rx = max(0.0, cur[0] - tu["base_rx"])
                tx = max(0.0, cur[1] - tu["base_tx"])
                today = (rx, tx, True)
            else:
                today = (tu["base_rx"], tu["base_tx"], tu["ok"])
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

    # 右键菜单
    menu = tk.Menu(root, tearoff=0, bg="#161b22", fg="#c9d1d9",
                  activebackground="#1f6feb", activeforeground="#ffffff")
    menu.add_command(label="查看历史 / 详情",
                     command=lambda: build_history_window(
                         root, getattr(MONITOR, "iface", None)))
    menu.add_command(label="置顶 / 置底", command=lambda: root.attributes(
        "-topmost", not root.attributes("-topmost")))
    menu.add_command(label="退出", command=lambda: (root.destroy(), sys.exit(0)))
    root.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))

    def update():
        paint()
        tick["n"] += 1
        if tick["n"] % 30 == 5:  # 约每 30 秒刷新一次系统今日总用量
            refresh_today_usage()
        root.after(1000, update)

    root.after(200, update)
    root.after(1500, refresh_today_usage)  # 首次加载
    root.mainloop()  # 阻塞直到窗口关闭


def build_history_window(parent=None, iface=None):
    """历史用量窗口（深色主题，与悬浮窗同配色）：
    各软件每日明细矩阵（读取 Windows 自带「数据使用量」记录，按当前网卡归档）。
    parent 为 None 时独立运行（自带 mainloop）；否则作为子窗口挂在 parent 上。"""
    import tkinter as tk
    from tkinter import ttk
    import queue

    # ---- 深色主题（GitHub Dark 配色，与悬浮窗一致） ----
    BG, PANEL, BORDER = "#0d1117", "#161b22", "#30363d"
    FG, MUTED, TITLE = "#c9d1d9", "#8b949e", "#e6edf3"
    PURPLE = "#bc8cff"
    HEAD_BG, SEL_BG, HOVER = "#21262d", "#1f6feb", "#2d333b"
    MW = ("Microsoft YaHei",)

    def apply_style(root):
        st = ttk.Style(root)
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("Treeview", background=BG, foreground=FG,
                     fieldbackground=BG, borderwidth=0, rowheight=28,
                     font=MW + (9,))
        st.configure("Treeview.Heading", background=HEAD_BG, foreground=MUTED,
                     borderwidth=0, relief="flat", font=MW + (9, "bold"),
                     padding=(6, 7))
        st.map("Treeview.Heading", background=[("active", HEAD_BG)])
        st.map("Treeview", background=[("selected", SEL_BG)],
               foreground=[("selected", "#ffffff")])
        st.configure("TButton", background=HEAD_BG, foreground=FG,
                     bordercolor=BORDER, borderwidth=1, focuscolor=BG,
                     padding=(8, 2), font=MW + (8,))
        st.map("TButton", background=[("active", HOVER),
                                      ("pressed", "#1c2128")])
        for orient in ("Vertical", "Horizontal"):
            st.configure(orient + ".TScrollbar", background="#3d444d",
                         troughcolor=BG, bordercolor=BG, arrowcolor=MUTED,
                         relief="flat")
            st.map(orient + ".TScrollbar",
                   background=[("active", "#484f58"), ("pressed", "#565d66")])

    def card(p):
        return tk.Frame(p, bg=PANEL, highlightbackground=BORDER,
                        highlightthickness=1, bd=0)

    def treebox(c, height):
        """带横竖滚动条的深色表格，返回 (tree, box)。"""
        box = tk.Frame(c, bg=PANEL)
        box.pack(fill="both", expand=True, padx=12, pady=(2, 10))
        t = ttk.Treeview(box, show="headings", height=height)
        t.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(box, orient="vertical", command=t.yview)
        vsb.pack(side="right", fill="y")
        t.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(box, orient="horizontal", command=t.xview)
        hsb.pack(side="bottom", fill="x")
        t.configure(xscrollcommand=hsb.set)
        return t, box

    def open_sys_settings():
        try:
            os.startfile("ms-settings:data-usage")
        except Exception:
            try:
                webbrowser.open("ms-settings:data-usage")
            except Exception:
                pass

    if iface is None:
        try:
            iface = detect_usb_tether_iface()[0]
        except Exception:
            iface = None

    own_root = parent is None
    win = tk.Tk() if own_root else tk.Toplevel(parent)
    win.title("USB 网络共享 · 各软件每日明细")
    win.configure(bg=BG)
    win.geometry("860x600")
    try:
        win.minsize(800, 520)
    except Exception:
        pass
    apply_style(win)

    # ---- 顶部：标题 + 网卡 ----
    head = card(win)
    head.pack(fill="x", padx=14, pady=(14, 10))
    hr = tk.Frame(head, bg=PANEL)
    hr.pack(fill="x", padx=14, pady=(12, 12))
    tk.Label(hr, text="USB 网络共享 · 历史用量", bg=PANEL, fg=TITLE,
             font=MW + (14, "bold")).pack(side="left")
    tk.Label(hr, text="网卡：%s" % (iface or "未检测到"), bg=PANEL, fg=MUTED,
             font=MW + (8,)).pack(side="right", pady=(6, 0))

    # ---- 各软件每日明细（可折叠分区） ----
    c1 = card(win)
    c1.pack(fill="both", expand=True, padx=14, pady=(0, 10))
    h = tk.Frame(c1, bg=PANEL)
    h.pack(fill="x", padx=12, pady=(10, 2))
    accent = tk.Frame(h, width=3, height=14, bg=PURPLE)
    accent.pack(side="left", padx=(0, 7))
    arrow = tk.Label(h, text="▾", bg=PANEL, fg=MUTED, font=MW + (9,),
                     cursor="hand2")
    arrow.pack(side="left")
    ttl = tk.Label(h, text="各软件每日明细", bg=PANEL, fg=TITLE,
                   font=MW + (10, "bold"), cursor="hand2")
    ttl.pack(side="left", padx=(4, 0))
    hint = tk.Label(h, text="最近 %d 天 · 每列为当天上下行合计 · 首行为总用量"
                    % APP_USAGE_DAYS, bg=PANEL, fg=MUTED, font=MW + (8,),
                    cursor="hand2")
    hint.pack(side="left", padx=(8, 0))
    ttk.Button(h, text="刷新", width=6,
               command=lambda: load_app_data()).pack(side="right")
    ttk.Button(h, text="系统设置页", width=11,
               command=open_sys_settings).pack(side="right", padx=(0, 8))
    app_status = tk.Label(h, text="", bg=PANEL, fg=MUTED, font=MW + (8,))
    app_status.pack(side="right", padx=(0, 12))

    day_tree, box = treebox(c1, 12)
    day_tree.tag_configure("total", foreground=TITLE, font=MW + (9, "bold"))

    # 折叠：点头部（标题/提示/accent 条）收起或展开表格
    def toggle(e=None):
        if box.winfo_ismapped():
            box.pack_forget()
            arrow.config(text="▸")
        else:
            box.pack(fill="both", expand=True, padx=12, pady=(2, 10))
            arrow.config(text="▾")

    for w in (h, accent, arrow, ttl, hint):
        w.bind("<Button-1>", toggle)

    tk.Label(win, text="数据来自 Windows 自带「数据使用量」记录"
                       "（设置 → 网络和 Internet → 高级网络设置 → 数据使用量），"
                       "按当前网卡归档，含投屏等本地传输；无记录的日期显示 “—”。",
             bg=BG, fg=MUTED, font=MW + (8,), justify="left",
             wraplength=800).pack(anchor="w", padx=16, pady=(0, 12))

    result_q = queue.Queue()
    state = {"loading": False}

    def _fill_day_tree(matrix, dates):
        cols = ("name",) + tuple(dates) + ("total",)
        day_tree["columns"] = cols
        for c in cols:
            text = "软件" if c == "name" else c
            if dates and c == dates[-1]:
                text += " ·今天"
            day_tree.heading(c, text=text)
            day_tree.column(c, width=200 if c == "name" else
                            (84 if c != "total" else 100),
                            minwidth=60,
                            anchor="w" if c == "name" else "e")
        for it in day_tree.get_children():
            day_tree.delete(it)
        if matrix is None:
            day_tree.insert("", "end", values=("（未找到该网卡的系统记录）",))
            return
        rows = []
        for name, per in matrix.items():
            tot = sum(v[0] + v[1] for v in per.values())
            rows.append((name, per, tot))
        rows.sort(key=lambda x: x[2], reverse=True)
        shown = [r for r in rows if r[2] > 0][:15]
        if not shown:
            day_tree.insert("", "end", values=("（该时段无记录）",))
            return
        # 总用量行（置顶，始终可见）：全部软件按天汇总
        vals = ["总用量"]
        for d in dates:
            s = sum(r[1].get(d, [0, 0])[0] + r[1].get(d, [0, 0])[1]
                    for r in rows)
            vals.append(human_size(s) if s else "—")
        vals.append(human_size(sum(r[2] for r in rows)))
        day_tree.insert("", "end", values=tuple(vals), tags=("total",))
        for name, per, tot in shown:
            vals = [name]
            for d in dates:
                v = per.get(d)
                vals.append(human_size(v[0] + v[1]) if v else "—")
            vals.append(human_size(tot))
            day_tree.insert("", "end", values=tuple(vals))
        rest = [r for r in rows if r[2] > 0][15:]
        if rest:
            vals = ["其他 %d 个软件" % len(rest)]
            for d in dates:
                s = sum(r[1].get(d, [0, 0])[0] + r[1].get(d, [0, 0])[1]
                        for r in rest)
                vals.append(human_size(s) if s else "—")
            vals.append(human_size(sum(r[2] for r in rest)))
            day_tree.insert("", "end", values=tuple(vals))

    def load_app_data():
        if state["loading"]:
            return
        if not HAVE_APPUSAGE:
            app_status.config(text="未安装 winrt 组件，无法读取系统记录")
            _fill_day_tree(None, [])
            return
        state["loading"] = True
        app_status.config(text="加载中…")

        def work():
            # WinRT 阻塞调用不能在 STA 线程中执行；显式以 MTA 初始化本线程
            try:
                import pythoncom
                pythoncom.CoInitializeEx(0)  # COINIT_MULTITHREADED
            except Exception:
                pass
            matrix = None
            dates = []
            err = ""
            try:
                matrix, dates = app_usage_by_day(iface, APP_USAGE_DAYS)
            except Exception as e:
                err = str(e)
            result_q.put((err, matrix, dates))

        threading.Thread(target=work, daemon=True).start()

    def drain_queue():
        try:
            while True:
                err, matrix, dates = result_q.get_nowait()
                state["loading"] = False
                if err:
                    app_status.config(text="读取失败：" + err[:40])
                elif matrix is None:
                    app_status.config(text="未找到该网卡的系统记录"
                                       "（需已连接 USB 网络共享）")
                elif not matrix:
                    app_status.config(text="该时段无记录")
                else:
                    app_status.config(text="来源：系统「数据使用量」")
                _fill_day_tree(matrix, dates)
        except queue.Empty:
            pass

    def poll_queue():  # 每 250ms 取一次后台结果；每 20s 自动重载
        drain_queue()
        poll_queue.n += 1
        if poll_queue.n % 80 == 3:
            load_app_data()
        win.after(250, poll_queue)

    poll_queue.n = 0
    win.after(300, load_app_data)
    win.after(250, poll_queue)
    if own_root:
        win.mainloop()
    return win


def main():
    global MONITOR
    p = argparse.ArgumentParser(description="USB 网络共享流量悬浮窗")
    p.add_argument("--iface", help="指定网卡名称，如 '以太网 2'")
    p.add_argument("--float", action="store_true",
                   help=argparse.SUPPRESS)  # 兼容旧快捷方式（悬浮窗已是默认模式）
    p.add_argument("--history", action="store_true", help="打开历史用量窗口")
    p.add_argument("--install-startup", action="store_true", help="创建开机自启快捷方式")
    p.add_argument("--install-shortcut", action="store_true", help="创建桌面快捷方式")
    p.add_argument("--uninstall-startup", action="store_true", help="删除开机自启快捷方式")
    args = p.parse_args()

    if args.install_startup:
        install_startup()
        return
    if args.install_shortcut:
        install_shortcut()
        return
    if args.uninstall_startup:
        uninstall_startup()
        return

    # 独立历史用量窗口（每日总量 + 各软件每日明细）
    if args.history:
        if not HAVE_GUI:
            print("[ERR] 当前环境不支持 GUI（缺少 tkinter）。", flush=True)
            return
        build_history_window(None)
        return

    # 默认（也是唯一）运行模式：桌面悬浮窗
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
