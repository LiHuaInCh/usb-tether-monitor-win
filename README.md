# USB 网络共享流量监控（Windows 桌面悬浮窗）

一个常驻桌面的小悬浮窗：实时显示 USB 网络共享（手机共享网络 / USB 绑网）的
**下载 / 上传速度**、**本次累计**、**今日总用量**和**监测时长**；并带一个历史面板，
按天展示**每个软件**的用量明细（直接读取 Windows 自带的「数据使用量」记录）。

单文件、纯本地、无联网上报、不占用 Web 端口。

![深色壁纸](docs/preview-dark.png)
![浅色壁纸](docs/preview-light.png)

## 功能

- **自动识别网卡**：按网卡描述（RNDIS / USB Ethernet）找到 USB 网络共享接口，手机插拔后自动重新识别
- **实时速度**：下载 / 上传速度每秒刷新，数值过长时自动缩小字号防止两列重叠
- **今日总用量**：读取 Windows 自带「数据使用量」记录（设置 → 网络和 Internet →
  高级网络设置 → 数据使用量），按当前网卡归档，每 30 秒刷新
- **历史面板**（右键悬浮窗 →「查看历史 / 详情」）：
  - 各软件每日明细矩阵：最近 7 天 × 每个软件的每日用量，首行为「总用量」
  - 数据与系统设置页一致，含投屏等本地传输（系统口径）
- **桌面悬浮窗**：
  - 半透明深色面板（80% 不透明度），可拖动到屏幕任意位置
  - 文字带柔和暗影，浅色壁纸下也清晰
  - 支持开机自启（写入「启动」文件夹的快捷方式，无黑窗口）

## 运行环境

- Windows 10 / 11
- **系统 Python 3.9+**（需自带 tkinter；注意部分精简版 Python 不含 tkinter）
- 依赖：

  ```bash
  pip install psutil pywin32 Pillow
  # 读取「今日总用量 / 各软件明细」还需要：
  pip install winrt-Windows.Networking.Connectivity winrt-Windows.Foundation winrt-Windows.Foundation.Collections
  ```

  未安装 winrt 绑定时，悬浮窗的速度 / 本次累计仍可用，仅系统记录相关功能不可用。

## 使用

```bash
# 默认模式：桌面悬浮窗
python usb_tether_monitor.py

# 独立打开历史用量面板
python usb_tether_monitor.py --history

# 指定网卡（默认自动识别）
python usb_tether_monitor.py --iface "以太网 2"

# 创建开机自启 / 桌面快捷方式 / 删除开机自启
python usb_tether_monitor.py --install-startup
python usb_tether_monitor.py --install-shortcut
python usb_tether_monitor.py --uninstall-startup
```

右键悬浮窗菜单：查看历史 / 详情、置顶 / 置底、退出。
**拖动**：按住卡片任意位置即可拖动（整卡都可接收鼠标事件）。

> 快捷方式默认指向当前运行本脚本的 `pythonw`（无控制台窗口）。如需指定其它
> 解释器（例如自带 tkinter 的系统 Python），设置环境变量
> `USB_TETHER_PYTHONW` 为对应 `pythonw.exe` 路径后再创建快捷方式。

## 实现要点

- **流量采样**：`psutil.net_io_counters(pernic=True)` 每秒读网卡字节数，差分得速率
- **网卡识别**：`pywin32` + WMI 查询 `Win32_NetworkAdapter`（不依赖 powershell）
- **按软件用量**：WinRT `ConnectionProfile.GetAttributedNetworkUsageAsync`
  （按网卡归档、按应用归因），在 MTA 工作线程中同步获取后固化为普通列表
- **悬浮窗渲染**：Tk 只做消息循环，内容用 Pillow 离屏渲染（2x 超采样抗锯齿），
  通过 `UpdateLayeredWindow`（ULW_ALPHA）贴到分层窗口：
  - 整卡铺一层带 alpha 的底色（alpha>0 的像素才能被鼠标命中）
  - 相比 Tk 的 `-transparentcolor` 色键方案，不会出现「透明区域点击穿透到桌面」
  - 文字统一垫一层高斯模糊暗影，任何壁纸下都清晰
- **已知边界**：同一网卡上系统记录不区分「上网流量」与「投屏等本地流量」；
  各软件明细依赖 Windows 自己的记录（约保留 30 天，未连接的日期显示 `—`）

## 许可

[MIT](LICENSE)
