import keyboard
import win32gui
import win32con
import threading
import time
import sys
import os
import pystray
from PIL import Image, ImageDraw


def create_icon():
    """创建托盘图标"""
    # 创建一个简单的蓝色方块图标
    image = Image.new('RGB', (64, 64), color='blue')
    dc = ImageDraw.Draw(image)
    dc.rectangle([16, 16, 48, 48], fill='white')
    return image


class TopmostTool:
    def __init__(self):
        self.topmost_windows = set()
        self.running = True
        self.tray_icon = None
        self.setup()

    def toggle_topmost(self):
        """切换窗口置顶状态"""
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                return

            title = win32gui.GetWindowText(hwnd) or "未知窗口"

            if hwnd in self.topmost_windows:
                # 取消置顶
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self.topmost_windows.remove(hwnd)
                # 更新托盘提示
                self.update_tray_tooltip(f"已取消置顶: {title}")
            else:
                # 设置置顶
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                self.topmost_windows.add(hwnd)
                # 更新托盘提示
                self.update_tray_tooltip(f"已置顶: {title}")
        except Exception as e:
            print(f"操作失败: {e}")

    def update_tray_tooltip(self, message):
        """更新托盘提示信息"""
        if self.tray_icon:
            self.tray_icon.title = f"窗口置顶工具 - {message}"

    def exit_program(self, icon=None, item=None):
        """退出程序"""
        self.running = False
        # 取消所有窗口的置顶
        for hwnd in list(self.topmost_windows):
            try:
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            except:
                pass
        if self.tray_icon:
            self.tray_icon.stop()
        os._exit(0)

    def show_status(self, icon=None, item=None):
        """显示当前状态"""
        status = f"已置顶 {len(self.topmost_windows)} 个窗口"
        self.update_tray_tooltip(status)

    def setup_tray(self):
        """设置系统托盘"""
        try:
            from pystray import MenuItem as item

            menu = (
                item('显示状态', self.show_status),
                item('退出', self.exit_program)
            )

            self.tray_icon = pystray.Icon(
                "topmost_tool",
                create_icon(),
                "窗口置顶工具 - 就绪",
                menu
            )

            # 在单独线程中运行托盘图标
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
        except Exception as e:
            print(f"托盘图标创建失败: {e}")

    def monitor_topmost(self):
        """监控并保持置顶状态"""
        while self.running:
            for hwnd in list(self.topmost_windows):
                if win32gui.IsWindow(hwnd):
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            time.sleep(0.1)

    def setup(self):
        """初始化设置"""
        # 设置快捷键
        keyboard.add_hotkey('ctrl+alt+t', self.toggle_topmost)
        keyboard.add_hotkey('ctrl+alt+q', self.exit_program)

        # 设置系统托盘
        self.setup_tray()

        # 启动监控线程
        threading.Thread(target=self.monitor_topmost, daemon=True).start()

    def run(self):
        """运行程序"""
        try:
            keyboard.wait()
        except KeyboardInterrupt:
            self.exit_program()


if __name__ == "__main__":
    # 检查依赖
    try:
        import pystray
        from PIL import Image
    except ImportError:
        print("请安装依赖: pip install pystray pillow")
        sys.exit(1)

    TopmostTool().run()