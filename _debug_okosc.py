import win32gui, win32con
from pywinauto import Application
import config

# 창 핸들 찾아서 복원
def find_okosc_handle():
    result = []
    def cb(hwnd, _):
        t = win32gui.GetWindowText(hwnd)
        if any(kw in t for kw in config.OKOSC_WINDOW_KEYWORDS):
            result.append((hwnd, t))
    win32gui.EnumWindows(cb, None)
    return result

handles = find_okosc_handle()
print(f"OKOSC 핸들 목록: {handles}")

if handles:
    hwnd, title = handles[0]
    # 창 복원 (최소화 해제)
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)

    # Application 연결
    app = Application(backend="win32").connect(handle=hwnd)
    win = app.window(handle=hwnd)
    print(f"\n연결된 창: {win.window_text()}")
    print("\n=== 컨트롤 목록 (depth=4) ===")
    win.print_control_identifiers(depth=4)
