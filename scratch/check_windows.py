import ctypes
from ctypes import wintypes

user32 = ctypes.windll.user32

def enum_windows_callback(hwnd, extra):
    lpdwProcessId = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(lpdwProcessId))
    if lpdwProcessId.value == 25584:
        length = user32.GetWindowTextLengthW(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buff, length + 1)
        visible = user32.IsWindowVisible(hwnd)
        print(f'HWND {hwnd}: Visible={visible}, Title="{buff.value}"')
    return True

WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
user32.EnumWindows(WNDENUMPROC(enum_windows_callback), 0)

