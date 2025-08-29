from inactive import Inactive
import win32gui
def main():
    hwnd = win32gui.FindWindow(None, "无标题 - 记事本")
    hwnd1 = win32gui.FindWindow(None, "记事本")
    print(hwnd,hwnd1)
    inactive = Inactive(hwnd)
    inactive.press('a')
    inactive.press('space')
    pass

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(e)