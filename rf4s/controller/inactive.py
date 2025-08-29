import win32gui
import win32ui
import win32con
import win32api
from PIL import Image
from pyscreeze import Box

PRIMARY = "primary"

KEY_MAP = {
    'a':97,
    'b':98,
    'c':99,
    'd':100,
    'e':101,
    'f':102,
    'g':103,
    'h':104,
    'i':105,
    'j':106,
    'k':107,
    'l':108,
    'm':109,
    'n':110,
    'o':111,
    'p':112,
    'q':113,
    'r':114,
    's':115,
    't':116,
    'u':117,
    'v':118,
    'w':119,
    'x':120,
    'y':121,
    'z':122,
    '0':48,
    '1':49,
    '2':50,
    '3':51,
    '4':52,
    '5':53,
    '6':54,
    '7':55,
    '8':56,
    '9':57,
    ' ':32,
    'enter':13,
    'backspace':8,
    'tab':9,
    'shift':16,
    'ctrl':17,
    'alt':18,
    'esc':27,
    'left':37,
    'up':38,
    'right':39,
    'down':40,
    'f1':112,
    'f2':113,
    'f3':114,
    'f4':115,
    'f5':116,
    'f6':117,
    'f7':118,
    'f8':119,
    'f9':120,
    'f10':121,
    'f11':122,
    'f12':123,
    'print':42,
    'scroll':45,
    'pause':19,
    'insert':45,
    'delete':46,
    'home':36,
    'end':35,
    'pageup':33,
    'pagedown':34,
    'numlock':144,
    'capslock':20,
    'semicolon':59,
    'equals':61,
    'minus':189,
    'period':190,
    'slash':191,
    'backquote':192,
    'bracketleft':219,
    'bracketright':221,
    'comma':188,
    'period':190,
    'quote':222,
    'semicolon':59,
    'slash':191,
    'backslash':220,
    'grave':96,
    'numpad0':96,
    'numpad1':97,
    'numpad2':98,
    'numpad3':99,
    'numpad4':100,
    'numpad5':101,
    'numpad6':102,
    'numpad7':103,
    'numpad8':104,
    'numpad9':105,
    'numpadmultiply':106,
    'numpadadd':107,
    'numpadsubtract':109,
    'numpaddecimal':110,
    'numpaddivide':111,
    'numpadenter':13,
    'numpadmultiply':55,
    'numpadadd':78,
    'numpadsubtract':74,
    'numpaddecimal':83,
    'numpaddivide':53,
    'space':32,
    'tab':9,
}
class Inactive:
    def convert_image(self,bitmap):
        bmpinfo = bitmap.GetInfo()
        bmpbits = bitmap.GetBitmapBits(True)
        pil_im = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpbits, 'raw', 'BGRX', 0, 1)
        return pil_im


    def screenshot(self,bound:Box):
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(self.dcObj, bound.width, bound.height)
        self.cDC.SelectObject(dataBitMap)
        self.cDC.BitBlt((0,0),(bound.width, bound.height) ,self.dcObj, (bound.x,bound.y), win32con.SRCCOPY)
        win32gui.DeleteObject(dataBitMap.GetHandle())
        return self.convert_image(dataBitMap)

    def __init__(self,hwnd) -> None:
        self.hwnd =hwnd
        self.wDC = win32gui.GetWindowDC(hwnd)
        self.dcObj=win32ui.CreateDCFromHandle( self.wDC)
        self.cDC= self.dcObj.CreateCompatibleDC()
        self.PyCWnd = win32ui.CreateWindowFromHandle(hwnd)
        pass

    def destroy(self,hwnd)->None:
         # Free Resources
        self.dcObj.DeleteDC()
        self.cDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, self.wDC)


    def press(self,key:str)->None:
        self.keyDown(key)
        self.keyUp(key)

    def keyDown(self,key:str)->None:
        vk = KEY_MAP[key.lower()]
        if vk is None:
            raise ValueError(f"Invalid key: {key}")
        win32gui.PostMessage(self.hwnd,win32con.WM_KEYDOWN,0x41,win32api.MAKEWORD(1,win32api.MapVirtualKey(0x41,0)<<16))

    def keyUp(self,key:str)->None:
        vk = KEY_MAP[key.lower()]
        if vk is None:
            raise ValueError(f"Invalid key: {key}")
        win32gui.PostMessage(self.hwnd,win32con.WM_KEYUP,0x41,win32api.MAKEWORD(0,win32api.MapVirtualKey(0x41,0)<<16))
        

    def mouseDown(self,x=None, y=None, button=PRIMARY):
        if button==PRIMARY:
            self.PyCWnd.SendMessage(win32con.WM_LBUTTONDOWN,win32con.VK_SPACE,0)
            self.PyCWnd.SendMessage(win32con.WM_LBUTTONUP)
        else:
            self.PyCWnd.SendMessage(win32con.WM_RBUTTONDOWN)
            self.PyCWnd.SendMessage(win32con.WM_RBUTTONUP)

    def mouseMove(self,x:float,y:float):
        lParam = win32api.MAKELONG(x, y)
        self.PyCWnd.SendMessage(win32con.WM_MOUSEMOVE,0,lParam)