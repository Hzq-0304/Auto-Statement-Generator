"""在创建任何Tk窗口前启用DPI感知，避免Windows把界面当位图拉伸。"""
import ctypes
import sys
import tkinter.font as tkfont
from tkinter import ttk


def enable_dpi_awareness():
    if sys.platform != 'win32':
        return
    try:
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        fn = user32.SetProcessDpiAwarenessContext
        fn.argtypes = [ctypes.c_void_p]
        fn.restype = ctypes.c_bool
        if fn(ctypes.c_void_p(-4)):
            return
        if ctypes.get_last_error() == 5:
            return  # 已由安装版的manifest设置，不能再次覆盖。
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def window_dpi(window):
    if sys.platform == 'win32':
        try:
            fn = ctypes.windll.user32.GetDpiForWindow
            fn.argtypes = [ctypes.c_void_p]
            fn.restype = ctypes.c_uint
            value = fn(window.winfo_id())
            if value:
                return value
        except (AttributeError, OSError):
            pass
    return round(window.winfo_fpixels('1i'))


def apply_fonts(root, dpi=None):
    dpi = dpi or window_dpi(root)
    root.tk.call('tk', 'scaling', dpi / 72)
    families = set(tkfont.families(root))
    family = next((f for f in ('Microsoft YaHei UI', 'Microsoft YaHei', 'Noto Sans CJK SC') if f in families), 'Arial')
    for name in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont', 'TkHeadingFont', 'TkCaptionFont', 'TkSmallCaptionFont', 'TkTooltipFont'):
        tkfont.nametofont(name, root=root).configure(family=family, size=10 if name=='TkTooltipFont' else 11)
    font = tkfont.nametofont('TkDefaultFont', root=root)
    style = ttk.Style(root)
    style.configure('.', font='TkDefaultFont')
    style.configure('Treeview', font='TkDefaultFont', rowheight=font.metrics('linespace') + round(10*dpi/96))
    style.configure('Treeview.Heading', font='TkHeadingFont')
    style.configure('TButton', padding=(round(8*dpi/96), round(5*dpi/96)))
    root.option_add('*Font', 'TkDefaultFont')
    return dpi


def fit_window(window, width, height):
    scale = window_dpi(window) / 96
    sw, sh = window.winfo_screenwidth(), window.winfo_screenheight()
    width = min(round(width*scale), sw-50)
    height = min(round(height*scale), sh-100)
    window.geometry(f'{max(640,width)}x{max(480,height)}')
    window.minsize(min(round(850*scale),sw-80), min(round(620*scale),sh-100))
