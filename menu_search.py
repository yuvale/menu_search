# coding: utf-8

import win32gui as w32g
import win32gui_struct as w32g_s
import win32con
import win32api
import re
import locale

MENU_TEXT_ESCAPE_RE = re.compile(r'&(.)')

_, DEFAULT_ENCODING = locale.getdefaultlocale()

def get_window_handle(text):
    class Arg(object):
        def __init__(self, text):
            self.text = text
            self.hwnd = None

    def enum_callback(hwnd, arg):
        t = w32g.GetWindowText(hwnd).decode(DEFAULT_ENCODING)
        if arg.hwnd == None and arg.text in t:
            arg.hwnd = hwnd
        return 1
        
    arg = Arg(text)
    w32g.EnumWindows(enum_callback, arg)
    return arg.hwnd

def get_menu_item_info(hMenu, index):
    menu_item_info = w32g_s.EmptyMENUITEMINFO()
    w32g.GetMenuItemInfo(hMenu, index, True, menu_item_info[0])
    _, _, _, hSubMenu, _, _, _, text, _ = \
        w32g_s.UnpackMENUITEMINFO(menu_item_info[0])
    text = escape_menu_text(text.decode(DEFAULT_ENCODING))
    return text, hSubMenu

def escape_menu_text(t):
    "Replaces '&&' with '&', removes all other '&'s."
    return MENU_TEXT_ESCAPE_RE.sub(r'\1', t)
    
def traverse_menu(hMenu):
    count = w32g.GetMenuItemCount(hMenu)
    items = []
    for i in xrange(count):
        text, hSubMenu = get_menu_item_info(hMenu, i)
        item_id = w32g.GetMenuItemID(hMenu, i)
        if hSubMenu == 0:
            items.append((text, item_id, None))
        else:
            items.append((text, item_id, (hSubMenu, traverse_menu(hSubMenu))))
    return items
    
def run_menu_item(hwnd, menu_id):
    win32api.PostMessage(hwnd,
                         win32con.WM_COMMAND,
                         win32api.MAKELONG(menu_id, 0),
                         0)