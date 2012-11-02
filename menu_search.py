# coding: utf-8

import win32gui as w32g
import win32gui_struct as w32g_s
import win32con
import win32api

def get_window_handle(text):
    class Arg(object):
        def __init__(self, text):
            self.text = text
            self.hwnd = None

    def enum_callback(hwnd, arg):
        t = w32g.GetWindowText(hwnd)
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
    return text, hSubMenu
    
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