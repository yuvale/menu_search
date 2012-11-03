# coding: utf-8

import win32gui as w32g
import win32gui_struct as w32g_s
import win32con
import win32api
import re
import locale
import sys

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
    item_type, _, _, hSubMenu, _, _, _, text, _ = \
        w32g_s.UnpackMENUITEMINFO(menu_item_info[0])
    text = escape_menu_text(text.decode(DEFAULT_ENCODING))
    return item_type, text, hSubMenu

def escape_menu_text(t):
    "Replaces '&&' with '&', removes all other '&'s."
    return MENU_TEXT_ESCAPE_RE.sub(r'\1', t)
    
def walk_menu(hMenu):
    def walk_menu_helper(hMenu, path):
        items = []
        sub_menus = []
        
        count = w32g.GetMenuItemCount(hMenu)
        for i in xrange(count):
            item_type, text, hSubMenu = get_menu_item_info(hMenu, i)
            if item_type & (win32con.MF_SEPARATOR | win32con.MF_BITMAP):
                # Skip non-string items.
                continue
            if hSubMenu == 0:
                item_id = w32g.GetMenuItemID(hMenu, i)
                items.append((text, item_id))
            else:
                sub_menus.append((text, hSubMenu))
        
        if items:
            yield path, items
        
        for text, hSubMenu in sub_menus:
            new_path = path + [text]
            for item in walk_menu_helper(hSubMenu, new_path):
                yield item
    
    for item in walk_menu_helper(hMenu, []):
        yield item
    
def run_menu_item(hwnd, menu_id):
    win32api.PostMessage(hwnd,
                         win32con.WM_COMMAND,
                         win32api.MAKELONG(menu_id, 0),
                         0)

def main():
    title = sys.argv[1]
    
    win_handle = get_window_handle(title)
    if win_handle is None:
        print 'No window handle found.'
        return
    print 'Found window handle %d.' % win_handle
    
    menu_handle = w32g.GetMenu(win_handle)
    if menu_handle == 0:
        print 'No menu found.'
        return
    
    for path, items in walk_menu(menu_handle):
        path_text = '->'.join(path)
        for text, id in items:
            print '%s: %s [%d]' % (path_text, text, id)
            
if __name__ == '__main__':
    main()