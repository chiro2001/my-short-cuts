import platform
from pynput import keyboard
from tkinter import *
from tkinter import messagebox
import getopt
from tkinter import ttk
import time
import json
import os
import threading
import traceback
import win32api
import win32con
import win32gui_struct
import win32gui

_main = None


class SysTrayIcon(object):
    QUIT = 'QUIT'
    SPECIAL_ACTIONS = [QUIT]
    FIRST_ID = 1314

    def __init__(self,
                 icon,
                 hover_text,
                 menu_options,
                 on_quit=None,
                 default_menu_index=None,
                 window_class_name=None, ):
        self.icon = icon
        self.hover_text = hover_text
        self.on_quit = on_quit

        menu_options = menu_options + (('退出', None, self.QUIT),)
        self._next_action_id = self.FIRST_ID
        self.menu_actions_by_id = set()
        self.menu_options = self._add_ids_to_menu_options(list(menu_options))
        self.menu_actions_by_id = dict(self.menu_actions_by_id)
        del self._next_action_id

        self.default_menu_index = (default_menu_index or 0)
        self.window_class_name = window_class_name or "SysTrayIconPy"

        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.refresh_icon,
                       win32con.WM_DESTROY: self.destroy,
                       win32con.WM_COMMAND: self.command,
                       win32con.WM_USER + 20: self.notify, }
        # 注册窗口类。
        window_class = win32gui.WNDCLASS()
        window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = self.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map  # 也可以指定wndproc.
        self.classAtom = win32gui.RegisterClass(window_class)

    def show_icon(self):
        # 创建窗口。
        hinst = win32gui.GetModuleHandle(None)
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(self.classAtom,
                                          self.window_class_name,
                                          style,
                                          0,
                                          0,
                                          win32con.CW_USEDEFAULT,
                                          win32con.CW_USEDEFAULT,
                                          0,
                                          0,
                                          hinst,
                                          None)
        win32gui.UpdateWindow(self.hwnd)
        self.notify_id = None
        self.refresh_icon()

        win32gui.PumpMessages()

    def show_menu(self):
        menu = win32gui.CreatePopupMenu()
        self.create_menu(menu, self.menu_options)
        # win32gui.SetMenuDefaultItem(menu, 1000, 0)

        pos = win32gui.GetCursorPos()
        # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu,
                                win32con.TPM_LEFTALIGN,
                                pos[0],
                                pos[1],
                                0,
                                self.hwnd,
                                None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    def destroy(self, hwnd, msg, wparam, lparam):
        if self.on_quit: self.on_quit(self)  # 运行传递的on_quit
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # 退出托盘图标

    def notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:  # 双击左键
            pass  # self.execute_menu_option(self.default_menu_index + self.FIRST_ID)
        elif lparam == win32con.WM_RBUTTONUP:  # 单击右键
            self.show_menu()
        elif lparam == win32con.WM_LBUTTONUP:  # 单击左键
            nid = (self.hwnd, 0)
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
            win32gui.PostQuitMessage(0)  # 退出托盘图标
            if _main:
                _main.root.deiconify()
        return True

    def _add_ids_to_menu_options(self, menu_options):
        result = []
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
                self.menu_actions_by_id.add((self._next_action_id, option_action))
                result.append(menu_option + (self._next_action_id,))
            else:
                result.append((option_text,
                               option_icon,
                               self._add_ids_to_menu_options(option_action),
                               self._next_action_id))
            self._next_action_id += 1
        return result

    def refresh_icon(self, **data):
        hinst = win32gui.GetModuleHandle(None)
        if os.path.isfile(self.icon):  # 尝试找到自定义图标
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst,
                                       self.icon,
                                       win32con.IMAGE_ICON,
                                       0,
                                       0,
                                       icon_flags)
        else:  # 找不到图标文件 - 使用默认值
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if self.notify_id:
            message = win32gui.NIM_MODIFY
        else:
            message = win32gui.NIM_ADD
        self.notify_id = (self.hwnd,
                          0,
                          win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                          win32con.WM_USER + 20,
                          hicon,
                          self.hover_text)
        win32gui.Shell_NotifyIcon(message, self.notify_id)

    def create_menu(self, menu, menu_options):
        for option_text, option_icon, option_action, option_id in menu_options[::-1]:
            if option_icon:
                option_icon = self.prep_menu_icon(option_icon)

            if option_id in self.menu_actions_by_id:
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                wID=option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                self.create_menu(submenu, option_action)
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)

    def prep_menu_icon(self, icon):
        # 首先加载图标。
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hdcBitmap = win32gui.CreateCompatibleDC(0)
        hdcScreen = win32gui.GetDC(0)
        hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x, ico_y)
        hbmOld = win32gui.SelectObject(hdcBitmap, hbm)
        # 填满背景。
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
        win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)
        # "GetSysColorBrush返回缓存的画笔而不是分配新的画笔。"
        #  - 暗示没有DeleteObject
        # 画出图标
        win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
        win32gui.SelectObject(hdcBitmap, hbmOld)
        win32gui.DeleteDC(hdcBitmap)

        return hbm

    def command(self, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        self.execute_menu_option(id)

    def execute_menu_option(self, id):
        menu_action = self.menu_actions_by_id[id]
        if menu_action == self.QUIT:
            win32gui.DestroyWindow(self.hwnd)
        else:
            menu_action(self)


def get_lower_case_name(text):
    lst = []
    for index, char in enumerate(text):
        if char.isupper() and index != 0:
            lst.append("_")
        lst.append(char)

    return "".join(lst).lower()


class MyShortCuts:
    class Settings:
        FILENAME = 'settings.json'
        DEFAULT = {
            'saves': [
                {
                    'keys': ('Ctrl', 'Alt', 'T'),
                    'command': 'cmd',
                    'path': 'D:\\Programs',
                    'time': time.time()
                }
            ]
        }

        def __init__(self):
            if not os.path.exists(MyShortCuts.Settings.FILENAME):
                self.new()
            self.data = {}
            self.load()

        @staticmethod
        def new():
            with open(MyShortCuts.Settings.FILENAME, 'w') as f:
                json.dump(MyShortCuts.Settings.DEFAULT, f)

        def load(self):
            with open(MyShortCuts.Settings.FILENAME, 'r') as f:
                try:
                    self.data = json.load(f)
                except json.decoder.JSONDecodeError:
                    self.new()
                    self.load()

        def save(self):
            if 'saves' in self.data:
                tmp = []
                for d in self.data['saves']:
                    if 'keys' in d:
                        d['keys'] = list(d['keys'])
                        tmp.append(d)
                self.data['saves'] = tmp
            with open(MyShortCuts.Settings.FILENAME, 'w') as f:
                json.dump(self.data, f)

    @staticmethod
    def cmp_keys(keys1: set, keys2: set) -> bool:
        # 设置优先级和对应按键
        return keys1 == keys2

    @staticmethod
    def parse_keys_str(keys_str: str) -> set:
        keys = keys_str.split('+')
        # 自动去重
        return set(keys)

    def make_map_to_val(self):
        s1 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        # for s in s1:
        #     self.map_to_code[s.lower()] = s
        special = {
            'ctrl_l': 'Ctrl',
            'ctrl_r': 'Ctrl',
            'alt_l': 'Alt',
            'alt_r': 'Alt',
            'cmd': 'Win',
            '¿': '/',
            '¼': ',',
            'Þ': "'",
            'º': ';',
            'Û': '[',
            'Ý': ']',
            '½': '-',
            '»': '=',
            '¾': '.',
            'À': '`',
        }
        comm = ['Shift', 'Tab', 'Esc', 'Space', 'Enter', 'CapsLock',
                'Backspace', 'Left', 'Right', 'Up', 'Down',
                'Insert', 'End', 'PageDown', 'PageUp', 'Delete', 'Home',
                'PrintScreen', 'ScrollLock', 'Pause']
        comm_dict = {}
        functions = {}
        for i in range(1, 13, 1):
            functions['f%d' % i] = 'F%d' % i
        for c in comm:
            comm_dict[get_lower_case_name(c)] = c
        self.map_to_value = {}
        for s in s1:
            self.map_to_value[s] = s
        for i in range(0, 10):
            self.map_to_value[str(i)] = str(i)
        self.map_to_value.update(special)
        self.map_to_value.update(comm_dict)
        self.map_to_value.update(functions)
        #         for ignore in ignores:
        #             del self.map_to_value[ignore]
        # print(self.map_to_value)

    """
    请先调用上一个函数
    """

    # def make_map_to_code(self):
    #     if len(self.map_to_value) == 0:
    #         self.make_map_to_val()
    #     self.map_to_code = {}
    #     for m in self.map_to_value:
    #         if type(self.map_to_value[m]) is list:
    #             pass

    def key_code(self, key):
        try:
            key_code = chr(key.vk)
        except AttributeError:
            key_code = key.name
        # print('key_code =', key_code)
        return key_code

    @staticmethod
    def keys_str(keys: set) -> str:
        keys_str = ''
        keys_tmp = list(keys)
        keys_tmp.sort()
        for k in keys_tmp:
            keys_str = keys_str + '+' + k
        # 注意排序...
        keys_str = keys_str[1:]
        return keys_str

    def key_val(self, key) -> str or None:
        key_code = self.key_code(key)
        if key_code in self.map_to_value:
            return self.map_to_value[key_code]
        return None

    def key_code_val(self, key_code) -> str or None:
        if key_code in self.map_to_value:
            return self.map_to_value[key_code]
        return None

    UNIQUE_FILE = "%s\\my-short-cuts.tmp" % os.environ.get('TEMP')

    def __init__(self, root=None, silent=False):
        self.silent = silent
        self.ignores = ['media_play_pause', 'media_previous', 'media_next']
        self.call_up = ['Win+Alt+K', 'Win+Ctrl+K']

        # 提示是否独占
        if os.path.exists(self.UNIQUE_FILE):
            result = messagebox.askyesno('检测到已经打开了一个程序', '是否退出？')
            if result:
                exit(1)

        try:
            with open(self.UNIQUE_FILE, 'w') as f:
                f.write('my-short-cuts')
        except Exception:
            traceback.print_exc()

        # self.map_to_code = {}
        self.map_to_value = {}
        self.make_map_to_val()

        self.root = root
        if self.root is None:
            self.root = Tk()
        self.title = "快捷键管理"
        self.root.title(self.title)
        # 禁止最大化
        self.root.resizable(width=False, height=False)
        # 主表格
        self.table = ttk.Treeview(self.root)
        self.table['column'] = ['keys', 'command', 'path']
        self.table.grid(row=0, columnspan=3, column=0)

        # 设置宽度
        self.table.column('#0', width=150)
        self.table.column('keys', width=150)
        self.table.column('command', width=280)
        self.table.column('path', width=400)

        # 设置列名
        self.table.heading('keys', text='快捷键')
        self.table.heading('command', text='命令')
        self.table.heading('path', text='路径')
        self.table.bind('<Double-Button-1>', self.clicked)
        self.table.bind('<Return>', self.clicked)

        Button(self.root, text="增加", command=self.add_item).grid(row=1, column=0, sticky=W + E)
        Button(self.root, text="删除", command=self.delete_item).grid(row=1, column=1, sticky=W + E)
        Button(self.root, text="最小化", command=self.enter_mini_mode).grid(row=1, column=2, sticky=W + E)

        self.settings = None
        self.data = None
        self.data_set = {}
        self.data_init()

        self.update_data()

        self.var_keys = None
        self.var_cmd = None
        self.var_path = None
        self.top = None

        # 已经按下的按键
        self.pressed = set([])

        ###########################     开始托盘程序嵌入     #####################################
        icons = os.getcwd() + r'\icon.ico'
        # print(icons)
        hover_text = "快捷键管理"  # 悬浮于图标上方时的提示
        menu_options = ()
        self.sysTrayIcon = SysTrayIcon(icons, hover_text, menu_options, on_quit=self.exit, default_menu_index=1)

        self.root.bind("<Unmap>", lambda event: self.Unmap() if self.root.state() == 'iconic' else False)
        self.root.protocol('WM_DELETE_WINDOW', self.exit)
        self.root.resizable(0, 0)

    def data_init(self):
        self.settings = MyShortCuts.Settings()
        self.data = list(self.settings.data['saves'])
        tmp = []
        for d in self.data:
            d['keys'] = set(d['keys'])
            tmp.append(d)
        self.data = tmp
        self.data_set = {}

    def update_data(self):
        self.clear_all()
        for item in self.data:
            keys_set = item.get('keys', set())
            keys_str = self.keys_str(keys_set)
            command = item.get('command', 'echo')
            path = item.get('path', 'D:\\Programs')
            time_stamp = item.get('time', time.time())
            time_array = time.localtime(time_stamp)
            style_time = time.strftime("%Y-%m-%d %H:%M:%S", time_array)
            self.table.insert('', self.data.index(item),
                              text=style_time,
                              values=(keys_str, command, path))
            self.data_set[keys_str] = item

    def add_item(self):
        self.settings.data['saves'].append({
            "keys": ["None"],
            "command": "cmd",
            "path": "D:\\",
            "time": time.time()
        })
        self.settings.save()
        self.data_init()
        self.update_data()

    def delete_item(self):
        item = self.get_item_now()
        item_text = self.table.item(item, "values")
        data = self.data_set[item_text[0]]
        # print(data)
        if item is None:
            return
        keys = data.get('keys')
        # print(keys)
        for d in self.data:
            if d['keys'] == keys:
                del self.data[self.data.index(d)]
                self.settings.data['saves'] = self.data
                self.settings.save()
                self.data_init()
                self.update_data()
                return
        print("WARNING: delete none!")

    def get_item_now(self):
        items = self.table.selection()
        if len(items) == 0:
            return None
        item = items[0]
        return item

    def clicked(self, event=None):
        item = self.get_item_now()
        item_text = self.table.item(item, "values")
        data = self.data_set[item_text[0]]
        # print(data)
        keys = ''
        for k in data['keys']:
            keys = keys + '+' + k
        # 注意排序...
        keys = keys[1:]
        top = Toplevel()
        if platform.system() == 'Windows':
            top.attributes("-toolwindow", 1)
            top.attributes("-topmost", 1)
        top.resizable(width=False, height=False)

        self.var_keys = StringVar(top)
        self.var_cmd = StringVar(top)
        self.var_path = StringVar(top)

        self.var_keys.set(keys)
        self.var_cmd.set(data['command'])
        self.var_path.set(data['path'])

        top.title('设置快捷键')
        Label(top, text='快捷键').grid(row=0, column=0)
        Entry(top, textvariable=self.var_keys).grid(row=0, column=1)
        Label(top, text='命令').grid(row=1, column=0)
        Entry(top, textvariable=self.var_cmd).grid(row=1, column=1)
        Label(top, text='路径').grid(row=2, column=0)
        Entry(top, textvariable=self.var_path).grid(row=2, column=1)
        Button(top, text='确定', command=self.confirm_settings).grid(row=3, columnspan=2, stick=W + E)
        self.top = top
        self.top.mainloop()

    def confirm_settings(self, event=None):
        # print(self.var_cmd.get(), self.var_keys.get())
        keys = self.parse_keys_str(self.var_keys.get())
        cmd = self.var_cmd.get()
        path = self.var_path.get()
        # print(keys)
        item = self.get_item_now()
        item_text = self.table.item(item, "values")
        data = self.data_set[item_text[0]]
        # print(data)
        try:
            index = self.data.index(data)
        except ValueError:
            return
        data['keys'] = keys
        data['command'] = cmd
        data['path'] = path
        self.data[index] = data
        self.settings.data['saves'] = self.data
        self.settings.save()
        self.data_init()
        self.update_data()
        self.top.destroy()

    def clear_all(self):
        x = self.table.get_children()
        for item in x:
            self.table.delete(item)

    def enter_mini_mode(self):
        self.root.state('iconic')

    def loop(self):
        th = threading.Thread(target=self.hook)
        th.setDaemon(True)
        th.start()
        if self.silent:
            # self.root.withdraw()
            self.enter_mini_mode()
        self.root.mainloop()

    @staticmethod
    def run_cmd(cmd, path):
        try:
            # print('#1:', os.path.abspath(os.path.curdir))
            os.chdir(path)
            # base = os.path.abspath(os.path.curdir)
            # print('#2:', os.path.abspath(os.path.curdir))
            os.system('start ' + cmd)
        except Exception:
            traceback.print_exc()

    def start_cmd(self, cmd: str, path="D:\\"):
        th = threading.Thread(target=self.run_cmd, args=(cmd, path))
        th.setDaemon(True)
        base = os.path.abspath(os.path.curdir)
        th.start()
        time.sleep(1)
        os.chdir(base)
        # print("#3:", os.path.abspath(os.path.curdir))

    def on_press(self, key):
        key_code = self.key_code(key)
        if key_code not in self.pressed:
            self.pressed.add(key_code)
            print(self.pressed)
            keys = set(map(self.key_code_val, list(self.pressed)))
            # print(keys)
            # TODO: 快捷键调出设置
            # print(self.cmp_keys(keys, self.parse_keys_str(self.call_up[0])))
            for d in self.data:
                # print('test:', d)
                if self.cmp_keys(d['keys'], keys):
                    self.start_cmd(d['command'], d.get('path', 'D:\\Programs'))
                    self.pressed = set([])

    def on_release(self, key):
        key_code = self.key_code(key)
        if key_code in self.pressed:
            self.pressed.remove(key_code)
        else:
            self.pressed = set([])
        # print(list(map(self.key_code_val, list(self.pressed))))

    def hook(self):
        with keyboard.Listener(
                on_press=self.on_press,
                on_release=self.on_release)as listener:
            listener.join()

    def switch_icon(self, _sysTrayIcon, icons='D:\\2.ico'):
        _sysTrayIcon.icon = icons
        _sysTrayIcon.refresh_icon()
        # 点击右键菜单项目会传递SysTrayIcon自身给引用的函数，所以这里的_sysTrayIcon = self.sysTrayIcon

    def Unmap(self):
        self.root.withdraw()
        self.sysTrayIcon.show_icon()

    def exit(self, _sysTrayIcon=None):
        # 删除独占文件
        if os.path.exists(self.UNIQUE_FILE):
            try:
                os.remove(self.UNIQUE_FILE)
            except Exception:
                traceback.print_exc()
        self.root.destroy()
        # print('exit...')


if __name__ == '__main__':
    opts, args = getopt.getopt(sys.argv[1:], '-s', [])
    _silent = False
    for name, val in opts:
        if '-s' == name:
            _silent = True
    _main = MyShortCuts(silent=_silent)
    _main.loop()
