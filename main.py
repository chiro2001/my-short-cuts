import platform
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import time
import json
import os


class MyShortCuts:
    class Settings:
        FILENAME = 'settings.json'
        DEFAULT = {
            'saves': [
                {
                    'keys': ['Ctrl', 'Alt', 'T'],
                    'command': 'cmd',
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
                self.data = json.load(f)

        def save(self):
            with open(MyShortCuts.Settings.FILENAME, 'w') as f:
                json.dump(self.data, f)

    @staticmethod
    def make_keys_str(keys: list) -> str:
        # 设置优先级和对应按键
        pass

    @staticmethod
    def parse_keys_str(keys_str: str) -> list:
        pass

    def make_map_to_val(self):
        s1 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        for s in s1:
            self.map_to_code[s.lower()] = s
        special = {
            'ctrl_l': 'Ctrl',
            'ctrl_r': 'Ctrl',
            'alt_l': 'Alt',
            'alt_r': 'Alt',
            'cmd': 'Win',
        }
        comm = ['Shift', 'Tab', 'Esc', 'Space', 'Enter', 'Backspace', 'Left', 'Right', 'Up', 'Down']

    def __init__(self, root=None):
        self.map_to_code = {}
        self.map_to_value = {}

        self.settings = MyShortCuts.Settings()
        self.root = root
        if self.root is None:
            self.root = Tk()
        self.title = "快捷键管理"
        self.root.title(self.title)
        # 禁止最大化
        self.root.resizable(width=False, height=False)
        # 主表格
        self.table = ttk.Treeview(self.root)
        self.table['column'] = ['keys', 'command']
        self.table.pack()

        # 设置宽度
        self.table.column('keys', width=150)
        self.table.column('command', width=280)
        self.table.column('#0', width=150)

        # 设置列名
        self.table.heading('keys', text='快捷键')
        self.table.heading('command', text='命令')
        self.table.bind('<Double-Button-1>', self.clicked)
        self.table.bind('<Return>', self.clicked)

        self.data = list(self.settings.data['saves'])
        self.data_set = {}

        self.update_data()

        self.var_keys = None
        self.var_cmd = None
        self.top = None

    def update_data(self):
        self.clear_all()
        for item in self.data:
            keys_list = item.get('keys', [])
            keys = ''
            for k in keys_list:
                keys = keys + '+' + k
            # 注意排序...
            keys = keys[1:]
            command = item.get('command', 'echo')
            time_stamp = item.get('time', time.time())
            time_array = time.localtime(time_stamp)
            style_time = time.strftime("%Y-%m-%d %H:%M:%S", time_array)
            self.table.insert('', self.data.index(item),
                              text=style_time,
                              values=(keys, command))
            self.data_set[keys] = item

    def clicked(self, event=None):
        item = self.table.selection()[0]
        item_text = self.table.item(item, "values")
        data = self.data_set[item_text[0]]
        print(data)
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

        self.var_keys.set(keys)
        self.var_cmd.set(data['command'])

        top.title('设置快捷键')
        Label(top, text='快捷键').grid(row=0, column=0)
        Entry(top, textvariable=self.var_keys).grid(row=0, column=1)
        Label(top, text='命令').grid(row=1, column=0)
        Entry(top, textvariable=self.var_cmd).grid(row=1, column=1)
        Button(top, text='确定', command=self.confirm_settings).grid(row=2, columnspan=2, stick=W + E)
        self.top = top
        self.top.mainloop()

    def confirm_settings(self, event=None):
        print(self.var_cmd.get())
        self.top.destroy()

    def clear_all(self):
        x = self.table.get_children()
        for item in x:
            self.table.delete(item)

    def loop(self):
        self.root.mainloop()


if __name__ == '__main__':
    _main = MyShortCuts()
    _main.loop()
