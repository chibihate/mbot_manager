import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32con
from pywinauto import findwindows
import ctypes

WINDOW_WIDTH = 700
WINDOW_HEIGHT = 300
TOP_LEFT_X = 0
TOP_LEFT_Y = 0

class MBot():
    def __init__(self, mbot):
        self.mbot = mbot
        self.name = ""
        self.spy_player_checkbox_state = None
        self.refresh_spy_button = None
        self.spy_combo = None

    def __str__(self):
        return f"name: {self.mbot.name}"

    def is_valid(self):
        return win32gui.IsWindow(self.mbot.handle)

    def find_nth_element_by_name(self, name, index):
        if not self.is_valid():
            return

        children = self.mbot.children()
        for i, child in enumerate(children):
            nth_child = children[i + index] if i + index < len(children) else None
            if child.name == name:
                return nth_child

    def get_name(self):
        if not self.is_valid():
            return

        if self.name == "":
            element = self.find_nth_element_by_name("Hide client after relogin", 1)
            name = element.name
            split_name = name.split(":")
            if len(split_name) > 1 and split_name[0] == "Name":
                self.name = split_name[1].strip()
            else:
                self.name = split_name[0].strip()
        return self.name

    def set_spy_player_checkbox_state(self):
        if not self.is_valid():
            return

        if not self.spy_player_checkbox_state:
            self.spy_player_checkbox_state = self.find_nth_element_by_name("Spy", 6)

        state = win32gui.SendMessage(self.spy_player_checkbox_state.handle, win32con.BM_GETCHECK, 0, 0)
        if state != win32con.BST_CHECKED:
            win32gui.PostMessage(self.spy_player_checkbox_state.handle, win32con.BM_CLICK, 0, 0)

    def refresh_spy(self):
        if not self.is_valid():
            return

        if not self.refresh_spy_button:
            self.refresh_spy_button = self.find_nth_element_by_name("Spy", 5)

        win32gui.PostMessage(self.refresh_spy_button.handle, win32con.BM_CLICK, 0, 0)

    def get_list_of_guild(self):
        if not self.is_valid():
            return

        if not self.spy_combo:
            self.spy_combo = self.find_nth_element_by_name("Spy", 10)

        guilds = set()
        count = win32gui.SendMessage(self.spy_combo.handle, win32con.CB_GETCOUNT, 0, 0)
        for i in range(count):
            length = win32gui.SendMessage(self.spy_combo.handle, win32con.CB_GETLBTEXTLEN, i, 0)
            buffer = ctypes.create_unicode_buffer(length + 1)
            win32gui.SendMessage(self.spy_combo.handle, win32con.CB_GETLBTEXT, i, buffer)
            value =  buffer.value
            if "[" in value and "]" in value:
                guild = value.split("[")[1].split("]")[0]
                guild = guild.split("*")[0].strip()
                guilds.add(guild)
        return guilds

    def get_list_of_member_in_guild(self, list_of_guild):
        if not self.is_valid():
            return

        if not self.spy_combo:
            self.spy_combo = self.find_nth_element_by_name("Spy", 10)

        members = []
        count = win32gui.SendMessage(self.spy_combo.handle, win32con.CB_GETCOUNT, 0, 0)
        for i in range(count):
            length = win32gui.SendMessage(self.spy_combo.handle, win32con.CB_GETLBTEXTLEN, i, 0)
            buffer = ctypes.create_unicode_buffer(length + 1)
            win32gui.SendMessage(self.spy_combo.handle, win32con.CB_GETLBTEXT, i, buffer)
            value =  buffer.value
            if "[" in value and "]" in value:
                guild = value.split("[")[1].split("]")[0]
                guild = guild.split("*")[0].strip()
                if guild in list_of_guild:
                    members.append(value)
        return members

class FilterGuild(tk.Tk):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Filter Guild")
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.bot_name_list = None
        self.guild_name_list = None
        self.mbot_list = []
        self.guild_combo = None
        self.guild_combo_var = tk.StringVar()
        self.guild_options = []
        self.log_label = None
        self.log_read_text = None

        self.create_widgets()

    def create_bot_list_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)
        ttk.Label(frame, text="mBot").pack(anchor='w')

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.bot_name_list = tk.Listbox(frame, selectmode=tk.SINGLE, height=15, yscrollcommand=scrollbar.set, exportselection=False)
        self.bot_name_list.pack(side=tk.LEFT, fill='both', expand=True)

        scrollbar.config(command=self.bot_name_list.yview)

    def create_button_list_left_frame(self, parent_frame):
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill='x', pady=2)

        ttk.Button(button_frame, text="Refresh mBots", command=self.refresh_mbot_list).pack(fill='x', pady=1)

    def create_left_frame(self, x, y, width, height):
        frame = ttk.Frame(self)
        frame.place(x=x, y=y, width=width, height=height)

        self.create_bot_list_frame(frame)
        self.create_button_list_left_frame(frame)

    def create_guild_list_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)
        ttk.Label(frame, text="Guild filter").pack(anchor='w')

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.guild_name_list = tk.Listbox(frame, selectmode=tk.SINGLE, height=10, yscrollcommand=scrollbar.set, exportselection=False)
        self.guild_name_list.pack(side=tk.LEFT, fill='both', expand=True)

        scrollbar.config(command=self.guild_name_list.yview)

    def create_button_list_center_frame(self, parent_frame):
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill='x', pady=2)

        ttk.Button(button_frame, text="Refresh Guild", command=self.refresh_guild_list).pack(fill='x', pady=1)
        self.guild_combo = ttk.Combobox(button_frame, textvariable=self.guild_combo_var, values=self.guild_options, state="readonly")
        self.guild_combo.pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Add Guild", command=self.add_guild).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Remove Guild", command=self.remove_guild).pack(fill='x', pady=1)

    def create_center_frame(self, x, y, width, height):
        frame = ttk.Frame(self)
        frame.place(x=x, y=y, width=width, height=height)

        self.create_guild_list_frame(frame)
        self.create_button_list_center_frame(frame)

    def create_log_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)

        self.log_label = ttk.Label(frame, text="Result")
        self.log_label.pack(anchor='w')

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_read_text = tk.Text(frame, height=15, state="disabled")
        self.log_read_text.pack(anchor='w')
        self.log_read_text.config(state="normal")

        scrollbar.config(command=self.log_read_text.yview)

    def create_button_list_right_frame(self, parent_frame):
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill='x', pady=2)

        ttk.Button(button_frame, text="Refresh result", command=self.refresh_result).pack(fill='x', pady=1)

    def create_right_frame(self, x, y, width, height):
        frame = ttk.Frame(self)
        frame.place(x=x, y=y, width=width, height=height)

        self.create_log_frame(frame)
        self.create_button_list_right_frame(frame)

    def create_widgets(self):
        left_frame_width = WINDOW_WIDTH * 2 / 7
        left_frame_height = WINDOW_HEIGHT
        left_frame_x = TOP_LEFT_X
        left_frame_y = TOP_LEFT_Y
        self.create_left_frame(x=left_frame_x, y=left_frame_y, width=left_frame_width, height=left_frame_height)

        center_frame_width = WINDOW_WIDTH * 3 / 14
        center_frame_height = WINDOW_HEIGHT
        center_frame_x = left_frame_width
        center_frame_y = TOP_LEFT_Y
        self.create_center_frame(x=center_frame_x, y=center_frame_y, width=center_frame_width, height=center_frame_height)

        right_frame_width = WINDOW_WIDTH - center_frame_width - left_frame_width
        right_frame_height = WINDOW_HEIGHT
        right_frame_x = WINDOW_WIDTH - right_frame_width
        right_frame_y = TOP_LEFT_Y
        self.create_right_frame(x=right_frame_x, y=right_frame_y, width=right_frame_width, height=right_frame_height)

        self.refresh_mbot_list()

    def refresh_mbot_list(self):
        mbot_list = findwindows.find_elements(class_name="#32770", visible_only=False, title_re=".*mBot.*")
        if mbot_list:
            self.mbot_list.clear()
            self.bot_name_list.delete(0, tk.END)
            sorted_mbot_list = sorted(mbot_list, key=lambda e: e.name)
            for mbot in sorted_mbot_list:
                self.mbot_list.append(MBot(mbot))
                self.bot_name_list.insert(tk.END, mbot.name)
            sorted_mbot_list.clear()
            mbot_list.clear()

    def refresh_guild_list(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        mbot_list[0].set_spy_player_checkbox_state()
        mbot_list[0].refresh_spy()
        content = mbot_list[0].get_list_of_guild()
        self.guild_options = list(content)
        if len(self.guild_options) != 0:
            self.guild_combo["values"] = self.guild_options
            self.guild_combo.current(0)

    def add_guild(self):
        if self.guild_combo_var.get() != "":
            name = self.guild_combo_var.get()
            if name not in self.guild_name_list.get(0, tk.END):
                self.guild_name_list.insert(tk.END, name)

    def remove_guild(self):
        selection = self.guild_name_list.curselection()
        if selection:
            self.guild_name_list.delete(selection[0])

    def refresh_result(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        mbot_list[0].set_spy_player_checkbox_state()
        mbot_list[0].refresh_spy()
        list_of_guild = self.guild_name_list.get(0, tk.END)
        content = mbot_list[0].get_list_of_member_in_guild(list_of_guild)
        if content is not None:
            self.log_label.configure(text=f"Result - {mbot_list[0].get_name()}")
            self.log_read_text.delete("1.0", "end")
            for line in content:
                self.log_read_text.insert("end", line)
                self.log_read_text.insert("end", "\n")
            total = len(content)
            self.log_read_text.insert("end", f"Total: {total}")
            self.log_read_text.see("end")

if __name__ == "__main__":
    app = FilterGuild()
    app.mainloop()
