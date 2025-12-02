import tkinter as tk
import win32gui
import win32con
from tkinter import ttk, messagebox
from pywinauto import findwindows
import time
import ctypes

SW_HIDE = 0
SW_SHOW = 5

class MBotClient():
    def __init__(self, mbot):
        self.mbot = mbot
        self.client_section = None
        self.delay_section = None
        self.save_settings_section = None
        self.log_off_section = None
        self.start_client_section = None
        self.kill_client_section = None
        self.reset_section = None
        self.stats_section = None
        self.hp_value = None
        self.mp_value = None
        self.name = ""

    def __str__(self):
        return f"name: {self.mbot.name}"

    def find_element_by_name(self, name):
        children = self.mbot.children()
        for child in children:
            if child.name == name:
                return child

    def find_element_by_next_element(self, name):
        children = self.mbot.children()
        for i, child in enumerate(children):
            next_child = children[i + 1] if i + 1 < len(children) else None
            if next_child.name == name:
                return child

    def find_nth_element_by_name(self, name, index):
        children = self.mbot.children()
        for i, child in enumerate(children):
            nth_child = children[i + index] if i + index < len(children) else None
            if child.name == name:
                return nth_child

    def extract_progress_bar(self, num_string):
        parts = num_string.split("/")
        current = int(parts[0].replace(",", "").strip())
        total = int(parts[1].replace(",", "").strip())
        return current * 100 / total

    def get_hp(self):
        if self.hp_value == None:
            self.hp_value = self.find_nth_element_by_name("HP", 6)
        return self.extract_progress_bar(self.hp_value.name)

    def get_mp(self):
        if self.mp_value == None:
            self.mp_value = self.find_nth_element_by_name("MP", 6)
        return self.extract_progress_bar(self.mp_value.name)

    def get_name(self):
        if self.name == "":
            self.name = self.find_nth_element_by_name("Hide client after relogin", 1)
        return self.name.name
    
    def get_stats(self):
        if self.stats_section == None:
            self.stats_section = self.find_nth_element_by_name("Stop training", 2)
        return self.stats_section.name
    
    def get_kills_per_hour(self):
        text = self.get_stats()
        sections = text.split("\n\n")  # Each block separated by empty line
        for sec in sections:
            if sec.startswith("Per hour"):
                per_hour_section = sec
                break

        for line in per_hour_section.splitlines():
            if line.startswith("Kills:"):
                kills_per_hour = line.split(":")[1].strip()
                break
        return f"Kills/hour: {kills_per_hour}"


    def set_delay(self, time_delay = 9999):
        if self.delay_section == None:
            self.delay_section = self.find_element_by_next_element("minutes before relogin")
        win32gui.SendMessage(self.delay_section.handle, win32con.WM_SETTEXT, 0, "")
        time.sleep(0.1)
        win32gui.SendMessage(self.delay_section.handle, win32con.WM_SETTEXT, 0, str(time_delay))
        time.sleep(0.1)

    def save_settings(self):
        if self.save_settings_section == None:
            self.save_settings_section = self.find_element_by_name("Save settings")
        win32gui.SendMessage(self.save_settings_section.handle, win32con.BM_CLICK, 0, 0)
        time.sleep(0.1)

    def log_off(self):
        if self.log_off_section == None:
            self.log_off_section = self.find_element_by_name("Log Off")
        win32gui.PostMessage(self.log_off_section.handle, win32con.BM_CLICK, 0, 0)
        time.sleep(0.1)
        confirmation_list = findwindows.find_elements(class_name="#32770", title="Confirmation")
        for confirmation in confirmation_list:
            children = confirmation.children()
            for child in children:
                if child.name == "&Yes":
                    win32gui.SendMessage(child.handle, win32con.BM_CLICK, 0, 0)
                    time.sleep(0.1)

    def start_client(self):
        if self.start_client_section == None:
            self.start_client_section = self.find_element_by_name("Start Client")
        win32gui.PostMessage(self.start_client_section.handle, win32con.BM_CLICK, 0, 0)
        time.sleep(0.1)

    def kill_client(self):
        if self.kill_client_section == None:
            self.kill_client_section = self.find_element_by_name("Kill Client")
        win32gui.PostMessage(self.kill_client_section.handle, win32con.BM_CLICK, 0, 0)
        time.sleep(0.1)
        confirmation_list = findwindows.find_elements(class_name="#32770", title="Confirmation")
        for confirmation in confirmation_list:
            children = confirmation.children()
            for child in children:
                if child.name == "&Yes":
                    win32gui.SendMessage(child.handle, win32con.BM_CLICK, 0, 0)
                    time.sleep(0.1)

    def show_client(self):
        ctypes.windll.user32.ShowWindow(self.mbot.handle, SW_SHOW)

    def hide_client(self):
        ctypes.windll.user32.ShowWindow(self.mbot.handle, SW_HIDE)

    def reset_client(self):
        if self.reset_section == None:
            self.reset_section = self.find_element_by_name("Reset")
        win32gui.PostMessage(self.reset_section.handle, win32con.BM_CLICK, 0, 0)
        time.sleep(0.1)


class ItemHpMpRow:
    def __init__(self, parent, row):
        self.frame = parent
        # Create widgets
        self.name_label = ttk.Label(self.frame, text="Name: X")
        self.kill_label = ttk.Label(self.frame, text="")
        self.hp_label = ttk.Label(self.frame, text="HP:")
        self.mp_label = ttk.Label(self.frame, text="  MP:")
        self.hp_value = ttk.Progressbar(self.frame, mode="determinate", length=100)
        self.mp_value = ttk.Progressbar(self.frame, mode="determinate", length=100)

        # Place widgets using grid
        self.hp_label.grid(row=row, column=0, padx=0, pady=2, sticky="w")
        self.hp_value.grid(row=row, column=1, padx=0, pady=2, sticky="w")
        self.mp_label.grid(row=row, column=2, padx=0, pady=2, sticky="w")
        self.mp_value.grid(row=row, column=3, padx=0, pady=2, sticky="w")
        self.name_label.grid(row=row, column=4, padx=5, pady=2, sticky="w")
        self.kill_label.grid(row=row, column=5, padx=5, pady=2, sticky="w")

    def destroy(self):
        self.hp_label.destroy()
        self.hp_value.destroy()
        self.mp_label.destroy()
        self.mp_value.destroy()
        self.name_label.destroy()
        self.kill_label.destroy()

class MBotManager(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Quản Lý Mbot")
        self.geometry("700x560")
        self.resizable(False, False)
        self.mbot_list = []
        self.hp_mp_list = []
        self.right_frame = None

        self.create_widgets()

    def create_widgets(self):
        left_frame = ttk.Frame(self)
        left_frame.place(x=10, y=10, width=200, height=280)

        ttk.Label(left_frame, text="Danh Sách mBots").pack(anchor='w')

        self.bot_name_list = tk.Listbox(left_frame, selectmode=tk.MULTIPLE, height=5)
        self.bot_name_list.pack(fill='both', expand=True)

        self.right_frame = ttk.Frame(self)
        self.right_frame.place(x=220, y=20, width=500, height=520)
        self.right_frame.grid_columnconfigure(0, minsize=10)
        self.right_frame.grid_columnconfigure(1, minsize=100)
        self.right_frame.grid_columnconfigure(2, minsize=10)
        self.right_frame.grid_columnconfigure(3, minsize=100)
        self.right_frame.grid_columnconfigure(4, minsize=100)
        self.right_frame.grid_columnconfigure(5, minsize=100)

        self.refresh_list()

        bot_frame = ttk.Frame(self)
        bot_frame.place(x=10, y=290, width=200, height=260)

        ttk.Button(bot_frame, text="Làm mới danh sách", command=self.refresh_list).pack(fill='x', pady=2)
        ttk.Button(bot_frame, text="Chọn tất cả mbot hiện có", command=self.select_all).pack(fill='x', pady=2)
        ttk.Button(bot_frame, text="Bỏ chọn tất cả mbot hiện có", command=self.deselect_all).pack(fill='x', pady=2)
        ttk.Button(bot_frame, text="Start client mbot đang chọn", command=self.start_client_selected).pack(fill='x', pady=2)
        ttk.Button(bot_frame, text="Kill client mbot đang chọn", command=self.kill_client_selected).pack(fill='x', pady=2)
        ttk.Button(bot_frame, text="Log Off mbot đang chọn", command=self.log_off_selected).pack(fill='x', pady=2)
        ttk.Button(bot_frame, text="Hiện mbot đang chọn", command=self.show_selected).pack(fill='x', pady=2)
        ttk.Button(bot_frame, text="Ẩn mbot đang chọn", command=self.hide_selected).pack(fill='x', pady=2)
        ttk.Button(bot_frame, text="Reset mbot đang chọn", command=self.reset_selected).pack(fill='x', pady=2)

    def refresh_list(self):
        mbot_list = findwindows.find_elements(class_name="#32770", visible_only=False)
        if mbot_list:
            self.mbot_list.clear()
            self.bot_name_list.delete(0, tk.END)
            sorted_mbot_list = sorted(mbot_list, key=lambda e: e.name)
            for mbot in sorted_mbot_list:
                self.mbot_list.append(MBotClient(mbot))
                self.bot_name_list.insert(tk.END, mbot.name)
            sorted_mbot_list.clear()
            mbot_list.clear()

            for item in self.hp_mp_list:
                item.destroy()
            self.hp_mp_list.clear()

    def update_hp_mp(self):
        size_of_list = self.bot_name_list.size()
        if not self.hp_mp_list:
            for i in range(size_of_list):
                self.hp_mp_list.append(ItemHpMpRow(self.right_frame, i))

        for i in range(size_of_list):
            self.hp_mp_list[i].name_label.configure(text=self.mbot_list[i].get_name())
            self.hp_mp_list[i].hp_value['value'] = self.mbot_list[i].get_hp()
            self.hp_mp_list[i].mp_value['value'] = self.mbot_list[i].get_mp()
            self.hp_mp_list[i].kill_label.configure(text=self.mbot_list[i].get_kills_per_hour())
        app.after(3000, self.update_hp_mp)

    def select_all(self):
        self.bot_name_list.select_set(0, tk.END)

    def deselect_all(self):
        self.bot_name_list.select_clear(0, tk.END)

    def start_client_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một mbot để đóng.")
            return
        mbot_list = [self.mbot_list[i] for i in selected]
        for index, mbot in enumerate(mbot_list):
            mbot.set_delay(index%4 + 2)
            mbot.save_settings()
            mbot.start_client()
            time.sleep(15)

    def kill_client_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một mbot để đóng.")
            return
        mbot_list = [self.mbot_list[i] for i in selected]
        for mbot in mbot_list:
            mbot.kill_client()
            time.sleep(0.1)

    def log_off_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một mbot để đóng.")
            return
        mbot_list = [self.mbot_list[i] for i in selected]
        for mbot in mbot_list:
            mbot.set_delay()
            mbot.save_settings()
            mbot.log_off()
            time.sleep(1)

    def show_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một mbot để đóng.")
            return
        mbot_list = [self.mbot_list[i] for i in selected]
        for mbot in mbot_list:
            mbot.show_client()
            time.sleep(0.1)

    def hide_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một mbot để đóng.")
            return
        mbot_list = [self.mbot_list[i] for i in selected]
        for mbot in mbot_list:
            mbot.hide_client()
            time.sleep(0.1)

    def reset_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một mbot để đóng.")
            return
        mbot_list = [self.mbot_list[i] for i in selected]
        for mbot in mbot_list:
            mbot.reset_client()
            time.sleep(0.1)

if __name__ == "__main__":
    app = MBotManager()
    app.update_hp_mp()
    app.mainloop()
