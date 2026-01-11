import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import win32gui
import win32con
import win32com.client
from pywinauto import findwindows
import ctypes
import os
import base64
import json
import subprocess
import uiautomation as auto
import re
from collections import defaultdict
from pywinauto.timings import always_wait_until, TimeoutError

FILE_NAME = "accounts.json"
accounts = None

WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 650
TOP_LEFT_X = 0
TOP_LEFT_Y = 0
CHAT_BUTTON_TEXTS = ["Allchat", "PM", "Party", "Guild", "Global", "Academy", "GM", "Union", "Unique"]
INVENTORY_OPTIONS = ["Avatar", "Fellow", "Guildstorage", "Inventory", "Pet", "Storage"]
LOG_OPTIONS = ["Refresh", "Pieces", "Log", "Active buffs"]

tcvn3_to_unicode = {
    # Lowercase a
    'µ':'à', '¸':'á', '¶':'ả', '·':'ã', '¹':'ạ',
    '¨':'ă', '»':'ằ', '¾':'ắ', '¼':'ẳ', '½':'ẵ', 'Æ':'ặ',
    '©':'â', 'Ç':'ầ', 'Ê':'ấ', 'È':'ẩ', 'É':'ẫ', 'Ë':'ậ',

    # Lowercase d
    '®':'đ',

    # Lowercase e
    'Ì':'è', 'Ð':'é', 'Î':'ẻ', 'Ï':'ẽ', 'Ñ':'ẹ',
    'ª':'ê', 'Ò':'ề', 'Õ':'ế', 'Ó':'ể', 'Ô':'ễ', 'Ö':'ệ',

    # Lowercase i
    '×':'ì', 'Ý':'í', 'Ø':'ỉ', 'Ü':'ĩ', 'Þ':'ị',

    # Lowercase o
    'ß':'ò', 'ã':'ó', 'á':'ỏ', 'â':'õ', 'ä':'ọ',
    '«':'ô', 'å':'ồ', 'è':'ố', 'æ':'ổ', 'ç':'ỗ', 'é':'ộ',
    '¬':'ơ', 'ê':'ờ', 'í':'ớ', 'ë':'ở', 'ì':'ỡ', 'î':'ợ',

    # Lowercase u
    'ï':'ù', 'ó':'ú', 'ñ':'ủ', 'ò':'ũ', 'ô':'ụ',
    '­':'ư', 'õ':'ừ', 'ø':'ứ', 'ö':'ử', '÷':'ữ', 'ù':'ự',

    # Lowercase y
    'ú':'ỳ', 'ý':'ý', 'û':'ỷ', 'ü':'ỹ', 'þ':'ỵ',

    # Uppercase D
    "§": "Đ",

    # Uppercase E
    "£": "Ê",

    # Upper case O
    "¤": "Ô", "¥": "Ơ",

    # Upper case U
    "¦": "Ư"
}

def tcvn3_to_unicode_text(text):
    """Convert TCVN3 text to Unicode."""
    return ''.join(tcvn3_to_unicode.get(char, char) for char in text)

def extract_progress_bar(num_string):
    parts = num_string.split("/")
    current = int(parts[0].replace(",", "").strip())
    total = int(parts[1].replace(",", "").strip())
    return current * 100 / total

def load_accounts():
    if os.path.exists(FILE_NAME):
        with open(FILE_NAME, "r") as f:
            return json.load(f)
    return []

def save_accounts():
    with open(FILE_NAME, "w") as f:
        json.dump(accounts, f, indent=4)

def click_confirmation(class_name="#32770", title="Confirmation", text="&Yes", is_re=False, timeout=1, retry_interval=0.1):
    """
    Waits for a confirmation dialog and clicks the text button if found.

    Args:
        class: Class to match
        title (str): Dialog title to match.
        text (str): Text of the button (default "&Yes" for standard Windows dialogs).
        is_re(bool): True if we want to search regular expression
        timeout (float): Max time to wait for the dialog.
        retry_interval (float): How often to check for the dialog.

    Returns:
        bool: True if Yes button was clicked, False if timed out or not found.
    """
    try:
        @always_wait_until(timeout, retry_interval)
        def wait_and_click():
            if is_re:
                confirmation_list = findwindows.find_elements(class_name=class_name, title_re=title)
            else:
                confirmation_list = findwindows.find_elements(class_name=class_name, title=title)
            for confirmation in confirmation_list:
                children = confirmation.children()
                for child in children:
                    if child.name == text:
                        win32gui.SendMessage(child.handle, win32con.BM_CLICK, 0, 0)
                        return True
            return False

        return wait_and_click()
    except TimeoutError:
        return False

class MBot():
    def __init__(self, mbot):
        self.mbot = mbot
        self.delay_edit = None
        self.save_settings_button = None
        self.log_off_button = None
        self.start_client_button = None
        self.kill_client_button = None
        self.show_hide_client_button = None
        self.reset_section = None
        self.stats_section = None
        self.hp_value = None
        self.mp_value = None
        self.name = ""
        self.current_position_button = None
        self.start_training_button = None
        self.stop_training_button = None
        self.inventory_combo = None
        self.inventory_refresh_button = None
        self.inventory_items = None
        self.clear_button = None
        self.log_edit = None
        self.drops_checkbox_state = None
        self.who_attacked_you_checkbox_state = None
        self.spy_player_checkbox_state = None
        self.refresh_spy_button = None
        self.spy_combo = None
        self.spy_log = None
        self.chat_buttons_dict = {}

    def __str__(self):
        return f"name: {self.mbot.name}"

    def is_valid(self):
        return win32gui.IsWindow(self.mbot.handle)

    def find_element_by_name(self, name):
        if not self.is_valid():
            return

        children = self.mbot.children()
        for child in children:
            if child.name == name:
                return child

    def find_element_by_next_element(self, name):
        if not self.is_valid():
            return

        children = self.mbot.children()
        for i, child in enumerate(children):
            next_child = children[i + 1] if i + 1 < len(children) else None
            if next_child.name == name:
                return child

    def find_nth_element_by_name(self, name, index):
        if not self.is_valid():
            return

        children = self.mbot.children()
        for i, child in enumerate(children):
            nth_child = children[i + index] if i + index < len(children) else None
            if child.name == name:
                return nth_child

    def get_hp(self):
        if not self.is_valid():
            return

        if not self.hp_value:
            self.hp_value = self.find_nth_element_by_name("HP", 6)

        return extract_progress_bar(self.hp_value.name)

    def get_mp(self):
        if not self.is_valid():
            return

        if not self.mp_value:
            self.mp_value = self.find_nth_element_by_name("MP", 6)

        return extract_progress_bar(self.mp_value.name)

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

    def get_stats(self):
        if not self.is_valid():
            return

        if not self.stats_section:
            self.stats_section = self.find_nth_element_by_name("Stop training", 2)

        return self.stats_section.name

    def get_kills_per_hour(self):
        if not self.is_valid():
            return

        text = self.get_stats()
        sections = text.split("\n\n")
        for sec in sections:
            if sec.startswith("Per hour"):
                per_hour_section = sec
                break

        for line in per_hour_section.splitlines():
            if line.startswith("Kills:"):
                kills_per_hour = line.split(":")[1].strip()
                kills_per_hour = kills_per_hour.split(".")[0].strip()
                break
        return f"K/H: {kills_per_hour}"

    def get_edit_content(self, handle):
        if not self.is_valid():
            return

        length = win32gui.SendMessage(handle, win32con.WM_GETTEXTLENGTH, 0, 0)
        buffer = ctypes.create_unicode_buffer(length + 1)
        win32gui.SendMessage(handle, win32con.WM_GETTEXT, length + 1, buffer)

        full_text = buffer.value
        lines = full_text.splitlines()
        last_100 = "\n".join(lines[-100:])
        content = tcvn3_to_unicode_text(last_100)
        return content

    def get_chat_content(self, button_name):
        if not self.is_valid():
            return

        keyword = "Use colored chat"
        if button_name not in self.chat_buttons_dict:
            self.chat_buttons_dict[button_name] = self.find_nth_element_by_name(keyword, CHAT_BUTTON_TEXTS.index(button_name)+1)
        return self.get_edit_content(self.chat_buttons_dict[button_name].handle)

    def set_delay(self, is_default = True):
        if not self.is_valid():
            return

        if not self.delay_edit:
            self.delay_edit = self.find_element_by_next_element("minutes before relogin")

        delay = 999
        if not is_default:
            name = self.get_name()
            for acc in accounts:
                if acc["character"] == name:
                    delay = acc["delay_time"]
                    break

        win32gui.SendMessage(self.delay_edit.handle, win32con.WM_SETTEXT, 0, "")
        win32gui.SendMessage(self.delay_edit.handle, win32con.WM_SETTEXT, 0, str(delay))

    def save_settings(self):
        if not self.is_valid():
            return

        if not self.save_settings_button:
            self.save_settings_button = self.find_element_by_name("Save settings")

        win32gui.SendMessage(self.save_settings_button.handle, win32con.BM_CLICK, 0, 0)

    def log_off(self):
        if not self.is_valid():
            return

        if not self.log_off_button:
            self.log_off_button = self.find_element_by_name("Log Off")

        win32gui.PostMessage(self.log_off_button.handle, win32con.BM_CLICK, 0, 0)
        click_confirmation()

    def start_client(self):
        if not self.is_valid():
            return

        if not self.start_client_button:
            self.start_client_button = self.find_element_by_name("Start Client")

        print("start_client")
        win32gui.PostMessage(self.start_client_button.handle, win32con.BM_CLICK, 0, 0)

    def kill_client(self):
        if not self.is_valid():
            return

        if not self.kill_client_button:
            self.kill_client_button = self.find_element_by_name("Kill Client")

        win32gui.PostMessage(self.kill_client_button.handle, win32con.BM_CLICK, 0, 0)
        click_confirmation()

    def kill_mbot(self):
        if not self.is_valid():
            return

        win32gui.PostMessage(self.mbot.handle, win32con.WM_CLOSE, 0, 0)
        click_confirmation()

    def show_hide_mbot(self):
        if not self.is_valid():
            return

        if win32gui.IsWindowVisible(self.mbot.handle):
            ctypes.windll.user32.ShowWindow(self.mbot.handle, 0)
        else:
            ctypes.windll.user32.ShowWindow(self.mbot.handle, 5)

    def show_hide_client(self):
        if not self.is_valid():
            return

        if not self.show_hide_client_button:
            self.show_hide_client_button = self.find_element_by_name("Show / Hide Client")

        win32gui.PostMessage(self.show_hide_client_button.handle, win32con.BM_CLICK, 0, 0)

    def reset_mbot(self):
        if not self.is_valid():
            return

        if not self.reset_section:
            self.reset_section = self.find_element_by_name("Reset")

        win32gui.PostMessage(self.reset_section.handle, win32con.BM_CLICK, 0, 0)

    def get_current_position(self):
        if not self.is_valid():
            return

        if not self.current_position_button:
            self.current_position_button = self.find_element_by_name("Get current position")

        win32gui.PostMessage(self.current_position_button.handle, win32con.BM_CLICK, 0, 0)

    def start_training(self):
        if not self.is_valid():
            return

        if not self.start_training_button:
            self.start_training_button = self.find_element_by_name("Start training")

        win32gui.PostMessage(self.start_training_button.handle, win32con.BM_CLICK, 0, 0)

    def stop_training(self):
        if not self.is_valid():
            return

        if not self.stop_training_button:
            self.stop_training_button = self.find_element_by_name("Stop training")

        win32gui.PostMessage(self.stop_training_button.handle, win32con.BM_CLICK, 0, 0)

    def set_inventory_combo(self, index):
        if not self.is_valid():
            return

        if not self.inventory_combo:
            self.inventory_combo = self.find_nth_element_by_name("Inventory", 1)

        win32gui.SendMessage(self.inventory_combo.handle, win32con.CB_SETCURSEL, index, 0)

    def refresh_inventory(self):
        if not self.is_valid():
            return

        if not self.inventory_refresh_button:
            self.inventory_refresh_button = self.find_nth_element_by_name("Inventory", 2)

        win32gui.PostMessage(self.inventory_refresh_button.handle, win32con.BM_CLICK, 0, 0)

    def get_inventory_items(self):
        if not self.is_valid():
            return

        if not self.inventory_items:
            self.inventory_items = self.find_nth_element_by_name("Inventory", 3)

        count = win32gui.SendMessage(self.inventory_items.handle, win32con.LB_GETCOUNT, 0, 0)
        if count <= 0:
            return []

        items = []
        for i in range(count):
            length = win32gui.SendMessage(self.inventory_items.handle, win32con.LB_GETTEXTLEN, i, 0)
            if length <= 0:
                continue

            buffer = ctypes.create_unicode_buffer(length + 1)
            win32gui.SendMessage(self.inventory_items.handle, win32con.LB_GETTEXT, i, buffer)

            text = buffer.value
            text_convert = tcvn3_to_unicode_text(text)
            items.append(text_convert)
        return items

    def get_log(self):
        if not self.is_valid():
            return

        if not self.log_edit:
            self.log_edit = self.find_nth_element_by_name("Weaponswitch", 1)

        return self.get_edit_content(self.log_edit.handle)

    def clear_log(self):
        if not self.is_valid():
            return

        if not self.clear_button:
            self.clear_button = self.find_element_by_name("Clear")

        win32gui.PostMessage(self.clear_button.handle, win32con.BM_CLICK, 0, 0)

    def get_drops_checkbox_state(self):
        if not self.is_valid():
            return

        if not self.drops_checkbox_state:
            self.drops_checkbox_state = self.find_element_by_name("Drops")

        state = win32gui.SendMessage(self.drops_checkbox_state.handle, win32con.BM_GETCHECK, 0, 0)
        return state == win32con.BST_CHECKED

    def set_drops_checkbox_state(self, desired_state):
        if not self.is_valid():
            return

        if not self.drops_checkbox_state:
            self.drops_checkbox_state = self.find_element_by_name("Drops")

        current = self.get_drops_checkbox_state()
        if current != desired_state:
            win32gui.PostMessage(self.drops_checkbox_state.handle, win32con.BM_CLICK, 0, 0)

    def get_who_attacked_you_checkbox_state(self):
        if not self.is_valid():
            return

        if not self.who_attacked_you_checkbox_state:
            self.who_attacked_you_checkbox_state = self.find_element_by_name("Players who attacked you")

        state = win32gui.SendMessage(self.who_attacked_you_checkbox_state.handle, win32con.BM_GETCHECK, 0, 0)
        return state == win32con.BST_CHECKED

    def set_who_attacked_you_checkbox_state(self, desired_state):
        if not self.is_valid():
            return

        if not self.who_attacked_you_checkbox_state:
            self.who_attacked_you_checkbox_state = self.find_element_by_name("Players who attacked you")

        current = self.get_who_attacked_you_checkbox_state()
        if current != desired_state:
            win32gui.PostMessage(self.who_attacked_you_checkbox_state.handle, win32con.BM_CLICK, 0, 0)

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

    def get_active_buffs(self):
        result = []
        is_found = False
        start_collecting = False
        if not self.is_valid():
            return

        if not self.spy_combo:
            self.spy_combo = self.find_nth_element_by_name("Spy", 10)

        if not self.spy_log:
            self.spy_log = self.find_nth_element_by_name("Spy", 11)

        count = win32gui.SendMessage(self.spy_combo.handle, win32con.CB_GETCOUNT, 0, 0)
        pattern = re.compile(rf"^Name:\s+{self.get_name()}$")

        for i in range(count):
            win32gui.SendMessage(self.spy_combo.handle, win32con.WM_KEYDOWN, win32con.VK_DOWN, 0)
            content = self.get_edit_content(self.spy_log.handle)
            lines = content.splitlines()
            for line in lines:
                if pattern.search(line):
                    is_found = True
                    continue

                if is_found:
                    if line.startswith("Active buffs:"):
                        start_collecting = True
                        continue

                    if start_collecting:
                        line = line.lstrip("\t")
                        result.append(line)
                
            if is_found:
                return result


class ItemHpMpRow:
    def __init__(self, parent, row):
        self.frame = parent
        self.name_label = ttk.Label(self.frame, text="Name: X")
        self.kill_label = ttk.Label(self.frame, text="")
        self.hp_label = ttk.Label(self.frame, text="HP:")
        self.mp_label = ttk.Label(self.frame, text="  MP:")
        self.hp_value = ttk.Progressbar(self.frame, mode="determinate", length=100)
        self.mp_value = ttk.Progressbar(self.frame, mode="determinate", length=100)

        self.name_label.grid(row=row, column=0, padx=5, pady=2, sticky="w")
        self.hp_label.grid(row=row, column=1, padx=0, pady=2, sticky="w")
        self.hp_value.grid(row=row, column=2, padx=0, pady=2, sticky="w")
        self.mp_label.grid(row=row, column=3, padx=0, pady=2, sticky="w")
        self.mp_value.grid(row=row, column=4, padx=0, pady=2, sticky="w")
        self.kill_label.grid(row=row, column=5, padx=5, pady=2, sticky="w")

    def destroy(self):
        self.hp_label.destroy()
        self.hp_value.destroy()
        self.mp_label.destroy()
        self.mp_value.destroy()
        self.name_label.destroy()
        self.kill_label.destroy()

class Monitor(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.bot_name_list = None
        self.mbot_list = []
        self.update_hp_mp_job = None
        self.hp_mp_list = []
        self.hp_mp_frame = None
        self.update_chat_job = None
        self.chat_label = None
        self.chat_read_text = None
        self.log_label = None
        self.log_read_text = None
        self.inventory_combo = None
        self.inventory_options = INVENTORY_OPTIONS
        self.current_log = LOG_OPTIONS[2]
        self.inventory_combo_var = tk.StringVar()
        self.drops_check_var = tk.StringVar()
        self.who_attacked_you_check_var = tk.StringVar()

        self.create_widgets()
        self.update_hp_mp()

    def process_mbots(self, mbot_list, actions):
        def process_next(index=0):
            if index >= len(mbot_list):
                return

            mbot = mbot_list[index]
            current_time = 0
            for delay_ms, method in actions:
                current_time += delay_ms
                self.after(current_time, lambda m=mbot, meth=method: meth(m))

            total_delay = sum(delay for delay, _ in actions)
            self.after(total_delay + 100, lambda: process_next(index + 1))

        process_next(0)

    def create_bot_list_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)
        ttk.Label(frame, text="mBot").pack(anchor='w')

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.bot_name_list = tk.Listbox(frame, selectmode=tk.EXTENDED, height=15, yscrollcommand=scrollbar.set, exportselection=False)
        self.bot_name_list.pack(side=tk.LEFT, fill='both', expand=True)
        self.bot_name_list.bind("<<ListboxSelect>>", self.show_mbot_log)

        scrollbar.config(command=self.bot_name_list.yview)

    def show_mbot_log(self, event):
        selection = self.bot_name_list.curselection()
        if selection:
            if self.current_log == LOG_OPTIONS[0]:
                self.update_inventory_log()
            elif self.current_log == LOG_OPTIONS[1]:
                self.update_inventory_pieces_log()
            elif self.current_log == LOG_OPTIONS[2]:
                self.update_log()
            elif self.current_log == LOG_OPTIONS[3]:
                self.update_active_buffs()
            self.update_drops_state()
            self.update_who_attacked_you_state()

    def create_button_list_frame(self, parent_frame):
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill='x', pady=2)

        ttk.Button(button_frame, text="Refresh mBots", command=self.refresh_list).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Select all mBots", command=self.select_all).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Deselect all mBots", command=self.deselect_all).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Start client selected mBots", command=self.start_client_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Kill client selected mBots", command=self.kill_client_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Kill mBot selected mBots", command=self.kill_mbot_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Log Off selected mBots", command=self.log_off_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Show/Hide selected mBots", command=self.show_hide_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Show/Hide Client selected mBots", command=self.show_hide_client_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Reset selected mBots", command=self.reset_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Get current position selected mBots", command=self.get_current_position_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Start training selected mBots", command=self.start_training_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Stop training selected mBots", command=self.stop_training_selected).pack(fill='x', pady=1)

    def create_left_frame(self, x, y, width, height):
        frame = ttk.Frame(self)
        frame.place(x=x, y=y, width=width, height=height)

        self.create_bot_list_frame(frame)
        self.create_button_list_frame(frame)

    def create_hp_mp_frame(self, parent_frame, width):
        self.hp_mp_frame = ttk.Frame(parent_frame)
        self.hp_mp_frame.pack(fill='x', pady=2)
        name_minsize = width / 5
        hp_mp_label_minsize = width / 25
        hp_mp_progress_bar_minsize = width / 5
        kill_per_hours_minsize = (width - hp_mp_label_minsize * 2 - hp_mp_progress_bar_minsize * 2) / 2

        self.hp_mp_frame.grid_columnconfigure(0, minsize=name_minsize)
        self.hp_mp_frame.grid_columnconfigure(1, minsize=hp_mp_label_minsize)
        self.hp_mp_frame.grid_columnconfigure(2, minsize=hp_mp_progress_bar_minsize)
        self.hp_mp_frame.grid_columnconfigure(3, minsize=hp_mp_label_minsize)
        self.hp_mp_frame.grid_columnconfigure(4, minsize=hp_mp_progress_bar_minsize)
        self.hp_mp_frame.grid_columnconfigure(5, minsize=kill_per_hours_minsize)

    def create_center_frame(self, x, y, width, height):
        frame = ttk.Frame(self)
        frame.place(x=x, y=y, width=width, height=height)

        ttk.Label(frame, text="Character").pack(anchor='w')
        self.create_hp_mp_frame(frame, width=width)

    def update_chat(self, mbot, button_name):
        if not mbot.is_valid():
            if self.update_chat_job:
                self.after_cancel(self.update_chat_job)
            return

        content = mbot.get_chat_content(button_name)
        if self.chat_read_text.get("1.0", "end-1c") != content:
            if content is not None:
                self.chat_label.configure(text=f"Chat - {mbot.get_name()}")
                self.chat_read_text.delete("1.0", "end")
                self.chat_read_text.insert("1.0", content)
                self.chat_read_text.see("end")
            else:
                self.chat_read_text.delete("1.0", "end")
                return

        self.update_chat_job = self.after(20000, lambda: self.update_chat(mbot, button_name))

    def start_update_chat(self, button_name):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        if self.update_chat_job:
            self.after_cancel(self.update_chat_job)

        self.update_chat(mbot_list[0], button_name)

    def create_chat_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)
        for i, text in enumerate(CHAT_BUTTON_TEXTS):
            btn = ttk.Button(frame, text=text, command=lambda t=text: self.start_update_chat(t))
            btn.grid(row=0, column=i, padx=1, pady=5)

        total_buttons = len(CHAT_BUTTON_TEXTS)
        for i in range(total_buttons):
            frame.grid_columnconfigure(i, weight=1)

        scrollbar = tk.Scrollbar(frame)
        scrollbar.grid(row=1, column=total_buttons, sticky="ns")

        self.chat_read_text = tk.Text(frame, height=15, state="disabled", yscrollcommand=scrollbar.set)
        self.chat_read_text.grid(row=1, column=0, columnspan=total_buttons, sticky="ew", padx=5, pady=5)
        self.chat_read_text.config(state="normal")

        scrollbar.config(command=self.chat_read_text.yview)

    def update_active_buffs(self):
        self.current_log = LOG_OPTIONS[3]
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        mbot_list[0].set_spy_player_checkbox_state()
        mbot_list[0].set_spy_player_checkbox_state()
        mbot_list[0].refresh_spy()
        content = mbot_list[0].get_active_buffs()
        if content is not None:
            self.log_label.configure(text=f"Inventory & Log - {mbot_list[0].get_name()}")
            self.log_read_text.delete("1.0", "end")
            for line in content:
                self.log_read_text.insert("end", line)
                self.log_read_text.insert("end", "\n")

    def update_inventory_pieces_log(self):
        self.current_log = LOG_OPTIONS[1]
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        selected_option = self.inventory_combo_var.get()
        index_option = self.inventory_options.index(selected_option)
        mbot_list[0].set_inventory_combo(index_option)
        mbot_list[0].refresh_inventory()
        content = mbot_list[0].get_inventory_items()
        if content is not None:
            self.log_label.configure(text=f"Inventory & Log - {mbot_list[0].get_name()}")
            item_totals = defaultdict(int)
            slot_item_totals = defaultdict(int)
            self.log_read_text.delete("1.0", "end")
            for line in content:
                match = re.search(r':\s*(.*?)\s*\((\d+)\s+pieces\)', line)
                if match:
                    item_name = match.group(1)
                    quantity = int(match.group(2))
                    item_totals[item_name] += quantity
                    slot_item_totals[item_name] += 1
            for item in sorted(item_totals):
                self.log_read_text.insert("end", f"{item}: {item_totals[item]} pieces - {slot_item_totals[item]} slots")
                self.log_read_text.insert("end", "\n")

    def update_inventory_log(self):
        self.current_log = LOG_OPTIONS[0]
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        selected_option = self.inventory_combo_var.get()
        index_option = self.inventory_options.index(selected_option)
        mbot_list[0].set_inventory_combo(index_option)
        mbot_list[0].refresh_inventory()
        content = mbot_list[0].get_inventory_items()
        if content is not None:
            self.log_label.configure(text=f"Inventory & Log - {mbot_list[0].get_name()}")
            self.log_read_text.delete("1.0", "end")
            for line in content:
                self.log_read_text.insert("end", line)
                self.log_read_text.insert("end", "\n")

    def update_log(self):
        self.current_log = LOG_OPTIONS[2]
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        content = mbot_list[0].get_log()
        if content is not None:
            self.log_label.configure(text=f"Inventory & Log - {mbot_list[0].get_name()}")
            self.log_read_text.delete("1.0", "end")
            self.log_read_text.insert("end", content)
            self.log_read_text.see("end")

    def clear_log(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.clear_log())
            ]
        )
        self.log_label.configure(text=f"Inventory & Log - {mbot_list[0].get_name()}")
        self.log_read_text.delete("1.0", "end")

    def update_drops_state(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        state  = mbot_list[0].get_drops_checkbox_state()
        self.drops_check_var.set(state)

    def update_who_attacked_you_state(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        state  = mbot_list[0].get_who_attacked_you_checkbox_state()
        self.who_attacked_you_check_var.set(state)

    def on_drops_check_change(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.set_drops_checkbox_state(self.drops_check_var.get())),
                (100, lambda mbot: mbot.save_settings())
            ]
        )

    def on_who_attacked_you_check_change(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.set_who_attacked_you_checkbox_state(self.who_attacked_you_check_var.get())),
                (100, lambda mbot: mbot.save_settings())
            ]
        )

    def create_inventory_log_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)
        columnspan_log = 15

        for i in range(columnspan_log):
            frame.grid_columnconfigure(i, weight=1)

        refresh_button = ttk.Button(frame, text="Refresh", command=self.update_inventory_log)
        refresh_button.grid(row=0, column=0, padx=0, pady=0)

        only_pieces_button = ttk.Button(frame, text="Only pieces", command=self.update_inventory_pieces_log)
        only_pieces_button.grid(row=0, column=1, padx=0, pady=0)

        self.inventory_combo = ttk.Combobox(frame, textvariable=self.inventory_combo_var, values=self.inventory_options, state="readonly")
        self.inventory_combo.current(3)
        self.inventory_combo.grid(row=0, column=2, padx=0, pady=0)

        active_buffs_button = ttk.Button(frame, text="Active buffs", command=self.update_active_buffs)
        active_buffs_button.grid(row=0, column=3, padx=0, pady=0)

        log_button = ttk.Button(frame, text="Log", command=self.update_log)
        log_button.grid(row=1, column=0, padx=0, pady=0)
        clear_button = ttk.Button(frame, text="Clear", command=self.clear_log)
        clear_button.grid(row=1, column=1, padx=0, pady=0)

        drops_checkbox_state = tk.Checkbutton(frame, text="Drops", variable=self.drops_check_var, command=self.on_drops_check_change)
        drops_checkbox_state.grid(row=1, column=13, padx=0, pady=0)

        who_attacked_you_checkbox_state = tk.Checkbutton(frame, text="Players who attacked you", variable=self.who_attacked_you_check_var, command=self.on_who_attacked_you_check_change)
        who_attacked_you_checkbox_state.grid(row=1, column=14, padx=0, pady=0)

        scrollbar = tk.Scrollbar(frame)
        scrollbar.grid(row=2, column=columnspan_log, sticky="ns")

        self.log_read_text = tk.Text(frame, height=14, state="disabled")
        self.log_read_text.grid(row=2, column=0, columnspan=columnspan_log, sticky="ew", padx=5, pady=5)
        self.log_read_text.config(state="normal")

        scrollbar.config(command=self.log_read_text.yview)

    def create_right_frame(self, x, y, width, height):
        frame = ttk.Frame(self)
        frame.place(x=x, y=y, width=width, height=height)

        self.chat_label = ttk.Label(frame, text="Chat")
        self.chat_label.pack(anchor='w')
        self.create_chat_frame(frame)
        self.log_label = ttk.Label(frame, text="Inventory & Log")
        self.log_label.pack(anchor='w')
        self.create_inventory_log_frame(frame)

    def create_widgets(self):
        left_frame_width = WINDOW_WIDTH * 2 / 12
        left_frame_height = WINDOW_HEIGHT
        left_frame_x = TOP_LEFT_X
        left_frame_y = TOP_LEFT_Y
        self.create_left_frame(x=left_frame_x, y=left_frame_y, width=left_frame_width, height=left_frame_height)

        center_frame_width = WINDOW_WIDTH * 4 / 12
        center_frame_height = WINDOW_HEIGHT
        center_frame_x = left_frame_width
        center_frame_y = TOP_LEFT_Y
        self.create_center_frame(x=center_frame_x, y=center_frame_y, width=center_frame_width, height=center_frame_height)

        right_frame_width = WINDOW_WIDTH - center_frame_width - left_frame_width
        right_frame_height = WINDOW_HEIGHT
        right_frame_x = WINDOW_WIDTH - right_frame_width
        right_frame_y = TOP_LEFT_Y
        self.create_right_frame(x=right_frame_x, y=right_frame_y, width=right_frame_width, height=right_frame_height)

        self.refresh_list()

    def refresh_list(self):
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

            if self.update_hp_mp_job:
                self.after_cancel(self.update_hp_mp_job)

            for item in self.hp_mp_list:
                item.destroy()
            self.hp_mp_list.clear()
            self.update_hp_mp()

    def update_hp_mp(self):
        size_of_list = self.bot_name_list.size()
        if len(self.hp_mp_list) != size_of_list:
            self.hp_mp_list.clear()
            for i in range(size_of_list):
                self.hp_mp_list.append(ItemHpMpRow(self.hp_mp_frame, i))

        for i in range(size_of_list):
            self.hp_mp_list[i].name_label.configure(text=self.mbot_list[i].get_name())
            self.hp_mp_list[i].hp_value['value'] = self.mbot_list[i].get_hp()
            self.hp_mp_list[i].mp_value['value'] = self.mbot_list[i].get_mp()
            self.hp_mp_list[i].kill_label.configure(text=self.mbot_list[i].get_kills_per_hour())
        self.update_hp_mp_job = self.after(2000, self.update_hp_mp)

    def select_all(self):
        self.bot_name_list.select_set(0, tk.END)

    def deselect_all(self):
        self.bot_name_list.select_clear(0, tk.END)

    def start_client_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.set_delay(False)),
                (100, lambda mbot: mbot.save_settings()),
                (100, lambda mbot: mbot.save_settings()),
                (10000, lambda mbot: mbot.start_client())
            ]
        )

    def kill_client_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        confirm = messagebox.askyesno("Confirmation", "Kill all client selected mBots ?")
        if not confirm:
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.kill_client())
            ]
        )

    def kill_mbot_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        confirm = messagebox.askyesno("Confirmation", "Close all selected mBots?")
        if not confirm:
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        if self.update_hp_mp_job:
            self.after_cancel(self.update_hp_mp_job)
            self.update_hp_mp_job = None

        reset_time = len(mbot_list) * 100 + 500
        self.after(reset_time, lambda: self.refresh_list())
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.kill_mbot())
            ]
        )

    def log_off_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        confirm = messagebox.askyesno("Confirmation", "Log off all selected mBots?")
        if not confirm:
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.set_delay()),
                (100, lambda mbot: mbot.save_settings()),
                (100, lambda mbot: mbot.save_settings()),
                (100, lambda mbot: mbot.log_off())
            ]
        )

    def show_hide_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.show_hide_mbot())
            ]
        )

    def show_hide_client_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.show_hide_client())
            ]
        )

    def reset_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.reset_mbot())
            ]
        )

    def get_current_position_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.get_current_position())
            ]
        )

    def start_training_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.set_delay(False)),
                (100, lambda mbot: mbot.save_settings()),
                (100, lambda mbot: mbot.save_settings()),
                (100, lambda mbot: mbot.start_training())
            ]
        )

    def stop_training_selected(self):
        selected = self.bot_name_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose at least one mBot")
            return

        mbot_list = [self.mbot_list[i] for i in selected]
        self.process_mbots(
            mbot_list,
            actions=[
                (0, lambda mbot: mbot.stop_training())
            ]
        )

class Account(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.account_list = None
        self.pending_start_mbot = None
        self.username_entry = None
        self.password_entry = None
        self.delay_entry = None
        self.character_entry = None
        self.username_var = None
        self.password_var = None
        self.delay_var = None
        self.character_var = None
        self.mbot_file_var = tk.StringVar()
        self.sro_folder_var = tk.StringVar()

        self.create_widgets()
        self.update_listbox()

    def process_mbots(self, mbot_list, actions):
        def process_next(index=0):
            if index >= len(mbot_list):
                return

            mbot = mbot_list[index]
            current_time = 0
            for delay_ms, method in actions:
                current_time += delay_ms
                self.after(current_time, lambda m=mbot, meth=method: meth(m))

            total_delay = sum(delay for delay, _ in actions)
            self.after(total_delay + 100, lambda: process_next(index + 1))

        process_next(0)

    def create_widgets(self):
        left_frame_width = WINDOW_WIDTH * 2 / 12
        left_frame_height = WINDOW_HEIGHT
        left_frame_x = TOP_LEFT_X
        left_frame_y = TOP_LEFT_Y
        self.create_left_frame(x=left_frame_x, y=left_frame_y, width=left_frame_width, height=left_frame_height)

        right_frame_width = WINDOW_WIDTH - left_frame_width
        right_frame_height = WINDOW_HEIGHT
        right_frame_x = WINDOW_WIDTH - right_frame_width
        right_frame_y = TOP_LEFT_Y
        self.create_right_frame(x=right_frame_x, y=right_frame_y, width=right_frame_width, height=right_frame_height)

    def create_left_frame(self, x, y, width, height):
        frame = ttk.Frame(self)
        frame.place(x=x, y=y, width=width, height=height)

        self.create_account_list_frame(frame)
        self.create_button_list_frame(frame)

    def create_account_list_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)
        ttk.Label(frame, text="Account").pack(anchor='w')

        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.account_list = tk.Listbox(frame, selectmode=tk.EXTENDED, height=15, yscrollcommand=scrollbar.set)
        self.account_list.pack(side=tk.LEFT, fill='both', expand=True)
        self.account_list.bind("<<ListboxSelect>>", self.show_account)

        scrollbar.config(command=self.account_list.yview)

    def create_button_list_frame(self, parent_frame):
        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill='x', pady=2)

        ttk.Button(button_frame, text="Delete account", command=self.delete_account).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Select all mBots", command=self.select_all).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Deselect all mBots", command=self.deselect_all).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Start mbBot & login SRO client", command=self.start_mbot_and_login_sro_selected).pack(fill='x', pady=1)
        ttk.Button(button_frame, text="Disable always-on-top SRO client", command=self.disable_always_on_top).pack(fill='x', pady=1)

    def select_all(self):
        self.account_list.select_set(0, tk.END)

    def deselect_all(self):
        self.account_list.select_clear(0, tk.END)

    def hide_mbot_and_sro_client(self, index):
        index_real = self.pending_start_mbot[index]
        character = accounts[index_real]["character"]
        mbot_list = findwindows.find_elements(class_name="#32770", title=f"[{character}] mBot v1.12b (vSRO 110)")
        if not mbot_list:
            return

        mbot_list_process = []
        mbot_list_process.append(MBot(mbot_list[0]))
        self.process_mbots(
            mbot_list_process,
            actions=[
                (0, lambda mbot: mbot.set_delay()),
                (100, lambda mbot: mbot.save_settings()),
                (100, lambda mbot: mbot.save_settings()),
                (100, lambda mbot: mbot.start_training()),
                (100, lambda mbot: mbot.show_hide_client()),
                (100, lambda mbot: mbot.show_hide_mbot())
            ]
        )
        self.after(2000, lambda: self.start_mbot_client(index + 1))

    def start_training_sro(self, index):
        index_real = self.pending_start_mbot[index]
        character = accounts[index_real]["character"]
        windows = findwindows.find_elements(class_name="CLIENT", title=character)
        if not windows:
            return self.after(5000, lambda: self.start_training_sro(index))

        handle = windows[0].handle
        win32gui.SetWindowPos(
            handle,
            win32con.HWND_NOTOPMOST,
            0, 0, 0, 0,
            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
        )

        mbot_list = findwindows.find_elements(class_name="#32770", title=f"[{character}] mBot v1.12b (vSRO 110)")
        if not mbot_list:
            return
        mbot = MBot(mbot_list[0])
        mbot.start_training()

        self.after(5000, lambda: self.hide_mbot_and_sro_client(index))

    def login_sro(self, index):
        def step_find_window():
            windows = findwindows.find_elements(class_name="CLIENT", title="SRO_Client")
            if not windows:
                return self.after(2000, lambda: self.login_sro(index))

            handle = windows[0].handle
            ctypes.windll.user32.ShowWindow(handle, 5)

            win32gui.SetWindowPos(
                handle,
                win32con.HWND_TOPMOST,
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
            )

            rect = win32gui.GetWindowRect(handle)
            left, top, right, bottom = rect

            width = right - left
            height = bottom - top
            center_x = left + width//2
            center_y = top + height//2

            self.after(200, lambda: step_click_center(center_x, center_y))

        def step_click_center(x, y):
            auto.Click(x, y)
            self.after(200, lambda: step_click_server(x, y))

        def step_click_server(x, y):
            auto.Click(x, y - 125)
            self.after(200, lambda: step_click_choose_server(x, y))

        def step_click_choose_server(x, y):
            auto.Click(x - 50, y + 200)
            self.after(200, lambda: step_enter_credentials())

        def step_enter_credentials():
            index_real = self.pending_start_mbot[index]
            username = accounts[index_real]["username"]
            password = accounts[index_real]["password"]
            password_decode = base64.b64decode(password).decode("utf-8")
            auto.SendKeys('{Tab}', interval=0.08)
            auto.SendKeys(username, interval=0.08)
            auto.SendKeys('{Tab}', interval=0.08)
            auto.SendKeys(password_decode, interval=0.08)
            auto.SendKeys('{Enter}', interval=0.08)
            self.after(15000, lambda: self.start_training_sro(index))

        step_find_window()

    def start_client_sro(self, index):
        mbot_list = findwindows.find_elements(class_name="#32770", title="mBot v1.12b (vSRO 110)")
        if not mbot_list:
            return self.after(2000, lambda: self.start_client_sro(index))
        mbot = MBot(mbot_list[0])
        mbot.start_client()
        self.after(10000, lambda: self.login_sro(index))

    def confirm_if_need(self, index):
        confirm = click_confirmation(title="Electus.*", text="OK", is_re=True)
        if confirm:
            self.after(1000, lambda: self.confirm_if_need(index))
        else:
            self.after(2000, lambda: self.start_client_sro(index))

    def start_mbot_client(self, index):
        if index >= len(self.pending_start_mbot):
            return
        
        index_real = self.pending_start_mbot[index]
        character = accounts[index_real]["character"]
        mbot_list = findwindows.find_elements(class_name="#32770", title=f"[{character}] mBot v1.12b (vSRO 110)", visible_only=False)
        if mbot_list:
            return self.after(1000, lambda: self.start_mbot_client(index + 1))
        
        mbot_list = findwindows.find_elements(class_name="#32770", title=f"[{character} - DC] mBot v1.12b (vSRO 110)", visible_only=False)
        if mbot_list:
            return self.after(1000, lambda: self.start_mbot_client(index + 1))

        mbot_path = accounts[index_real]["mbot_file_path"]
        folder_path = os.path.normpath(os.path.dirname(mbot_path))
        subprocess.Popen(mbot_path,cwd=folder_path)

        self.after(15000, lambda: self.confirm_if_need(index))

    def start_mbot_and_login_sro_selected(self):
        selected = self.account_list.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please choose one mBot")
            return

        self.pending_start_mbot = selected
        self.start_mbot_client(0)

    def disable_always_on_top(self):
        windows = findwindows.find_elements(class_name="CLIENT", visible_only=False)
        for window in windows:
            hwnd = window.handle
            win32gui.SetWindowPos(
                    hwnd,
                    win32con.HWND_NOTOPMOST,
                    0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
                )

    def create_right_frame(self, x, y, width, height):
        frame = ttk.Frame(self)
        frame.place(x=x, y=y, width=width, height=height)

        ttk.Label(frame, text="Sign up").pack(anchor='w')
        self.create_sign_up_frame(frame)
        ttk.Label(frame, text="Detail").pack(anchor='w')
        self.create_detail_frame(frame)

    def create_sign_up_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)

        tk.Label(frame, text="Username:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.username_entry = tk.Entry(frame)
        self.username_entry.grid(row=0, column=1, padx=5, pady=2, sticky="w")

        tk.Label(frame, text="Password:").grid(row=0, column=2, padx=5, pady=2, sticky="w")
        self.password_entry = tk.Entry(frame, show="*")
        self.password_entry.grid(row=0, column=3, padx=5, pady=2, sticky="w")

        tk.Label(frame, text="Delay:").grid(row=0, column=4, padx=5, pady=2, sticky="w")
        self.delay_entry = tk.Entry(frame)
        self.delay_entry.grid(row=0, column=5, padx=5, pady=2, sticky="w")

        tk.Label(frame, text="Character:").grid(row=0, column=6, padx=5, pady=2, sticky="w")
        self.character_entry = tk.Entry(frame)
        self.character_entry.grid(row=0, column=7, padx=5, pady=2, sticky="w")

        tk.Button(frame, text="Add account", command=self.add_account).grid(row=0, column=8, padx=5, pady=2, sticky="w")

    def create_detail_frame(self, parent_frame):
        frame = ttk.Frame(parent_frame)
        frame.pack(fill='x', pady=2)

        tk.Label(frame, text="Username:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.username_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.username_var, state='readonly').grid(row=0, column=1, padx=5, pady=2, sticky="w")

        tk.Label(frame, text="Password:").grid(row=0, column=2, padx=5, pady=2, sticky="w")
        self.password_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.password_var, show="*").grid(row=0, column=3, padx=5, pady=2, sticky="w")

        tk.Label(frame, text="Delay:").grid(row=0, column=4, padx=5, pady=2, sticky="w")
        self.delay_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.delay_var).grid(row=0, column=5, padx=5, pady=2, sticky="w")

        tk.Label(frame, text="Character:").grid(row=0, column=6, padx=5, pady=2, sticky="w")
        self.character_var = tk.StringVar()
        tk.Entry(frame, textvariable=self.character_var).grid(row=0, column=7, padx=5, pady=2, sticky="w")

        tk.Button(frame, text="Update account", command=self.update_account).grid(row=0, column=8, padx=5, pady=2, sticky="w")

        mbot_file_button = tk.Button(frame, text="Select file start mBot", command=self.select_mbot_file)
        mbot_file_button.grid(row=1, column=0, padx=5, pady=5)
        tk.Entry(frame, textvariable=self.mbot_file_var, state='readonly').grid(row=1, column=1, padx=5, pady=2, sticky="we", columnspan=8)

        sro_folder_button = tk.Button(frame, text="   Select folder SRO   ", command=self.select_sro_folder)
        sro_folder_button.grid(row=2, column=0, padx=5, pady=5)
        tk.Entry(frame, textvariable=self.sro_folder_var, state='readonly').grid(row=2, column=1, padx=5, pady=2, sticky="we", columnspan=8)

    def select_mbot_file(self):
        if not self.username_var.get():
            messagebox.showwarning("Warning", "Please choose one account")
            return

        mbot_path = filedialog.askopenfilename(
            title="Select mBot file",
            filetypes=[("Applications", "*.exe"), ("All files", "*.*")]
        )
        if mbot_path:
            mbot_path = mbot_path.replace("/", "\\")
            self.mbot_file_var.set(mbot_path)
            folder_path = os.path.normpath(os.path.dirname(mbot_path))
            for f in os.listdir(folder_path):
                if f.strip().lower() == "config.ini":
                    config_path = os.path.join(folder_path, f)
                    break

            if not config_path:
                return

            with open(config_path, "r", encoding="utf-16-le", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if line.lower().startswith("srodir="):
                        srodir_value = line.split("=", 1)[1]
                        self.sro_folder_var.set(srodir_value)
                        break
                    else:
                        continue

    def select_sro_folder(self):
        if not self.username_var.get():
            messagebox.showwarning("Warning", "Please choose one account")
            return

        shell = win32com.client.Dispatch("Shell.Application")
        folder = shell.BrowseForFolder(0, "Select Folder", 0, 0)
        if folder:
            self.sro_folder_var.set(folder.Self.Path)

    def update_listbox(self):
        self.account_list.delete(0, tk.END)
        for account in accounts:
            self.account_list.insert(tk.END, account["username"])

    def show_account(self, event):
        selection = self.account_list.curselection()
        if selection:
            index = selection[0]
            account = accounts[index]
            self.username_var.set(account["username"])
            self.password_var.set(base64.b64decode(account["password"]).decode('utf-8'))
            self.delay_var.set(str(account.get("delay_time", 0)))
            self.character_var.set(account["character"])
            self.mbot_file_var.set(account["mbot_file_path"])
            self.sro_folder_var.set(account["sro_folder_path"])

    def add_account(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        delay_time = self.delay_entry.get().strip()
        character = self.character_entry.get().strip()

        if not username or not password:
            messagebox.showwarning("Input Error", "Username and password required!")
            return

        if not character:
            messagebox.showwarning("Input Error", "Character required!")
            return

        try:
            delay_time_int = int(delay_time) if delay_time else 0
        except ValueError:
            messagebox.showwarning("Input Error", "Delay time must be an integer!")
            return

        for account in accounts:
            if account["username"] == username:
                messagebox.showerror("Error", "Username already exists!")
                return

        encoded_password = base64.b64encode(password.encode('utf-8')).decode('utf-8')
        accounts.append({"username": username, "password": encoded_password, "delay_time": delay_time_int, "character": character,"mbot_file_path": "", "sro_folder_path": ""})
        save_accounts()
        self.update_listbox()
        self.username_entry.delete(0, tk.END)
        self.password_entry.delete(0, tk.END)
        self.delay_entry.delete(0, tk.END)
        self.character_entry.delete(0, tk.END)
        messagebox.showinfo("Success", "Account added!")

    def update_folder_sro(self):
        sro_folder_path = self.sro_folder_var.get()
        mbot_path = self.mbot_file_var.get()

        folder_path = os.path.normpath(os.path.dirname(mbot_path))
        config_path = None

        for f in os.listdir(folder_path):
            if f.strip().lower() == "config.ini":
                config_path = os.path.join(folder_path, f)
                break

        if not config_path:
            return

        with open(config_path, "r", encoding="utf-16-le", errors="ignore") as f:
            lines = f.readlines()

        for i, line in enumerate(lines):
            if line.strip().lower().startswith("srodir="):
                lines[i] = f"srodir={sro_folder_path}\n"
                break

        with open(config_path, "w", encoding="utf-16-le") as f:
            f.writelines(lines)

    def update_account(self):
        selection = self.account_list.curselection()
        if not selection:
            messagebox.showwarning("Select account", "Please select an account to update!")
            return
        index = selection[0]

        new_password = self.password_var.get()
        new_delay = self.delay_var.get()
        new_character = self.character_var.get()
        new_mbot_file_path = self.mbot_file_var.get()
        new_sro_folder_path = self.sro_folder_var.get()

        if not new_mbot_file_path or not new_sro_folder_path:
            messagebox.showwarning("Warning", "Please select file start mBot and select folder SRO")
            return
        self.update_folder_sro()

        try:
            new_delay_int = int(new_delay) if new_delay else 0
        except ValueError:
            messagebox.showwarning("Input Error", "Delay time must be an integer!")
            return

        encoded_password = base64.b64encode(new_password.encode('utf-8')).decode('utf-8')
        accounts[index]["password"] = encoded_password
        accounts[index]["delay_time"] = new_delay_int
        accounts[index]["character"] = new_character
        accounts[index]["mbot_file_path"] = new_mbot_file_path
        accounts[index]["sro_folder_path"] = new_sro_folder_path

        save_accounts()
        self.update_listbox()
        messagebox.showinfo("Updated", "Account updated!")

    def delete_account(self):
        selection = self.account_list.curselection()
        if not selection:
            messagebox.showwarning("Select account", "Please select an account to remove!")
            return
        index = selection[0]
        confirm = messagebox.askyesno("Confirmation", f"Delete '{accounts[index]['username']}'?")
        if confirm:
            accounts.pop(index)
            save_accounts()
            self.update_listbox()
            self.username_var.set("")
            self.password_var.set("")
            self.delay_var.set("")
            self.character_var.set("")
            self.mbot_file_var.set("")
            self.sro_folder_var.set("")

class MBotManager(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("mBot Controller")
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.resizable(False, False)

        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill='both')

        self.account_page = ttk.Frame(notebook)
        notebook.add(self.account_page, text="Account")

        self.monitor_page =  ttk.Frame(notebook)
        notebook.add(self.monitor_page, text="Monitor")

    def run(self):
        account = Account(self.account_page)
        account.pack(expand=True, fill='both')
        monitor = Monitor(self.monitor_page)
        monitor.pack(expand=True, fill='both')

if __name__ == "__main__":
    accounts = load_accounts()
    app = MBotManager()
    app.run()
    app.mainloop()
