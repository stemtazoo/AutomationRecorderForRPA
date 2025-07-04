"""Graphical user interface for recording automation actions on Windows."""

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import webbrowser
import logging
import io
from contextlib import redirect_stdout
import threading
import time
from ctypes import wintypes, windll, Structure, byref
import traceback

from pynput import mouse
from pynput import keyboard
import pygetwindow as gw
import pyautogui
import win32gui
from pywinauto.application import Application
from pywinauto.uia_defines import IUIA
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.uia_element_info import UIAElementInfo
from pywinauto import Desktop
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.controls.hwndwrapper import HwndWrapper
from ..utils.inspector_utils import (
    format_inspector_output,
    get_window_title_with_parent,
)

# POINT 構造体
class POINT(Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

def get_cursor_pos():
    """Return the current cursor position as an ``(x, y)`` tuple."""
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return pt.x, pt.y

def generate_code_example(elem):
    """Return example pywinauto code to click ``elem``."""
    props = []
    title = elem.window_text()
    ctrl_type = elem.element_info.control_type
    auto_id = elem.element_info.automation_id

    if title:
        props.append(f'title="{title}"')
    if ctrl_type:
        props.append(f'control_type="{ctrl_type}"')
    if auto_id:
        props.append(f'automation_id="{auto_id}"')

    if props:
        return f'dlg.child_window({", ".join(props)}).click_input()'
    else:
        return "# 要素を特定する情報が不足しています"

def get_element_under_mouse():
    """Return the UI element currently under the mouse cursor."""
    try:
        x, y = get_cursor_pos()
        elem = Desktop(backend="uia").from_point(x, y)
        return elem
    except ElementNotFoundError:
        return None

def get_element_info_text(elem):
    """Return formatted information text for ``elem`` or a not-found message."""
    if elem:
        title = elem.window_text()
        class_name = elem.element_info.class_name
        ctrl_type = elem.element_info.control_type
        auto_id = elem.element_info.automation_id
        rect = elem.rectangle()
        code = generate_code_example(elem)
        return (
            "=== 要素情報 ===\n"
            f"タイトル: {title}\n"
            f"クラス名: {class_name}\n"
            f"コントロールタイプ: {ctrl_type}\n"
            f"オートメーションID: {auto_id}\n"
            f"矩形: {rect}\n"
            f"コード例: {code}"
        )
    else:
        return "要素が見つかりませんでした。"


class AutomationRecorderApp:
    """Main window for recording and generating automation scripts."""
    def __init__(self):
        """Initializes the main application window and sets up the tabs."""
        self.root = tk.Tk()
        self.root.title("Automation Recorder")
        
        # ここにカスタムサイズ設定を追加！
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = int(screen_width * 0.8)
        height = int(screen_height * 0.8)
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(600, 400)
        
        self.backend_var = tk.StringVar(value="win32")  # デフォルトは win32

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both')

        self.setup_click_tab()
        self.setup_key_tab()
        self.setup_window_tab()
        self.setup_control_tab()
        self.setup_ui_inspector_tab()
        self.start_ui_inspector_hotkey_listener()

        self.listener = mouse.Listener(on_click=self.on_click)
        self.listener.start()

        self.screen_x = None
        self.screen_y = None

    def setup_click_tab(self):
        """Sets up the click tab for recording mouse click positions."""
        self.click_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.click_tab, text='クリック操作')

        self.text_widget_click = tk.Text(self.click_tab, wrap=tk.WORD, font=("Arial", 14), height=2)
        self.text_widget_click.pack(pady=20)
        self.text_widget_click.insert(tk.END, "Click outside this window to record position")
        self.text_widget_click.config(state=tk.DISABLED)

        self.operation_var_click = tk.StringVar(self.click_tab)
        self.operation_var_click.set("クリック方法を選択してください")
        self.operations_menu_click = tk.OptionMenu(
            self.click_tab, self.operation_var_click, "Left Click", "Right Click", "Double Click",
            "Move to", "Drag and Drop"
        )
        self.operations_menu_click.pack(pady=10)

        self.execute_button_click = tk.Button(self.click_tab, text="ボタンを押すとコードが生成されます", command=self.generate_click_code)
        self.execute_button_click.pack(pady=10)

        self.operation_label_click = tk.Label(self.click_tab, text="実行ボタンを押すと、クリップボードにコピーされます。", font=("Arial", 12))
        self.operation_label_click.pack(pady=5)

    def setup_key_tab(self):
        """Sets up the key tab for recording keyboard inputs."""
        self.key_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.key_tab, text='キー操作')

        self.text_widget_key = tk.Text(self.key_tab, wrap=tk.WORD, font=("Arial", 14), height=2)
        self.text_widget_key.pack(pady=20)
        self.text_widget_key.insert(tk.END, "Enter the key or text to generate code")
        self.text_widget_key.config(state=tk.DISABLED)

        self.operation_var_key = tk.StringVar(self.key_tab)
        self.operation_var_key.set("操作方法を選択してください")
        self.operations_menu_key = tk.OptionMenu(
            self.key_tab, self.operation_var_key, "Press Key", "Write Text", "Hotkey"
        )
        self.operations_menu_key.pack(pady=10)

        self.operation_label_key1 = tk.Label(self.key_tab, text="文字列を入力してください。(Hotkeyの場合、キーの次に追加されます。)", font=("Arial", 12))
        self.operation_label_key1.pack(pady=5)

        self.key_entry = tk.Entry(self.key_tab, font=("Arial", 14))
        self.key_entry.pack(pady=5)

        self.operation_label_key2 = tk.Label(self.key_tab, text="キーを選択してください。(Press KeyとHotkeyの場合、有効になります。)", font=("Arial", 12))
        self.operation_label_key2.pack(pady=5)

        self.special_keys_frame = tk.Frame(self.key_tab)
        self.special_keys_frame.pack(pady=5)

        self.special_keys = ["ctrl", "alt", "shift", "tab", "up", "down", "left", "right", "f1", "space", "enter"]
        self.special_key_vars = {}
        for key in self.special_keys:
            var = tk.BooleanVar()
            check = tk.Checkbutton(self.special_keys_frame, text=key.capitalize(), variable=var)
            check.pack(side=tk.LEFT, padx=5)
            self.special_key_vars[key] = var

        self.special_keys_link_label = tk.Label(self.key_tab, text="リストにないキーについては、次のリンクを参照してください:\nhttps://pyautogui.readthedocs.io/en/latest/keyboard.html#keyboard-keys", font=("Arial", 10), fg="blue", cursor="hand2")
        self.special_keys_link_label.pack(pady=5)
        self.special_keys_link_label.bind("<Button-1>", lambda e: self.open_url("https://pyautogui.readthedocs.io/en/latest/keyboard.html#keyboard-keys"))

        self.execute_button_key = tk.Button(self.key_tab, text="ボタンを押すとコードが生成されます", command=self.generate_key_code)
        self.execute_button_key.pack(pady=10)

        self.operation_label_key3 = tk.Label(self.key_tab, text="実行ボタンを押すと、クリップボードにコピーされます。", font=("Arial", 12))
        self.operation_label_key3.pack(pady=5)

    def setup_window_tab(self):
        """Sets up the window tab for listing all open windows."""
        self.window_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.window_tab, text='ウィンドウ一覧')

        self.text_widget_window = tk.Text(self.window_tab, wrap=tk.WORD, font=("Arial", 14), height=14)
        self.text_widget_window.pack(pady=20)

        self.execute_button_window = tk.Button(self.window_tab, text="ウィンドウを取得", command=self.get_windows)
        self.execute_button_window.pack(pady=10)

    def setup_control_tab(self):
        """Sets up the control tab for listing and saving window controls."""
        self.control_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.control_tab, text='ウィンドウコントロール')
        
        self.update_windows_button = tk.Button(self.control_tab, text="ウィンドウリストを更新", command=self.update_window_list)
        self.update_windows_button.pack(pady=10)

        self.window_list_var = tk.StringVar(self.control_tab)
        self.window_list_menu = tk.OptionMenu(self.control_tab, self.window_list_var, '')
        self.window_list_menu.pack(pady=10)

       # バックエンド選択ラジオボタン
        backend_frame = tk.LabelFrame(self.control_tab, text="バックエンドを選択", font=("Arial", 10))
        backend_frame.pack(pady=5)

        tk.Radiobutton(backend_frame, text="win32", variable=self.backend_var, value="win32").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(backend_frame, text="uia", variable=self.backend_var, value="uia").pack(side=tk.LEFT, padx=10)

        self.get_control_button = tk.Button(self.control_tab, text="コントロールを取得", command=self.get_window_controls)
        self.get_control_button.pack(pady=10)
        
        self.text_widget_control = tk.Text(self.control_tab, wrap=tk.WORD, font=("Arial", 14), height=14)
        self.text_widget_control.pack(pady=20)

        self.save_button_control = tk.Button(self.control_tab, text="コントロールを保存", command=self.save_controls_to_file)
        self.save_button_control.pack(pady=10)
        
    def get_windows(self):
        """Gets the titles of all open windows and displays them in the window tab."""
        try:
            windows = gw.getAllTitles()
            self.text_widget_window.config(state=tk.NORMAL)
            self.text_widget_window.delete("1.0", tk.END)
            for window in windows:
                if window.strip():  # 空のタイトルを無視
                    self.text_widget_window.insert(tk.END, f'{window}\n')
            self.text_widget_window.config(state=tk.DISABLED)
        except Exception as e:
            logging.error("An error occurred while getting window titles", exc_info=True)

    def update_window_list(self):
        """Updates the dropdown menu with the list of all open windows."""
        try:
            windows = [w for w in gw.getAllTitles() if w.strip()]
            menu = self.window_list_menu["menu"]
            menu.delete(0, "end")
            for window in windows:
                menu.add_command(label=window, command=lambda value=window: self.window_list_var.set(value))
            if windows:
                self.window_list_var.set(windows[0])
        except Exception as e:
            logging.error("An error occurred while updating the window list", exc_info=True)

    def get_window_controls(self):
        """Gets the control identifiers of the selected window and displays them in the control tab."""
        try:
            print(self.window_list_var.get())
            print('window control')
            selected_window = self.window_list_var.get()
            if not selected_window:
                return

            backend = self.backend_var.get()
            app = Application(backend=backend).connect(title=selected_window)
            window = app.window(title=selected_window)
            f = io.StringIO()
            with redirect_stdout(f):
                print(window.element_info.name)       # 実際に認識されたウィンドウタイトル
                print(window.element_info.class_name) # ウィンドウのクラス名
                print(window.children())
                window.print_control_identifiers()
            
            output = f.getvalue()

            self.text_widget_control.config(state=tk.NORMAL)
            self.text_widget_control.delete("1.0", tk.END)
            self.text_widget_control.insert(tk.END, output)
            self.text_widget_control.config(state=tk.DISABLED)
        except Exception as e:
            logging.error("An error occurred while getting window controls", exc_info=True)
            print(e)

    def save_controls_to_file(self):
        """Saves the displayed control identifiers to a text file."""
        try:
            controls = self.text_widget_control.get("1.0", tk.END)
            if controls.strip():
                file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
                if file_path:
                    with open(file_path, "w", encoding="utf-8") as file:
                        file.write(controls)
                    print(f'Controls saved to {file_path}')
        except Exception as e:
            logging.error("An error occurred while saving controls to file", exc_info=True)

    def on_click(self, x, y, button, pressed):
        try:
            if pressed:
                self.root.update_idletasks()
                screen_x, screen_y = pyautogui.position()

                if not (self.root.winfo_rootx() <= screen_x <= self.root.winfo_rootx() + self.root.winfo_width() and
                        self.root.winfo_rooty() <= screen_y <= self.root.winfo_rooty() + self.root.winfo_height()):
                    self.screen_x, self.screen_y = screen_x, screen_y
                    self.text_widget_click.config(state=tk.NORMAL)
                    self.text_widget_click.delete("1.0", tk.END)
                    self.text_widget_click.insert(tk.END, f'Clicked at: ({screen_x}, {screen_y})')
                    self.text_widget_click.config(state=tk.DISABLED)
                    print(f'Clicked at: ({screen_x}, {screen_y})')

                    # # 🛠 修正ここ！
                    # try:
                    #     element_info = UIAElementInfo.from_point(screen_x, screen_y)
                    #     ctrl = UIAWrapper(element_info)
                    #     info = ctrl.element_info
                    #     result = f"【クリック位置のコントロール】\n名前: {info.name}\n" \
                    #             f"クラス: {info.class_name}\nタイプ: {info.control_type}\n" \
                    #             f"AutomationId: {info.automation_id}"

                    #     self.text_widget_control.config(state=tk.NORMAL)
                    #     self.text_widget_control.delete("1.0", tk.END)
                    #     self.text_widget_control.insert(tk.END, result)
                    #     self.text_widget_control.config(state=tk.DISABLED)
                    #     print("クリック位置のコントロール取得完了")

                    # except Exception as ee:
                    #     print("コントロール取得時エラー:", ee)
        except Exception as e:
            print("コントロール取得時エラー:", e)

    def generate_click_code(self):
        """Generates and displays the code for the recorded click action."""
        try:
            operation = self.operation_var_click.get()
            if self.screen_x is None or self.screen_y is None:
                return

            code = ""
            if operation == "Left Click":
                code = f"pyautogui.click({self.screen_x}, {self.screen_y})"
            elif operation == "Right Click":
                code = f"pyautogui.rightClick({self.screen_x}, {self.screen_y})"
            elif operation == "Double Click":
                code = f"pyautogui.doubleClick({self.screen_x}, {self.screen_y})"
            elif operation == "Move to":
                code = f"pyautogui.moveTo({self.screen_x}, {self.screen_y})"
            elif operation == "Drag and Drop":
                code = f"pyautogui.dragTo({self.screen_x}, {self.screen_y}, duration=1)"

            self.text_widget_click.config(state=tk.NORMAL)
            self.text_widget_click.delete("1.0", tk.END)
            self.text_widget_click.insert(tk.END, code)
            self.text_widget_click.config(state=tk.DISABLED)
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
            print(f'Generated code: {code}')
        except Exception as e:
            logging.error("An error occurred while generating the click code", exc_info=True)

    def generate_key_code(self):
        """Generates and displays the code for the recorded key action."""
        try:
            operation = self.operation_var_key.get()
            key_value = self.key_entry.get()

            keys = []
            for key, var in self.special_key_vars.items():
                if var.get():
                    keys.append(key)
            if key_value:
                keys.append(key_value)

            code = ""
            if operation == "Press Key":
                if keys:
                    if len(keys) == 1:
                        code = f"pyautogui.press('{keys[0]}')"
                    else:
                        keys_str = "', '".join(keys)
                        code = f"pyautogui.hotkey('{keys_str}')"
            elif operation == "Write Text":
                if key_value:
                    code = f"pyautogui.write('{key_value}')"
            elif operation == "Hotkey":
                if keys:
                    keys_str = "', '".join(keys)
                    code = f"pyautogui.hotkey('{keys_str}')"

            self.text_widget_key.config(state=tk.NORMAL)
            self.text_widget_key.delete("1.0", tk.END)
            self.text_widget_key.insert(tk.END, code)
            self.text_widget_key.config(state=tk.DISABLED)
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
            print(f'Generated code: {code}')
        except Exception as e:
            logging.error("An error occurred while generating the key code", exc_info=True)

    def open_url(self, url):
        """Opens the given URL in the default web browser."""
        try:
            webbrowser.open_new(url)
        except Exception as e:
            logging.error("An error occurred while opening the URL", exc_info=True)

    def run(self):
        """Starts the main application loop."""
        try:
            self.root.mainloop()
        except Exception as e:
            logging.error("An error occurred in the main loop", exc_info=True)

    def start_ui_inspector_hotkey_listener(self):
        self.ctrl = False
        self.shift = False
        self.x = False

        def on_press(key):
            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                self.ctrl = True
            if key in [keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r]:
                self.shift = True
            # ここがポイント！
            if hasattr(key, 'char') and (key.char == 'x' or key.char == 'X' or key.char == '\x18'):
                self.x = True
            print(f"ctrl={self.ctrl} shift={self.shift} x={self.x}, key={key}")
            if self.ctrl and self.shift and self.x:
                print("Hotkey Detected! Ctrl+Shift+X")
                self.inspect_element_under_cursor()

        def on_release(key):
            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                self.ctrl = False
            if key in [keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r]:
                self.shift = False
            if hasattr(key, 'char') and (key.char == 'x' or key.char == 'X' or key.char == '\x18'):
                self.x = False

        listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        listener.daemon = True
        listener.start()

    def inspect_element_under_cursor(self):
        print("inspect_element_under_cursor called")
        try:
            x, y = pyautogui.position()
            hwnd = win32gui.WindowFromPoint((x, y))
            window_title = get_window_title_with_parent(hwnd)
            
            # dlgサンプル文字列
            dlg_code = f"""【dlg設定サンプル】
from pywinauto.application import Application
# backend は 'uia' または 'win32' から選べます
app = Application(backend="uia").connect(title="{window_title}")
dlg = app.window(title="{window_title}")
# ↓このdlg変数を使って下のコード例をそのまま利用できます！
"""
            
            # --- UIA側 ---
            try:
                from pywinauto.uia_element_info import UIAElementInfo
                from pywinauto.controls.uiawrapper import UIAWrapper
                uia_elem = UIAElementInfo.from_point(x, y)
                uia_wrap = UIAWrapper(uia_elem)
                uia_info = {
                    "name": uia_wrap.element_info.name,
                    "class_name": uia_wrap.element_info.class_name,
                    "control_type": uia_wrap.element_info.control_type,
                    "automation_id": uia_wrap.element_info.automation_id,
                    "rectangle": str(uia_wrap.element_info.rectangle),
                    "code_example": f'dlg.child_window(title=\"{uia_wrap.element_info.name}\", control_type=\"{uia_wrap.element_info.control_type}\", automation_id=\"{uia_wrap.element_info.automation_id}\").click_input()'
                }
            except Exception as e:
                print("UIA取得エラー", e)
                uia_info = {"name": "取得エラー", "class_name": str(e), "control_type": "", "automation_id": "", "rectangle": "", "code_example": ""}

            # --- Win32側 ---
            try:
                hwnd = win32gui.WindowFromPoint((x, y))
                win32_wrap = HwndWrapper(hwnd)
                win32_info = {
                    "window_text": win32_wrap.window_text(),
                    "class_name": win32_wrap.friendly_class_name(),
                    "handle": win32_wrap.handle,
                    "rectangle": str(win32_wrap.rectangle()),
                    "code_example": f'dlg.child_window(title=\"{win32_wrap.window_text()}\", class_name=\"{win32_wrap.friendly_class_name()}\", handle={win32_wrap.handle}).click()'
                }
            except Exception as e:
                print("Win32取得エラー", e)
                win32_info = {
                    "window_text": "取得エラー",
                    "class_name": str(e),
                    "handle": "",
                    "rectangle": "",
                    "code_example": ""
                }

            # --- 表示部分 ---
            result = (
                f"{dlg_code}\n"
                f"画面名: {window_title}\n\n"
                f"{format_inspector_output(uia_info, win32_info)}"
            )
            print("resultを表示します")
            self.text_widget_ui_inspector.config(state="normal")
            self.text_widget_ui_inspector.delete("1.0", "end")
            self.text_widget_ui_inspector.insert("end", result)
            self.text_widget_ui_inspector.config(state="disabled")
            print("テキストボックス更新完了")
        except Exception as e:
            print("inspect_element_under_cursor error:", e)
            self.text_widget_ui_inspector.config(state="normal")
            self.text_widget_ui_inspector.delete("1.0", "end")
            self.text_widget_ui_inspector.insert("end", f"エラー: {e}")
            self.text_widget_ui_inspector.config(state="disabled")

    def setup_ui_inspector_tab(self):
        """UI要素インスペクタタブを追加する"""
        self.ui_inspector_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.ui_inspector_tab, text='UI要素インスペクタ')
        
        label = tk.Label(self.ui_inspector_tab, text="Ctrl+Shift+Xでマウス下のUI要素情報を取得します。", font=("Arial", 12))
        label.pack(pady=5)
        
        self.text_widget_ui_inspector = tk.Text(self.ui_inspector_tab, wrap=tk.WORD, font=("Arial", 12), height=15)
        self.text_widget_ui_inspector.pack(padx=10, pady=10, fill="both", expand=True)
        self.text_widget_ui_inspector.insert("end", "Ctrl+Shift+Xを押すと、ここにUI要素情報が表示されます。")
        self.text_widget_ui_inspector.config(state="disabled")
        

if __name__ == "__main__":
    try:
        print("Automation Recorderを起動します...")
        app = AutomationRecorderApp()
        app.run()
    except Exception as e:
        print("起動時エラー:", e)
        traceback.print_exc()
        input("何かキーを押すと終了します")
