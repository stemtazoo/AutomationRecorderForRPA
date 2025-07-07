import tkinter as tk
from tkinter import ttk
import logging
import pyautogui
import win32gui
from pynput import keyboard
from pywinauto.controls.hwndwrapper import HwndWrapper
from pywinauto import Desktop
from pywinauto.findwindows import ElementNotFoundError
from ..utils.inspector_utils import format_inspector_output, get_window_title_with_parent


class UIInspectorTab:
    """Tab for inspecting UI elements under the mouse cursor."""

    def __init__(self, app):
        self.app = app
        self.frame = ttk.Frame(app.notebook)
        app.notebook.add(self.frame, text='UI要素インスペクタ')

        label = tk.Label(self.frame, text="Ctrl+Shift+Xでマウス下のUI要素情報を取得します。", font=("Arial", 12))
        label.pack(pady=5)

        self.text_widget = tk.Text(self.frame, wrap=tk.WORD, font=("Arial", 12), height=15)
        self.text_widget.pack(padx=10, pady=10, fill="both", expand=True)
        self.text_widget.insert("end", "Ctrl+Shift+Xを押すと、ここにUI要素情報が表示されます。")
        self.text_widget.config(state="disabled")

        self.ctrl = self.shift = self.x = False
        self.start_hotkey_listener()

    def start_hotkey_listener(self):
        def on_press(key):
            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                self.ctrl = True
            elif key in [keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r]:
                self.shift = True
            elif hasattr(key, "char") and key.char and key.char.lower() == "x":
                self.x = True
            if self.ctrl and self.shift and self.x:
                self.inspect_element_under_cursor()

        def on_release(key):
            if key in [keyboard.Key.ctrl_l, keyboard.Key.ctrl_r]:
                self.ctrl = False
            elif key in [keyboard.Key.shift, keyboard.Key.shift_l, keyboard.Key.shift_r]:
                self.shift = False
            elif hasattr(key, "char") and key.char and key.char.lower() == "x":
                self.x = False

        listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        listener.daemon = True
        listener.start()

    def generate_code_example(self, elem):
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
        return "# 要素を特定する情報が不足しています"

    def get_element_under_mouse(self):
        try:
            x, y = pyautogui.position()
            elem = Desktop(backend="uia").from_point(x, y)
            return elem
        except ElementNotFoundError:
            return None

    def inspect_element_under_cursor(self):
        try:
            elem = self.get_element_under_mouse()
            if not elem:
                result = "要素が見つかりませんでした。"
            else:
                x, y = pyautogui.position()
                hwnd = win32gui.WindowFromPoint((x, y))
                window_title = get_window_title_with_parent(hwnd)
                dlg_code = f"""【dlg設定サンプル】
from pywinauto.application import Application
# backend は 'uia' または 'win32' から選べます
app = Application(backend=\"uia\").connect(title=\"{window_title}\")
dlg = app.window(title=\"{window_title}\")
# ↓このdlg変数を使って下のコード例をそのまま利用できます！
"""
                uia_info = {
                    "name": elem.window_text(),
                    "class_name": elem.element_info.class_name,
                    "control_type": elem.element_info.control_type,
                    "automation_id": elem.element_info.automation_id,
                    "rectangle": str(elem.rectangle()),
                    "code_example": self.generate_code_example(elem),
                }
                win32_wrap = HwndWrapper(hwnd)
                win32_info = {
                    "window_text": win32_wrap.window_text(),
                    "class_name": win32_wrap.friendly_class_name(),
                    "handle": win32_wrap.handle,
                    "rectangle": str(win32_wrap.rectangle()),
                    "code_example": f'dlg.child_window(title=\"{win32_wrap.window_text()}\", class_name=\"{win32_wrap.friendly_class_name()}\", handle={win32_wrap.handle}).click()',
                }
                result = f"{dlg_code}\n画面名: {window_title}\n\n{format_inspector_output(uia_info, win32_info)}"

            self.text_widget.config(state="normal")
            self.text_widget.delete("1.0", "end")
            self.text_widget.insert("end", result)
            self.text_widget.config(state="disabled")
        except Exception:
            logging.error("inspect_element_under_cursor error", exc_info=True)
            self.text_widget.config(state="normal")
            self.text_widget.delete("1.0", "end")
            self.text_widget.insert("end", "エラーが発生しました。")
            self.text_widget.config(state="disabled")
