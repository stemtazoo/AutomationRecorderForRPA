import tkinter as tk
from tkinter import ttk
import pyautogui
import logging


class ClickTab:
    """Tab for recording mouse click positions."""

    def __init__(self, app):
        """クリック操作用のウィジェットを準備します。"""

        self.app = app
        self.frame = ttk.Frame(app.notebook)
        app.notebook.add(self.frame, text='クリック操作')

        self.text_widget_click = tk.Text(self.frame, wrap=tk.WORD, font=("Arial", 14), height=2)
        self.text_widget_click.pack(pady=20)
        self.text_widget_click.insert(tk.END, "Click outside this window to record position")
        self.text_widget_click.config(state=tk.DISABLED)

        self.operation_var_click = tk.StringVar(self.frame)
        self.operation_var_click.set("クリック方法を選択してください")
        self.operations_menu_click = tk.OptionMenu(
            self.frame,
            self.operation_var_click,
            "Left Click",
            "Right Click",
            "Double Click",
            "Move to",
            "Drag and Drop",
        )
        self.operations_menu_click.pack(pady=10)

        self.execute_button_click = tk.Button(self.frame, text="ボタンを押すとコードが生成されます", command=self.generate_click_code)
        self.execute_button_click.pack(pady=10)

        self.operation_label_click = tk.Label(
            self.frame,
            text="実行ボタンを押すと、クリップボードにコピーされます。",
            font=("Arial", 12),
        )
        self.operation_label_click.pack(pady=5)

        self.screen_x = None
        self.screen_y = None

    def on_click(self, x, y, button, pressed):
        """ウィンドウ外でのマウスクリック位置を取得して表示します。"""
        try:
            if pressed:
                self.app.root.update_idletasks()
                screen_x, screen_y = pyautogui.position()
                if not (
                    self.app.root.winfo_rootx()
                    <= screen_x
                    <= self.app.root.winfo_rootx() + self.app.root.winfo_width()
                    and self.app.root.winfo_rooty()
                    <= screen_y
                    <= self.app.root.winfo_rooty() + self.app.root.winfo_height()
                ):
                    self.screen_x, self.screen_y = screen_x, screen_y
                    self.text_widget_click.config(state=tk.NORMAL)
                    self.text_widget_click.delete("1.0", tk.END)
                    self.text_widget_click.insert(tk.END, f"Clicked at: ({screen_x}, {screen_y})")
                    self.text_widget_click.config(state=tk.DISABLED)
        except Exception:
            logging.error("Error detecting click position", exc_info=True)

    def generate_click_code(self):
        """選択した操作に対応するPyAutoGUIコードを生成します。"""
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
            self.app.root.clipboard_clear()
            self.app.root.clipboard_append(code)
        except Exception:
            logging.error("An error occurred while generating the click code", exc_info=True)
