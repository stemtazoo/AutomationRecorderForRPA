import tkinter as tk
from tkinter import ttk
import webbrowser
import logging
import pyautogui


class KeyTab:
    """Tab for recording keyboard operations."""

    def __init__(self, app):
        """キーボード操作タブのウィジェットを初期化します。"""

        self.app = app
        self.frame = ttk.Frame(app.notebook)
        app.notebook.add(self.frame, text='キー操作')

        self.text_widget_key = tk.Text(self.frame, wrap=tk.WORD, font=("Arial", 14), height=2)
        self.text_widget_key.pack(pady=20)
        self.text_widget_key.insert(tk.END, "Enter the key or text to generate code")
        self.text_widget_key.config(state=tk.DISABLED)

        self.operation_var_key = tk.StringVar(self.frame)
        self.operation_var_key.set("操作方法を選択してください")
        self.operations_menu_key = tk.OptionMenu(
            self.frame,
            self.operation_var_key,
            "Press Key",
            "Write Text",
            "Hotkey",
        )
        self.operations_menu_key.pack(pady=10)

        self.operation_label_key1 = tk.Label(
            self.frame,
            text="文字列を入力してください。(Hotkeyの場合、キーの次に追加されます。)",
            font=("Arial", 12),
        )
        self.operation_label_key1.pack(pady=5)

        self.key_entry = tk.Entry(self.frame, font=("Arial", 14))
        self.key_entry.pack(pady=5)

        self.operation_label_key2 = tk.Label(
            self.frame,
            text="キーを選択してください。(Press KeyとHotkeyの場合、有効になります。)",
            font=("Arial", 12),
        )
        self.operation_label_key2.pack(pady=5)

        self.special_keys_frame = tk.Frame(self.frame)
        self.special_keys_frame.pack(pady=5)

        self.special_keys = [
            "ctrl",
            "alt",
            "shift",
            "tab",
            "up",
            "down",
            "left",
            "right",
            "f1",
            "space",
            "enter",
        ]
        self.special_key_vars = {}
        for key in self.special_keys:
            var = tk.BooleanVar()
            check = tk.Checkbutton(self.special_keys_frame, text=key.capitalize(), variable=var)
            check.pack(side=tk.LEFT, padx=5)
            self.special_key_vars[key] = var

        self.special_keys_link_label = tk.Label(
            self.frame,
            text="リストにないキーについては、次のリンクを参照してください:\nhttps://pyautogui.readthedocs.io/en/latest/keyboard.html#keyboard-keys",
            font=("Arial", 10),
            fg="blue",
            cursor="hand2",
        )
        self.special_keys_link_label.pack(pady=5)
        self.special_keys_link_label.bind(
            "<Button-1>",
            lambda e: self.open_url(
                "https://pyautogui.readthedocs.io/en/latest/keyboard.html#keyboard-keys"
            ),
        )

        self.execute_button_key = tk.Button(self.frame, text="ボタンを押すとコードが生成されます", command=self.generate_key_code)
        self.execute_button_key.pack(pady=10)

        self.operation_label_key3 = tk.Label(
            self.frame,
            text="実行ボタンを押すと、クリップボードにコピーされます。",
            font=("Arial", 12),
        )
        self.operation_label_key3.pack(pady=5)

    def generate_key_code(self):
        """入力内容からキーボード操作用コードを作成します。"""
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
            self.app.root.clipboard_clear()
            self.app.root.clipboard_append(code)
        except Exception:
            logging.error("An error occurred while generating the key code", exc_info=True)

    def open_url(self, url):
        """指定されたURLを既定のブラウザで開きます。"""
        try:
            webbrowser.open_new(url)
        except Exception:
            logging.error("An error occurred while opening the URL", exc_info=True)
