import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import webbrowser
import logging
import io
from contextlib import redirect_stdout

from pynput import mouse
import pyautogui
import pygetwindow as gw
from pywinauto.application import Application

class AutomationRecorderApp:
    def __init__(self):
        """Initializes the main application window and sets up the tabs."""
        self.root = tk.Tk()
        self.root.title("Automation Recorder")
        self.root.geometry("850x600")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both')

        self.setup_click_tab()
        self.setup_key_tab()
        self.setup_window_tab()
        self.setup_control_tab()

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

        self.window_list_var = tk.StringVar(self.control_tab)
        self.window_list_menu = tk.OptionMenu(self.control_tab, self.window_list_var, '')
        self.window_list_menu.pack(pady=10)

        self.update_windows_button = tk.Button(self.control_tab, text="ウィンドウリストを更新", command=self.update_window_list)
        self.update_windows_button.pack(pady=10)

        self.text_widget_control = tk.Text(self.control_tab, wrap=tk.WORD, font=("Arial", 14), height=14)
        self.text_widget_control.pack(pady=20)

        self.get_control_button = tk.Button(self.control_tab, text="コントロールを取得", command=self.get_window_controls)
        self.get_control_button.pack(pady=10)

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
            selected_window = self.window_list_var.get()
            if not selected_window:
                return

            app = Application(backend="uia").connect(title=selected_window)
            window = app.window(title=selected_window)
            f = io.StringIO()
            with redirect_stdout(f):
                window.print_control_identifiers()
            
            output = f.getvalue()

            self.text_widget_control.config(state=tk.NORMAL)
            self.text_widget_control.delete("1.0", tk.END)
            self.text_widget_control.insert(tk.END, output)
            self.text_widget_control.config(state=tk.DISABLED)
        except Exception as e:
            logging.error("An error occurred while getting window controls", exc_info=True)

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
        """Handles mouse click events to record click positions."""
        try:
            if pressed:
                self.root.update_idletasks()
                self.window_x = self.root.winfo_rootx()
                self.window_y = self.root.winfo_rooty()
                self.window_width = self.root.winfo_width()
                self.window_height = self.root.winfo_height()
                
                screen_x, screen_y = pyautogui.position()
                if not (self.window_x <= screen_x <= self.window_x + self.window_width and
                        self.window_y <= screen_y <= self.window_y + self.window_height):
                    self.screen_x, self.screen_y = screen_x, screen_y
                    self.text_widget_click.config(state=tk.NORMAL)
                    self.text_widget_click.delete("1.0", tk.END)
                    self.text_widget_click.insert(tk.END, f'Clicked at: ({screen_x}, {screen_y})')
                    self.text_widget_click.config(state=tk.DISABLED)
                    print(f'Clicked at: ({screen_x}, {screen_y})')
        except Exception as e:
            logging.error("An error occurred while processing the click event", exc_info=True)

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

if __name__ == "__main__":
    app = AutomationRecorderApp()
    app.run()
