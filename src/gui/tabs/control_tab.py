import tkinter as tk
from tkinter import ttk, filedialog
import logging
import io
from contextlib import redirect_stdout
import pygetwindow as gw
from pywinauto.application import Application


class ControlTab:
    """Tab for listing and saving window controls."""

    def __init__(self, app):
        self.app = app
        self.frame = ttk.Frame(app.notebook)
        app.notebook.add(self.frame, text='ウィンドウコントロール')

        self.update_windows_button = tk.Button(self.frame, text="ウィンドウリストを更新", command=self.update_window_list)
        self.update_windows_button.pack(pady=10)

        self.window_list_var = tk.StringVar(self.frame)
        self.window_list_menu = tk.OptionMenu(self.frame, self.window_list_var, '')
        self.window_list_menu.pack(pady=10)

        backend_frame = tk.LabelFrame(self.frame, text="バックエンドを選択", font=("Arial", 10))
        backend_frame.pack(pady=5)
        tk.Radiobutton(backend_frame, text="win32", variable=app.backend_var, value="win32").pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(backend_frame, text="uia", variable=app.backend_var, value="uia").pack(side=tk.LEFT, padx=10)

        self.get_control_button = tk.Button(self.frame, text="コントロールを取得", command=self.get_window_controls)
        self.get_control_button.pack(pady=10)

        self.text_widget_control = tk.Text(self.frame, wrap=tk.WORD, font=("Arial", 14), height=14)
        self.text_widget_control.pack(pady=20)

        self.save_button_control = tk.Button(self.frame, text="コントロールを保存", command=self.save_controls_to_file)
        self.save_button_control.pack(pady=10)

    def update_window_list(self):
        try:
            windows = [w for w in gw.getAllTitles() if w.strip()]
            menu = self.window_list_menu["menu"]
            menu.delete(0, "end")
            for window in windows:
                menu.add_command(label=window, command=lambda value=window: self.window_list_var.set(value))
            if windows:
                self.window_list_var.set(windows[0])
        except Exception:
            logging.error("An error occurred while updating the window list", exc_info=True)

    def get_window_controls(self):
        # Always clear the text widget when attempting to get controls
        self.text_widget_control.config(state=tk.NORMAL)
        self.text_widget_control.delete("1.0", tk.END)
        self.text_widget_control.config(state=tk.DISABLED)

        try:
            selected_window = self.window_list_var.get()
            if not selected_window:
                return
            backend = self.app.backend_var.get()
            app = Application(backend=backend).connect(title=selected_window)
            window = app.window(title=selected_window)
            f = io.StringIO()
            with redirect_stdout(f):
                window.print_control_identifiers()
            output = f.getvalue()
            self.text_widget_control.config(state=tk.NORMAL)
            self.text_widget_control.insert(tk.END, output)
            self.text_widget_control.config(state=tk.DISABLED)
        except Exception:
            logging.error("An error occurred while getting window controls", exc_info=True)
            self.text_widget_control.config(state=tk.NORMAL)
            self.text_widget_control.delete("1.0", tk.END)
            self.text_widget_control.insert(tk.END, "コントロールを取得できません")
            self.text_widget_control.config(state=tk.DISABLED)

    def save_controls_to_file(self):
        try:
            controls = self.text_widget_control.get("1.0", tk.END)
            if controls.strip():
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".txt",
                    filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                )
                if file_path:
                    with open(file_path, "w", encoding="utf-8") as file:
                        file.write(controls)
        except Exception:
            logging.error("An error occurred while saving controls to file", exc_info=True)
