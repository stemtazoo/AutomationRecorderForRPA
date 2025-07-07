import tkinter as tk
from tkinter import ttk
import logging
import pygetwindow as gw


class WindowTab:
    """Tab for listing all open windows."""

    def __init__(self, app):
        self.app = app
        self.frame = ttk.Frame(app.notebook)
        app.notebook.add(self.frame, text='ウィンドウ一覧')

        self.text_widget_window = tk.Text(self.frame, wrap=tk.WORD, font=("Arial", 14), height=14)
        self.text_widget_window.pack(pady=20)

        self.execute_button_window = tk.Button(self.frame, text="ウィンドウを取得", command=self.get_windows)
        self.execute_button_window.pack(pady=10)

    def get_windows(self):
        try:
            windows = gw.getAllTitles()
            self.text_widget_window.config(state=tk.NORMAL)
            self.text_widget_window.delete("1.0", tk.END)
            for window in windows:
                if window.strip():
                    self.text_widget_window.insert(tk.END, f"{window}\n")
            self.text_widget_window.config(state=tk.DISABLED)
        except Exception:
            logging.error("An error occurred while getting window titles", exc_info=True)
