"""Graphical user interface for recording automation actions on Windows."""

import tkinter as tk
from tkinter import ttk
from pynput import mouse

from .tabs.click_tab import ClickTab
from .tabs.key_tab import KeyTab
from .tabs.window_tab import WindowTab
from .tabs.control_tab import ControlTab
from .tabs.ui_inspector_tab import UIInspectorTab


class AutomationRecorderApp:
    """Main window for recording and generating automation scripts."""

    def __init__(self):
        """ウィンドウと各タブを作成し、マウス監視を開始します。"""

        self.root = tk.Tk()
        self.root.title("Automation Recorder")

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width = int(screen_width * 0.8)
        height = int(screen_height * 0.8)
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(600, 400)

        self.backend_var = tk.StringVar(value="win32")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")

        self.click_tab = ClickTab(self)
        self.key_tab = KeyTab(self)
        self.window_tab = WindowTab(self)
        self.control_tab = ControlTab(self)
        self.ui_inspector_tab = UIInspectorTab(self)

        self.listener = mouse.Listener(on_click=self.click_tab.on_click)
        self.listener.start()

    def run(self):
        """アプリケーションのメインループを開始します。"""

        self.root.mainloop()


if __name__ == "__main__":
    app = AutomationRecorderApp()
    app.run()
