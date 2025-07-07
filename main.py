"""Entry point for the Automation Recorder GUI application."""

import logging
import os
from src.gui.automation_recorder import AutomationRecorderApp

# Ensure logs directory and file exist
os.makedirs("logs", exist_ok=True)
log_path = os.path.join("logs", "app.log")
if not os.path.exists(log_path):
    open(log_path, "a").close()

# ログ設定
logging.basicConfig(filename=log_path, level=logging.ERROR,
                    format='%(asctime)s %(levelname)s: %(message)s')

if __name__ == "__main__":
    try:
        app = AutomationRecorderApp()
        app.run()
    except Exception as e:
        logging.error("An error occurred", exc_info=True)
