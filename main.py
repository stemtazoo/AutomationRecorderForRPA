"""Entry point for the Automation Recorder GUI application."""

import logging
from src.gui.automation_recorder import AutomationRecorderApp

# ログ設定
logging.basicConfig(filename='logs/app.log', level=logging.ERROR,
                    format='%(asctime)s %(levelname)s: %(message)s')

if __name__ == "__main__":
    try:
        app = AutomationRecorderApp()
        app.run()
    except Exception as e:
        logging.error("An error occurred", exc_info=True)
