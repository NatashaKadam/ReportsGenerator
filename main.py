import sys
import os
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QLocale

from core.data_manager import db_setup
from core.utilities import setup_assets, CustomTranslator
from ui.main_window import BillApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Initialize translator based on system locale
    app.translator = CustomTranslator(app, QLocale.system().name().split('_')[0])
    app.installTranslator(app.translator)

    if not setup_assets():
        sys.exit(1)
        
    db_setup()

    win = BillApp()
    win.showMaximized()
    sys.exit(app.exec())