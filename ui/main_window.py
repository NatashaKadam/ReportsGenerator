import os
from PyQt6.QtWidgets import QWidget, QStackedLayout
from PyQt6.QtGui import QIcon

from core.constants import SCRIPT_DIR
from .login_screen import LoginScreen
from .main_form import MainForm

class BillApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Reports Generator")
        self.setWindowIcon(QIcon(os.path.join(SCRIPT_DIR, "assets", "app_icon.png")))
        self.stack = QStackedLayout(self)

        self.login = LoginScreen(self.stack)
        self.form = MainForm()

        self.stack.addWidget(self.login)
        self.stack.addWidget(self.form)
        
    def changeEvent(self, event):
        if event.type() == event.type().LanguageChange:
            self.login.retranslate()
            self.form.retranslate()
        super().changeEvent(event)