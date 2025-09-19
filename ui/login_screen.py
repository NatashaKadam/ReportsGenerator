import hashlib
import time
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLineEdit, QPushButton, QLabel, QFrame, QApplication
from PyQt6.QtCore import QTimer, Qt
from ui.widgets.dialogs import show_message_box
from core.constants import ADMIN_USER, ADMIN_PASS_HASH, SESSION_TIMEOUT

class LoginScreen(QWidget):
    def __init__(self, stack):
        super().__init__()
        self.stack = stack
        self.login_attempts = 0
        self.max_attempts = 3
        self.lockout_time = 0
        self.setup_ui()

    def setup_ui(self):
        card = QFrame(objectName="LoginCard")
        layout = QVBoxLayout(card)
        layout.setSpacing(20)

        header = QLabel(self.tr("Login to Reports Generator"), objectName="Header")
        header.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.username = QLineEdit(placeholderText=self.tr("Username"))
        self.password = QLineEdit(placeholderText=self.tr("Password"), echoMode=QLineEdit.EchoMode.Password)
        self.login_btn = QPushButton(self.tr("Login"), clicked=self.try_login)

        self.username.returnPressed.connect(self.password.setFocus)
        self.password.returnPressed.connect(self.try_login)

        layout.addWidget(header)
        layout.addWidget(self.username)
        layout.addWidget(self.password)
        layout.addWidget(self.login_btn)

        outer_layout = QVBoxLayout(self)
        outer_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        outer_layout.addWidget(card)
        self.update_styles(True)
        
    def retranslate(self):
        self.findChild(QLabel, "Header").setText(self.tr("Login to Reports Generator"))
        self.username.setPlaceholderText(self.tr("Username"))
        self.password.setPlaceholderText(self.tr("Password"))
        self.login_btn.setText(self.tr("Login"))
        
    def tr(self, text):
        return QApplication.instance().tr(text)

    def try_login(self):
        current_time = time.time()
        if self.lockout_time > current_time:
            remaining = int(self.lockout_time - current_time)
            show_message_box(self.tr("Account Locked"), self.tr(f"Too many failed attempts. Try again in {remaining} seconds."))
            return

        username = self.username.text()
        password_hash = hashlib.sha256(self.password.text().encode()).hexdigest()

        if username == ADMIN_USER and password_hash == ADMIN_PASS_HASH:
            self.login_attempts = 0
            self.stack.setCurrentIndex(1)
            QTimer.singleShot(SESSION_TIMEOUT, self.session_timeout)
        else:
            self.login_attempts += 1
            remaining_attempts = self.max_attempts - self.login_attempts
            if remaining_attempts > 0:
                show_message_box(self.tr("Login Failed"), self.tr(f"Invalid credentials. {remaining_attempts} attempts remaining."))
            else:
                self.lockout_time = time.time() + 300
                show_message_box(self.tr("Account Locked"), self.tr("Too many failed attempts. Try again in 5 minutes."))

    def session_timeout(self):
        if self.stack.currentIndex() == 1:
            self.stack.setCurrentIndex(0)
            show_message_box(self.tr("Session Expired"), self.tr("Your session has expired. Please login again."))
            self.username.clear()
            self.password.clear()

    def update_styles(self, dark_mode):
        if dark_mode:
            self.setStyleSheet("""
                QWidget { background-color: #353535; }
                QFrame#LoginCard { background-color: #252525; border: 1px solid #444; border-radius: 8px; padding: 32px; max-width: 400px; }
                QLabel#Header { font-size: 24px; font-weight: 600; color: #FFFFFF; }
                QLineEdit { padding: 12px; border: 1px solid #444; border-radius: 4px; font-size: 14px; background-color: #353535; color: #FFFFFF; }
                QLineEdit:focus { border-color: #3F51B5; background-color: #444; }
                QPushButton { padding: 12px; background-color: #3F51B5; color: white; border: none; border-radius: 4px; font-size: 16px; font-weight: bold; }
                QPushButton:hover {{ background-color: #303F9F; }}
            """)
        else:
            self.setStyleSheet("""
                QWidget { background-color: #F5F5F5; }
                QFrame#LoginCard { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; padding: 32px; max-width: 400px; }
                QLabel#Header { font-size: 24px; font-weight: 600; color: #212121; }
                QLineEdit { padding: 12px; border: 1px solid #E0E0E0; border-radius: 4px; font-size: 14px; background-color: #FAFAFA; color: #212121; }
                QLineEdit:focus { border-color: #3F51B5; background-color: #FFFFFF; }
                QPushButton { padding: 12px; background-color: #3F51B5; color: white; border: none; border-radius: 4px; font-size: 16px; font-weight: bold; }
                QPushButton:hover {{ background-color: #303F9F; }}
            """)