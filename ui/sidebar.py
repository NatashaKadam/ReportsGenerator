import os
from PyQt6.QtWidgets import (
    QFrame, QHBoxLayout, QVBoxLayout, QListWidget, QLineEdit, QLabel,
    QToolButton, QSizePolicy, QApplication, QListWidgetItem, QMenu, QMessageBox
)
from PyQt6.QtCore import Qt, QPropertyAnimation, QEasingCurve, QSize, QDir, QPropertyAnimation
from PyQt6.QtGui import QIcon, QFont, QColor
from PyQt6.QtWidgets import QWidget
from PyQt6.QtWidgets import QApplication
from PyQt6.QtWidgets import (
    QFrame, QHBoxLayout, QVBoxLayout, QListWidget, QLineEdit, QLabel,
    QToolButton, QSizePolicy, QApplication, QListWidgetItem, QMenu, QMessageBox
)
import datetime

from core.constants import SCRIPT_DIR
from ui.widgets.dialogs import show_message_box
from core.data_manager import delete_session_from_db, load_sessions

class CollapsibleSidebar(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setFrameShadow(QFrame.Shadow.Raised)
        self.collapsed_width = 50
        self.expanded_width = 300
        self.setFixedWidth(self.expanded_width)
        self.main_layout = QHBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(5, 5, 5, 5)
        self.content_layout.setSpacing(5)
        self.search_bar = QLineEdit(placeholderText=self.tr("Search sessions..."))
        self.search_bar.setClearButtonEnabled(True)
        self.search_bar.setToolTip(self.tr("Search saved sessions"))
        self.search_bar.textChanged.connect(self.filter_sessions)
        self.history_label = QLabel(self.tr("Session History"))
        self.history_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.history_label.setObjectName("HistoryTitle")
        self.list_widget = QListWidget(self)
        self.list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.content_layout.addWidget(self.search_bar)
        self.content_layout.addWidget(self.history_label)
        self.content_layout.addWidget(self.list_widget, 1)
        self.toggle_button = QToolButton(self)
        self.toggle_button.setIcon(QIcon(os.path.join(SCRIPT_DIR, "assets", "left_arrow_icon.png")))
        self.toggle_button.setIconSize(QSize(24, 24))
        self.toggle_button.setToolTip(self.tr("Collapse/Expand Sidebar"))
        self.toggle_button.clicked.connect(self.toggle_state)
        self.main_layout.addWidget(self.content_widget, 1)
        self.main_layout.addWidget(self.toggle_button)
        self.animation = QPropertyAnimation(self, b"minimumWidth")
        self.animation.setDuration(500)
        self.animation.setEasingCurve(QEasingCurve.Type.InOutCubic)
        self.update_styles(True)
        
    def retranslate(self):
        self.search_bar.setPlaceholderText(self.tr("Search sessions..."))
        self.history_label.setText(self.tr("Session History"))
        self.toggle_button.setToolTip(self.tr("Collapse/Expand Sidebar"))
        parent_main_form = self.parent()
        if parent_main_form and hasattr(parent_main_form, 'refresh_sidebar'):
            parent_main_form.refresh_sidebar()

    def tr(self, text):
        return QApplication.instance().tr(text)

    def filter_sessions(self, text):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(Qt.ItemDataRole.UserRole) is not None:
                item.setHidden(text.lower() not in item.text().lower())

    def toggle_state(self):
        if self.width() > self.collapsed_width:
            self.animation.setStartValue(self.width())
            self.animation.setEndValue(self.collapsed_width)
            self.animation.start()
            self.toggle_button.setIcon(QIcon(os.path.join(SCRIPT_DIR, "assets", "right_arrow_icon.png")))
            for child in self.content_widget.findChildren(QWidget):
                child.hide()
            self.toggle_button.show()
        else:
            self.animation.setStartValue(self.width())
            self.animation.setEndValue(self.expanded_width)
            self.animation.start()
            self.toggle_button.setIcon(QIcon(os.path.join(SCRIPT_DIR, "assets", "left_arrow_icon.png")))
            for child in self.content_widget.findChildren(QWidget):
                child.show()

    def listWidget(self):
        return self.list_widget

    def searchBar(self):
        return self.search_bar

    def update_styles(self, dark_mode):
        if dark_mode:
            self.setStyleSheet("""
                QFrame { background-color: #353535; }
                QLineEdit { padding: 8px; border: 1px solid #444; border-radius: 4px; background-color: #252525; color: #FFFFFF; }
                QLabel#HistoryTitle { font-size: 14px; font-weight: bold; color: #FFFFFF; padding: 10px; background-color: #252525; border-bottom: 1px solid #444; border-top-left-radius: 8px; }
                QListWidget { border: none; background-color: #252525; padding-top: 5px; padding-bottom: 5px; }
                QListWidget::item { padding: 12px; border-bottom: 1px solid #444; color: #FFFFFF; }
                QListWidget::item:selected { background-color: #3F51B5; color: #FFFFFF; border-left: 3px solid #5C6BC0; }
                QListWidget::item:hover { background-color: #444; }
                QToolButton { background-color: #252525; border: none; }
            """)
        else:
            self.setStyleSheet("""
                QFrame { background-color: #FFFFFF; }
                QLineEdit { padding: 8px; border: 1px solid #CFD8DC; border-radius: 4px; background-color: #F8F8F8; color: #212121; }
                QLabel#HistoryTitle { font-size: 14px; font-weight: bold; color: #212121; padding: 10px; background-color: #F8F8F8; border-bottom: 1px solid #E0E0E0; border-top-left-radius: 8px; }
                QListWidget { border: none; background-color: #FFFFFF; padding-top: 5px; padding-bottom: 5px; }
                QListWidget::item { padding: 12px; border-bottom: 1px solid #F0F0F0; color: #212121; }
                QListWidget::item:selected { background-color: #E8EAF6; color: #303F9F; border-left: 3px solid #3F51B5; }
                QListWidget::item:hover { background-color: #F5F5F5; }
                QToolButton { background-color: #F8F8F8; border: none; }
            """)