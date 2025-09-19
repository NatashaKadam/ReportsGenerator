import os
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QDialogButtonBox,
    QTextEdit, QToolButton, QFileDialog, QLineEdit, QComboBox, QSpinBox, 
    QApplication, QAbstractButton, QSizePolicy, QFrame, QFormLayout
)
from PyQt6.QtCore import Qt, QSize, pyqtSignal, QAbstractAnimation, QVariantAnimation, QEasingCurve
from PyQt6.QtGui import QColor, QPalette, QPainter
from PyQt6.QtWidgets import QAbstractButton, QSizePolicy

class CustomMessageBox(QDialog):
    def __init__(self, title, text, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(450, 200)
        self.setStyleSheet("""
            QDialog { background-color: #353535; color: #FFFFFF; border: 1px solid #444; }
            QLabel { color: #FFFFFF; font-size: 14px; }
            QTextEdit { background-color: #252525; color: #FFFFFF; border: 1px solid #444; padding: 5px; }
            QPushButton { background-color: #3F51B5; color: white; border: none; padding: 8px; min-width: 80px; border-radius: 4px; }
            QPushButton:hover { background-color: #303F9F; }
        """)
        layout = QVBoxLayout(self)
        message_label = QLabel(text)
        message_label.setWordWrap(True)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(self.accept)
        layout.addWidget(message_label)
        layout.addWidget(button_box)

def show_message_box(title, message):
    dialog = CustomMessageBox(title, message)
    dialog.exec()

class MessageEditorDialog(QDialog):
    def __init__(self, current_text, parent=None):
        super().__init__(parent)
        self.setWindowTitle(self.tr("Edit Message"))
        self.setMinimumSize(700, 500)
        self.setModal(True)
        layout = QVBoxLayout(self)
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(current_text)
        font = self.text_edit.font()
        font.setPointSize(12)
        self.text_edit.setFont(font)
        layout.addWidget(self.text_edit)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def tr(self, text):
        return QApplication.instance().tr(text)

    def get_text(self):
        return self.text_edit.toPlainText()

class DetachedPreviewDialog(QDialog):
    closed = pyqtSignal()

    def __init__(self, preview_widget, parent=None):
        super().__init__(parent)
        self.setWindowTitle(self.tr("Detached Preview"))
        self.setMinimumSize(800, 600)
        self.setWindowFlags(Qt.WindowType.Window)
        self.preview_widget = preview_widget
        self.preview_widget.setStyleSheet("""
            QTextEdit {
                background-color: #FFFFFF;
                border: none;
                color: #212121;
                font-family: Arial, sans-serif;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        zoom_layout = QHBoxLayout()
        zoom_layout.addStretch()
        zoom_in_btn = QToolButton(objectName="zoomInButton")
        zoom_in_btn.setText("➕")
        zoom_in_btn.setToolTip(self.tr("Zoom In"))
        zoom_in_btn.clicked.connect(lambda: self.preview_widget.zoomIn(2))
        zoom_layout.addWidget(zoom_in_btn)
        zoom_out_btn = QToolButton(objectName="zoomOutButton")
        zoom_out_btn.setText("➖")
        zoom_out_btn.setToolTip(self.tr("Zoom Out"))
        zoom_out_btn.clicked.connect(lambda: self.preview_widget.zoomOut(2))
        zoom_layout.addWidget(zoom_out_btn)
        layout.addLayout(zoom_layout)
        layout.addWidget(self.preview_widget)

    def tr(self, text):
        return QApplication.instance().tr(text)

    def closeEvent(self, event):
        self.closed.emit()
        event.accept()
        
    def retranslate(self):
        self.setWindowTitle(self.tr("Detached Preview"))
        zoom_in_btn = self.findChild(QToolButton, "zoomInButton")
        if zoom_in_btn:
            zoom_in_btn.setToolTip(self.tr("Zoom In"))
        zoom_out_btn = self.findChild(QToolButton, "zoomOutButton")
        if zoom_out_btn:
            zoom_out_btn.setToolTip(self.tr("Zoom Out"))

class MaterialSwitch(QAbstractButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.setFixedSize(60, 30)
        self._thumb_color_on = QColor('#4CAF50')
        self._thumb_color_off = QColor('#BDBDBD')
        self._track_color_on = QColor('#A5D6A7')
        self._track_color_off = QColor('#E0E0E0')
        self._thumb_pos = 0.0
        self._animation = QVariantAnimation(self)
        self._animation.setStartValue(0.0)
        self._animation.setEndValue(1.0)
        self._animation.setDuration(250)
        self._animation.setEasingCurve(QEasingCurve.Type.InOutCubic)
        self._animation.valueChanged.connect(self._set_thumb_pos)
        self.toggled.connect(self.start_animation)
    
    def _set_thumb_pos(self, value):
        self._thumb_pos = value
        self.update()

    def start_animation(self, checked):
        if self._animation.state() == QAbstractAnimation.State.Running:
            self._animation.stop()
        if checked:
            self._animation.setDirection(QAbstractAnimation.Direction.Forward)
        else:
            self._animation.setDirection(QAbstractAnimation.Direction.Backward)
        self._animation.start()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        if self.isChecked():
            track_color = self._track_color_on
            thumb_color = self._thumb_color_on
        else:
            track_color = self._track_color_off
            thumb_color = self._thumb_color_off
        track_rect = self.rect().adjusted(2, 8, -2, -8)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(track_color)
        painter.drawRoundedRect(track_rect, track_rect.height() / 2, track_rect.height() / 2)
        thumb_radius = self.height() / 2 - 4
        thumb_x = track_rect.left() + self._thumb_pos * (track_rect.width() - thumb_radius * 2)
        painter.setBrush(thumb_color)
        painter.drawEllipse(int(thumb_x), track_rect.top(), int(thumb_radius * 2), int(thumb_radius * 2))

class SettingsDialog(QDialog):
    darkModeChanged = pyqtSignal(bool)
    languageChanged = pyqtSignal(str)
    autoSaveChanged = pyqtSignal(int)
    backupPathChanged = pyqtSignal(str)

    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle(self.tr("Settings"))
        self.setFixedSize(500, 300)
        layout = QFormLayout(self)
        self.dark_mode_switch = MaterialSwitch()
        self.dark_mode_switch.setChecked(self.settings.value("dark_mode", True, type=bool))
        self.dark_mode_switch.toggled.connect(self.darkModeChanged.emit)
        layout.addRow(QLabel(self.tr("Dark Mode")), self.dark_mode_switch)
        self.language_combo = QComboBox()
        self.language_combo.addItem(self.tr("English"), "en")
        self.language_combo.addItem(self.tr("Marathi"), "mr")
        current_locale = self.settings.value("language", "en")
        self.language_combo.setCurrentIndex(self.language_combo.findData(current_locale))
        self.language_combo.currentIndexChanged.connect(lambda index: self.languageChanged.emit(self.language_combo.itemData(index)))
        layout.addRow(QLabel(self.tr("Software Language")), self.language_combo)
        self.autosave_spinbox = QSpinBox()
        self.autosave_spinbox.setRange(1, 60)
        self.autosave_spinbox.setSuffix(" " + self.tr("minutes"))
        self.autosave_spinbox.setValue(self.settings.value("auto_save_interval", 5, type=int))
        self.autosave_spinbox.valueChanged.connect(self.autoSaveChanged.emit)
        layout.addRow(QLabel(self.tr("Auto-Save Interval (minutes)")), self.autosave_spinbox)
        self.backup_path_edit = QLineEdit(self.settings.value("backup_location", ""))
        self.backup_path_edit.setPlaceholderText(self.tr("No backup path set"))
        self.backup_path_edit.textChanged.connect(self.backupPathChanged.emit)
        browse_button = QPushButton(self.tr("Choose Location"))
        browse_button.clicked.connect(self.choose_backup_location)
        h_layout = QHBoxLayout()
        h_layout.addWidget(self.backup_path_edit)
        h_layout.addWidget(browse_button)
        layout.addRow(QLabel(self.tr("Backup & Export Location")), h_layout)

    def tr(self, text):
        return QApplication.instance().tr(text)

    def choose_backup_location(self):
        path = QFileDialog.getExistingDirectory(self, self.tr("Choose Location"), self.backup_path_edit.text())
        if path:
            self.backup_path_edit.setText(path)

    def retranslate(self):
        self.setWindowTitle(self.tr("Settings"))
        self.dark_mode_switch.update()
        self.language_combo.setItemText(0, self.tr("English"))
        self.language_combo.setItemText(1, self.tr("Marathi"))
        self.autosave_spinbox.setSuffix(" " + self.tr("minutes"))
        self.findChild(QLabel, self.tr("Dark Mode")).setText(self.tr("Dark Mode"))
        self.findChild(QLabel, self.tr("Software Language")).setText(self.tr("Software Language"))
        self.findChild(QLabel, self.tr("Auto-Save Interval (minutes)")).setText(self.tr("Auto-Save Interval (minutes)"))
        self.findChild(QLabel, self.tr("Backup & Export Location")).setText(self.tr("Backup & Export Location"))
        self.backup_path_edit.setPlaceholderText(self.tr("No backup path set"))
        self.findChild(QPushButton, self.tr("Choose Location")).setText(self.tr("Choose Location"))