from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QTabWidget, QScrollArea, QFrame,
    QFormLayout, QLabel, QLineEdit, QDateEdit, QComboBox, QTextEdit,
    QPushButton, QHBoxLayout, QApplication
)
from PyQt6.QtCore import QDate, Qt, pyqtSignal
from PyQt6.QtGui import QDoubleValidator

from .construction_items import ConstructionItemsWidget
from .excess_saving import ExcessSavingWidget
from .dialogs import MessageEditorDialog, QDialog

class MergedFormWidget(QWidget):
    something_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.inputs = {}
        self.is_dirty = False
        self.message_text = ""
        self.setup_ui()
        self.update_styles(True)

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)
        details_page_scroll = QScrollArea()
        details_page_scroll.setWidgetResizable(True)
        details_page = QWidget()
        details_page_scroll.setWidget(details_page)
        tab_widget.addTab(details_page_scroll, self.tr("Document Details"))
        details_layout = QVBoxLayout(details_page)
        form_layout = QFormLayout()
        form_layout.setVerticalSpacing(15)
        details_layout.addLayout(form_layout)
        fields_a = ["name", "name_work", "division", "constituency", "fund_head", "contractor", "deputy_engineer",
                    "date", "start_date", "end_date", "agreement_no", "work_order_no",
                    "acceptance_no", "mb_no", "letter_no", "vide_letter_no", "year",
                    "est_cost", "amt_rupes", "percentage_quoted", "send_to", "subject"]
        for field in fields_a:
            widget = None
            if field == "division":
                widget = QComboBox()
                widget.addItems(["West", "City", "East"])
                widget.currentIndexChanged.connect(self.set_dirty)
            elif "date" in field:
                widget = QDateEdit(calendarPopup=True, date=QDate.currentDate())
                widget.dateChanged.connect(self.set_dirty)
            elif field in ["est_cost", "amt_rupes", "percentage_quoted"]:
                widget = QLineEdit()
                widget.setValidator(QDoubleValidator())
                widget.textChanged.connect(self.set_dirty)
            else:
                widget = QLineEdit()
                widget.textChanged.connect(self.set_dirty)
            if widget:
                self.inputs[field] = widget
                label_text = self.tr(field.replace("_", " ").title())
                label = QLabel(label_text + ":")
                label.setObjectName(f"label_{field}")
                form_layout.addRow(label, widget)
        message_label = QLabel(self.tr("Message") + ":")
        message_label.setObjectName("message_label")
        form_layout.addRow(message_label)
        self.message_preview = QTextEdit()
        self.message_preview.setReadOnly(True)
        self.message_preview.setMaximumHeight(120)
        self.message_preview.setToolTip(self.tr("This is a preview. Click the button to edit the full message."))
        self.edit_message_btn = QPushButton(self.tr("Edit Message..."))
        self.edit_message_btn.clicked.connect(self.open_message_editor)
        message_layout = QVBoxLayout()
        message_layout.addWidget(self.message_preview)
        message_layout.addWidget(self.edit_message_btn, 0, Qt.AlignmentFlag.AlignRight)
        form_layout.addRow(message_layout)
        self.construction_items_widget = ConstructionItemsWidget()
        tab_widget.addTab(self.construction_items_widget, self.tr("Construction Items"))
        self.construction_items_widget.dirty_state_changed.connect(self.set_dirty)
        self.excess_saving_widget = ExcessSavingWidget()
        tab_widget.addTab(self.excess_saving_widget, self.tr("Excess/Saving Statement"))
        self.excess_saving_widget.dirty_state_changed.connect(self.set_dirty)
        self.construction_items_widget.rows_changed.connect(self.sync_excess_saving_table)

    def retranslate(self):
        self.findChild(QTabWidget).setTabText(0, self.tr("Document Details"))
        self.findChild(QTabWidget).setTabText(1, self.tr("Construction Items"))
        self.findChild(QTabWidget).setTabText(2, self.tr("Excess/Saving Statement"))
        for field, widget in self.inputs.items():
            label = self.findChild(QLabel, f"label_{field}")
            if label:
                label.setText(self.tr(field.replace("_", " ").title()) + ":")
        message_label = self.findChild(QLabel, "message_label")
        if message_label:
            message_label.setText(self.tr("Message") + ":")
        self.edit_message_btn.setText(self.tr("Edit Message..."))
        self.construction_items_widget.retranslate()
        self.excess_saving_widget.retranslate()

    def tr(self, text):
        return QApplication.instance().tr(text)

    def open_message_editor(self):
        dialog = MessageEditorDialog(self.message_text, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_text = dialog.get_text()
            if new_text != self.message_text:
                self.message_text = new_text
                self.message_preview.setPlainText(self.message_text)
                self.set_dirty()
        
    def sync_excess_saving_table(self):
        existing_excess_data = self.excess_saving_widget.gather_data()
        construction_data = self.construction_items_widget.gather_data()
        items_data = construction_data.get('items', [])
        for item in items_data:
            sr_no = item.get("sr_no")
            if sr_no in existing_excess_data:
                item['executed_quantity'] = existing_excess_data[sr_no].get('executed_quantity', item.get('quantity', '0'))
                item['remarks_excess_saving'] = existing_excess_data[sr_no].get('remarks_excess_saving', "As Per Site Condition")
        self.excess_saving_widget.update_table(items_data)

    def set_dirty(self):
        self.is_dirty = True
        self.something_changed.emit()

    def clear_dirty(self):
        self.is_dirty = False

    def gather_data(self):
        data = {}
        for key, widget in self.inputs.items():
            if isinstance(widget, QLineEdit): data[key] = widget.text()
            elif isinstance(widget, QDateEdit): data[key] = widget.date().toString("dd-MM-yyyy")
            elif isinstance(widget, QComboBox): data[key] = widget.currentText()
        data["message"] = self.message_text
        construction_data = self.construction_items_widget.gather_data()
        data.update(construction_data)
        excess_data = self.excess_saving_widget.gather_data()
        for item in data.get("items", []):
            sr_no = item.get("sr_no")
            if sr_no in excess_data:
                item.update(excess_data[sr_no])
        return data

    def load_data(self, data):
        for key, widget in self.inputs.items():
            if key in data:
                val = data.get(key, "")
                if isinstance(widget, QLineEdit): widget.setText(val)
                elif isinstance(widget, QDateEdit): widget.setDate(QDate.fromString(val, "dd-MM-yyyy"))
                elif isinstance(widget, QComboBox): widget.setCurrentText(val)
        self.message_text = data.get("message", "")
        self.message_preview.setPlainText(self.message_text)
        self.construction_items_widget.load_data(data)
        self.clear_dirty()

    def clear_form(self):
        for widget in self.inputs.values():
            if isinstance(widget, QLineEdit): widget.clear()
            elif isinstance(widget, QDateEdit): widget.setDate(QDate.currentDate())
            elif isinstance(widget, QComboBox): widget.setCurrentIndex(0)
        self.message_text = ""
        self.message_preview.clear()
        self.construction_items_widget.clear_form()
        self.clear_dirty()

    def update_styles(self, dark_mode):
        style = """
            QWidget {{ background-color: {bg}; color: {fg}; }} QLabel {{ color: {fg}; }}
            QTabWidget::pane {{ border: 1px solid {border}; }}
            QTabBar::tab {{ background: {bg_alt}; color: {fg}; padding: 10px; }}
            QTabBar::tab:selected {{ background: {bg}; border-top: 2px solid #3F51B5; }}
            QLineEdit, QDateEdit, QTextEdit, QComboBox, QSpinBox {{ padding: 10px; border: 1px solid {border}; border-radius: 4px; font-size: 14px; background-color: {bg_alt}; color: {fg}; min-height: 20px; }}
            QLineEdit:focus, QDateEdit:focus, QTextEdit:focus, QComboBox:focus, QSpinBox:focus {{ border: 2px solid #3F51B5; }}
            QTableWidget {{ background-color: {bg_alt}; color: {fg}; gridline-color: {border}; }}
            QHeaderView::section {{ background-color: {bg}; color: {fg}; padding: 5px; border: 1px solid {border}; }}
            QPushButton {{ padding: 5px; background-color: #3F51B5; color: white; border: none; border-radius: 4px; }}
            QPushButton:hover {{ background-color: #303F9F; }}
            QFrame, QScrollArea {{ background-color: transparent; border: none; color: {fg}; }}
            
            QComboBox QAbstractItemView {{
                border: 1px solid {border};
                background-color: {bg_card};
                color: {fg};
                selection-background-color: #3F51B5;
                selection-color: {fg};
            }}
        """.format(
            bg="#252525" if dark_mode else "#FFFFFF",
            fg="#FFFFFF" if dark_mode else "#212121",
            bg_alt="#353535" if dark_mode else "#F8F8F8",
            border="#444" if dark_mode else "#CFD8DC",
            bg_card="#252525" if dark_mode else "#FFFFFF"
        )
        self.setStyleSheet(style)