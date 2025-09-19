import os
import pandas as pd
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QLineEdit, QLabel, QPushButton,
    QHBoxLayout, QTableWidget, QTableWidgetItem, QHeaderView, QCompleter,
    QFrame, QComboBox, QApplication
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDoubleValidator
import re

from .dialogs import show_message_box
from core.constants import SSR_DATA_EXCEL

class ConstructionItemsWidget(QWidget):
    dirty_state_changed = pyqtSignal()
    rows_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.ssr_data = None
        self.setup_ui()
        self.load_ssr_data_from_excel()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0,0,0,0)
        
        item_entry_frame = QFrame()
        item_entry_layout = QFormLayout(item_entry_frame)
        item_entry_layout.setVerticalSpacing(10)

        self.description_combo = QComboBox()
        self.description_combo.setEditable(True)
        self.description_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        item_entry_layout.addRow(QLabel(self.tr("Item Description:")), self.description_combo)

        self.quantity_input = QLineEdit()
        self.quantity_input.setValidator(QDoubleValidator())
        item_entry_layout.addRow(QLabel(self.tr("Quantity:")), self.quantity_input)

        self.unit_input = QLineEdit(readOnly=True)
        item_entry_layout.addRow(QLabel(self.tr("Unit:")), self.unit_input)

        self.rate_input = QLineEdit(readOnly=True)
        item_entry_layout.addRow(QLabel(self.tr("Rate:")), self.rate_input)

        self.total_label = QLabel("₹0.00")
        self.total_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        item_entry_layout.addRow(QLabel(self.tr("Total Cost:")), self.total_label)
        
        button_layout = QHBoxLayout()
        self.add_button = QPushButton(self.tr("Add Item"))
        self.add_button.setToolTip(self.tr("Add the defined item to the table below."))
        button_layout.addWidget(self.add_button)
        
        signatories_frame = QFrame()
        signatories_layout = QFormLayout(signatories_frame)
        signatories_layout.setVerticalSpacing(10)
        self.jr_engineer_input = QLineEdit()
        self.deputy_engineer_input = QLineEdit()
        self.executive_engineer_input = QLineEdit()
        signatories_layout.addRow(QLabel(self.tr("Jr./Sect./Asst. Engineer:")), self.jr_engineer_input)
        signatories_layout.addRow(QLabel(self.tr("Deputy Engineer:")), self.deputy_engineer_input)
        signatories_layout.addRow(QLabel(self.tr("Executive Engineer:")), self.executive_engineer_input)

        self.items_table = QTableWidget()
        self.items_table.setColumnCount(11) 
        self.items_table.setHorizontalHeaderLabels([self.tr("Sr. No"), self.tr("Chapter"), self.tr("SSR Item No."), self.tr("Reference No."), self.tr("Description"), self.tr("Add. Spec."), self.tr("Unit"), self.tr("Rate"), self.tr("Qty"), self.tr("Total"), self.tr("Actions")])
        
        header = self.items_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch) 
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch) 
        
        for i, width in enumerate([40, 60, 80, 80, 250, 150, 60, 90, 60, 90, 70]):
            self.items_table.setColumnWidth(i, width)
        
        self.items_table.verticalHeader().setVisible(False)

        layout.addWidget(item_entry_frame)
        layout.addLayout(button_layout)
        layout.addWidget(self.items_table)
        layout.addWidget(signatories_frame)

        self.description_combo.currentTextChanged.connect(self.update_item_details)
        self.quantity_input.textChanged.connect(self.calculate_total)
        self.add_button.clicked.connect(self.add_to_table)
        
        self.quantity_input.textChanged.connect(self.dirty_state_changed.emit)
        self.description_combo.currentTextChanged.connect(self.dirty_state_changed.emit)
        self.jr_engineer_input.textChanged.connect(self.dirty_state_changed.emit)
        self.deputy_engineer_input.textChanged.connect(self.dirty_state_changed.emit)
        self.executive_engineer_input.textChanged.connect(self.dirty_state_changed.emit)
        self.items_table.itemChanged.connect(self.dirty_state_changed.emit)

    def tr(self, text):
        return QApplication.instance().tr(text)

    def load_ssr_data_from_excel(self):
        try:
            if not os.path.exists(SSR_DATA_EXCEL):
                show_message_box(self.tr("Excel Data Missing"), self.tr(f"Error: '{SSR_DATA_EXCEL}' not found.\nPlease ensure the file exists in the 'assets' folder."))
                return

            excel_data = pd.read_excel(SSR_DATA_EXCEL, sheet_name=None, header=1)

            self.ssr_data = pd.DataFrame()
            for df in excel_data.values():
                self.ssr_data = pd.concat([self.ssr_data, df], ignore_index=True)

            expected_excel_cols = {
                'Sr. No': 'sr_no', 'Chapter': 'chapter', 'SSR Item No.': 'ssr_item_no',
                'Reference No.': 'reference_no', 'Description of the item': 'description_of_the_item',
                'Additional Specification': 'additional_specification', 'Unit': 'unit',
                'Completed Rates': 'completed_rates'
            }

            self.ssr_data = self.ssr_data.rename(columns={k: v for k, v in expected_excel_cols.items() if k in self.ssr_data.columns})
            
            required_cols_standardized = ['description_of_the_item', 'unit', 'completed_rates', 'ssr_item_no']
            if not all(col in self.ssr_data.columns for col in required_cols_standardized):
                show_message_box(self.tr("Invalid Excel File"), self.tr("Excel file must contain required columns."))
                self.ssr_data = None
                return

            descriptions = self.ssr_data['description_of_the_item'].dropna().unique().tolist()
            self.description_combo.clear()
            self.description_combo.addItems(descriptions)

            completer = QCompleter(descriptions, self)
            completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            completer.setFilterMode(Qt.MatchFlag.MatchContains)
            self.description_combo.setCompleter(completer)

        except Exception as e:
            show_message_box(self.tr("Excel Load Error"), self.tr(f"An error occurred while reading the Excel file: {e}"))
            self.ssr_data = None

    def update_item_details(self, description):
        self.unit_input.clear()
        self.rate_input.clear()
        self.calculate_total()
        if self.ssr_data is None or not description: return
        item_data = self.ssr_data[self.ssr_data['description_of_the_item'] == description]
        if not item_data.empty:
            item = item_data.iloc[0]
            self.unit_input.setText(str(item.get('unit', '')))
            rate_val = item.get('completed_rates', 0.0)
            self.rate_input.setText(f"₹{float(rate_val):,.2f}" if pd.notna(rate_val) else "₹0.00")
        self.calculate_total()

    def calculate_total(self):
        try:
            quantity = float(self.quantity_input.text() or 0)
            rate = float(self.rate_input.text().replace("₹", "").replace(",", "") or 0)
            self.total_label.setText(f"₹{quantity * rate:,.2f}")
        except ValueError:
            self.total_label.setText("₹0.00")
        self.dirty_state_changed.emit()

    def add_to_table(self):
        if self.ssr_data is None: return show_message_box(self.tr("Data Not Loaded"), self.tr("SSR data not available."))
        description = self.description_combo.currentText()
        if not description: return show_message_box(self.tr("Invalid Item"), self.tr("Please select a valid item."))
        try:
            if float(self.quantity_input.text() or 0) <= 0: return show_message_box(self.tr("Invalid Input"), self.tr("Enter a positive quantity."))
        except ValueError: return show_message_box(self.tr("Invalid Quantity"), self.tr("Quantity must be a number."))
        item_data = self.ssr_data[self.ssr_data['description_of_the_item'] == description]
        if item_data.empty: return show_message_box(self.tr("Item Not Found"), self.tr("Selected item not in data source."))

        ssr_item = item_data.iloc[0]
        row = self.items_table.rowCount()
        self.items_table.insertRow(row)

        self.items_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self.items_table.setItem(row, 1, QTableWidgetItem(str(ssr_item.get('chapter', ''))))
        self.items_table.setItem(row, 2, QTableWidgetItem(str(ssr_item.get('ssr_item_no', ''))))
        self.items_table.setItem(row, 3, QTableWidgetItem(str(ssr_item.get('reference_no', ''))))
        self.items_table.setItem(row, 4, QTableWidgetItem(description))
        self.items_table.setItem(row, 5, QTableWidgetItem(str(ssr_item.get('additional_specification', ''))))
        self.items_table.setItem(row, 6, QTableWidgetItem(self.unit_input.text()))
        self.items_table.setItem(row, 7, QTableWidgetItem(self.rate_input.text()))
        self.items_table.setItem(row, 8, QTableWidgetItem(self.quantity_input.text()))
        self.items_table.setItem(row, 9, QTableWidgetItem(self.total_label.text()))

        delete_button = QPushButton(self.tr("Delete"))
        delete_button.clicked.connect(lambda _, r=row: self.remove_table_row(r))
        self.items_table.setCellWidget(row, 10, delete_button)

        self.dirty_state_changed.emit()
        self.rows_changed.emit()
        self.clear_entry_fields()

    def remove_table_row(self, row):
        self.items_table.removeRow(row)
        for r in range(self.items_table.rowCount()):
            self.items_table.item(r, 0).setText(str(r+1))
            delete_button = self.items_table.cellWidget(r, 10)
            if delete_button:
                try: delete_button.clicked.disconnect()
                except TypeError: pass
                delete_button.clicked.connect(lambda _, current_r=r: self.remove_table_row(current_r))

        self.dirty_state_changed.emit()
        self.rows_changed.emit()

    def clear_entry_fields(self):
        self.description_combo.setCurrentIndex(-1)
        self.description_combo.clearEditText()
        self.quantity_input.clear()
        self.unit_input.clear()
        self.rate_input.clear()
        self.total_label.setText("₹0.00")

    def clear_form(self):
        self.clear_entry_fields()
        self.items_table.setRowCount(0)
        self.jr_engineer_input.clear()
        self.deputy_engineer_input.clear()
        self.executive_engineer_input.clear()
        self.dirty_state_changed.emit()
        self.rows_changed.emit()

    def gather_data(self):
        items = []
        for row in range(self.items_table.rowCount()):
            items.append({
                "sr_no": self.items_table.item(row, 0).text(), "chapter": self.items_table.item(row, 1).text(),
                "ssr_no": self.items_table.item(row, 2).text(), "reference_no": self.items_table.item(row, 3).text(),
                "description": self.items_table.item(row, 4).text(), "additional_spec": self.items_table.item(row, 5).text(),
                "unit": self.items_table.item(row, 6).text(), "unit_rate": self.items_table.item(row, 7).text(),
                "quantity": self.items_table.item(row, 8).text(), "total": self.items_table.item(row, 9).text()
            })
        
        # Robustly handle potential non-numeric characters in the 'total' field
        def extract_float(s):
            if not s:
                return 0.0
            numeric_str = "".join(re.findall(r'[\d.]+', s.replace(",", "")))
            try:
                return float(numeric_str) if numeric_str else 0.0
            except ValueError:
                return 0.0
        
        overall_total_amount = sum(extract_float(i.get("total")) for i in items if i.get("total"))
        
        return {
            "items": items, "total_amount": f"₹{overall_total_amount:,.2f}",
            "signatory_jr_engineer": self.jr_engineer_input.text(),
            "signatory_deputy_engineer": self.deputy_engineer_input.text(),
            "signatory_exec_engineer": self.executive_engineer_input.text()
        }

    def load_data(self, data):
        self.items_table.setRowCount(0)
        items_to_load = data.get("items", [])
        for entry in items_to_load:
            row = self.items_table.rowCount()
            self.items_table.insertRow(row)
            self.items_table.setItem(row, 0, QTableWidgetItem(entry.get("sr_no", "")))
            self.items_table.setItem(row, 1, QTableWidgetItem(entry.get("chapter", "")))
            self.items_table.setItem(row, 2, QTableWidgetItem(entry.get("ssr_no", "")))
            self.items_table.setItem(row, 3, QTableWidgetItem(entry.get("reference_no", "")))
            self.items_table.setItem(row, 4, QTableWidgetItem(entry.get("description", "")))
            self.items_table.setItem(row, 5, QTableWidgetItem(entry.get("additional_spec", "")))
            self.items_table.setItem(row, 6, QTableWidgetItem(entry.get("unit", "")))
            self.items_table.setItem(row, 7, QTableWidgetItem(str(entry.get("unit_rate", ""))))
            self.items_table.setItem(row, 8, QTableWidgetItem(str(entry.get("quantity", ""))))
            self.items_table.setItem(row, 9, QTableWidgetItem(str(entry.get("total", ""))))
            delete_button = QPushButton(self.tr("Delete"))
            delete_button.clicked.connect(lambda _, r=row: self.remove_table_row(r))
            self.items_table.setCellWidget(row, 10, delete_button)
        self.jr_engineer_input.setText(data.get("signatory_jr_engineer", ""))
        self.deputy_engineer_input.setText(data.get("signatory_deputy_engineer", ""))
        self.executive_engineer_input.setText(data.get("signatory_exec_engineer", ""))
        self.dirty_state_changed.emit()
        self.rows_changed.emit()

    def retranslate(self):
        self.add_button.setText(self.tr("Add Item"))
        self.add_button.setToolTip(self.tr("Add the defined item to the table below."))
        self.items_table.setHorizontalHeaderLabels([self.tr("Sr. No"), self.tr("Chapter"), self.tr("SSR Item No."), self.tr("Reference No."), self.tr("Description"), self.tr("Add. Spec."), self.tr("Unit"), self.tr("Rate"), self.tr("Qty"), self.tr("Total"), self.tr("Actions")])
        for row in range(self.items_table.rowCount()):
            delete_button = self.items_table.cellWidget(row, 10)
            if delete_button:
                delete_button.setText(self.tr("Delete"))