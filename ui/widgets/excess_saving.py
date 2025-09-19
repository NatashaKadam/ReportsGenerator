from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QTableWidget, QHeaderView, QTableWidgetItem,
    QApplication
)
from PyQt6.QtCore import Qt, pyqtSignal

class ExcessSavingWidget(QWidget):
    dirty_state_changed = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._is_updating = False
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            self.tr("Item No."), self.tr("Tender Qty"), self.tr("Executed Qty"), self.tr("Unit"), 
            self.tr("Description of Item"), self.tr("Excess"), self.tr("Saving"), self.tr("Remarks")
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 60)
        self.table.setColumnWidth(1, 90)
        self.table.setColumnWidth(2, 90)
        self.table.setColumnWidth(3, 60)
        self.table.setColumnWidth(5, 70)
        self.table.setColumnWidth(6, 70)
        self.table.setColumnWidth(7, 150)
        self.table.verticalHeader().setVisible(False)
        self.table.itemChanged.connect(self._on_item_changed)
        layout.addWidget(self.table)
        
    def retranslate(self):
        self.table.setHorizontalHeaderLabels([
            self.tr("Item No."), self.tr("Tender Qty"), self.tr("Executed Qty"), self.tr("Unit"), 
            self.tr("Description of Item"), self.tr("Excess"), self.tr("Saving"), self.tr("Remarks")
        ])

    def tr(self, text):
        return QApplication.instance().tr(text)

    def _calculate_and_set_diff(self, row):
        tender_item = self.table.item(row, 1)
        executed_item = self.table.item(row, 2)
        excess_item = self.table.item(row, 5)
        saving_item = self.table.item(row, 6)
        if not all([tender_item, executed_item, excess_item, saving_item]):
            return
        try:
            tender_qty = float(tender_item.text())
            executed_qty = float(executed_item.text())
            diff = executed_qty - tender_qty
            if abs(diff) < 1e-9:
                excess_item.setText("-")
                saving_item.setText("-")
            elif diff > 0:
                excess_item.setText(f"{diff:.2f}")
                saving_item.setText("-")
            else:
                excess_item.setText("-")
                saving_item.setText(f"{-diff:.2f}")
        except (ValueError, TypeError):
            excess_item.setText("-")
            saving_item.setText("-")

    def _on_item_changed(self, item):
        if self._is_updating:
            return
        if item.column() == 2:
            self._calculate_and_set_diff(item.row())
            self.dirty_state_changed.emit()
        elif item.column() == 7:
            self.dirty_state_changed.emit()

    def update_table(self, items_data):
        self._is_updating = True
        self.table.setRowCount(0)
        for item_data in items_data:
            row = self.table.rowCount()
            self.table.insertRow(row)
            sr_no_item = QTableWidgetItem(item_data.get("sr_no", ""))
            tender_qty_item = QTableWidgetItem(item_data.get("quantity", "0"))
            executed_qty_val = item_data.get("executed_quantity", item_data.get("quantity", "0"))
            executed_qty_item = QTableWidgetItem(str(executed_qty_val))
            unit_item = QTableWidgetItem(item_data.get("unit", ""))
            desc_item = QTableWidgetItem(item_data.get("description", ""))
            excess_item = QTableWidgetItem("-")
            saving_item = QTableWidgetItem("-")
            remarks_val = item_data.get("remarks_excess_saving", "As Per Site Condition")
            remarks_item = QTableWidgetItem(remarks_val)
            for itm in [sr_no_item, tender_qty_item, unit_item, desc_item, excess_item, saving_item]:
                itm.setFlags(itm.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, sr_no_item)
            self.table.setItem(row, 1, tender_qty_item)
            self.table.setItem(row, 2, executed_qty_item)
            self.table.setItem(row, 3, unit_item)
            self.table.setItem(row, 4, desc_item)
            self.table.setItem(row, 5, excess_item)
            self.table.setItem(row, 6, saving_item)
            self.table.setItem(row, 7, remarks_item)
            self._calculate_and_set_diff(row)
        self._is_updating = False

    def gather_data(self):
        data = {}
        for row in range(self.table.rowCount()):
            sr_no = self.table.item(row, 0).text()
            if sr_no:
                data[sr_no] = {
                    "executed_quantity": self.table.item(row, 2).text(),
                    "excess": self.table.item(row, 5).text(),
                    "saving": self.table.item(row, 6).text(),
                    "remarks_excess_saving": self.table.item(row, 7).text()
                }
        return data

    def clear_form(self):
        self._is_updating = True
        self.table.setRowCount(0)
        self._is_updating = False
        self.dirty_state_changed.emit()