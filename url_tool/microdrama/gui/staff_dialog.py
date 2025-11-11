from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QGridLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from ..core.staff_db import delete_staff, list_staff, upsert_staff


class StaffDialog(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("审核人员库管理")
        self.resize(500, 400)
        self._build_ui()
        self._load_staff()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        self.table = QTableWidget(0, 2, self)
        self.table.setHorizontalHeaderLabels(["姓名", "身份证后四位"])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)

        layout.addWidget(self.table)

        button_layout = QGridLayout()
        self.add_button = QPushButton("新增")
        self.delete_button = QPushButton("删除")
        self.refresh_button = QPushButton("刷新")

        self.add_button.clicked.connect(self._add_staff)
        self.delete_button.clicked.connect(self._delete_staff)
        self.refresh_button.clicked.connect(self._load_staff)

        button_layout.addWidget(self.add_button, 0, 0)
        button_layout.addWidget(self.delete_button, 0, 1)
        button_layout.addWidget(self.refresh_button, 0, 2)
        layout.addLayout(button_layout)

    def _load_staff(self) -> None:
        staff = list_staff()
        self.table.setRowCount(len(staff))
        for row_index, (name, id_last4) in enumerate(staff.items()):
            self.table.setItem(row_index, 0, QTableWidgetItem(name))
            self.table.setItem(row_index, 1, QTableWidgetItem(id_last4))

    def _add_staff(self) -> None:
        name, ok = QInputDialog.getText(self, "新增人员", "请输入姓名")
        if not ok or not name.strip():
            return
        id_last4, ok = QInputDialog.getText(self, "新增人员", "请输入身份证后四位")
        if not ok or not id_last4.strip():
            return
        id_last4 = id_last4.strip()
        if len(id_last4) != 4 or not id_last4.isdigit():
            QMessageBox.warning(self, "提示", "身份证后四位必须为4位数字")
            return
        upsert_staff(name.strip(), id_last4)
        self._load_staff()

    def _delete_staff(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "提示", "请先选择要删除的人员")
            return
        name_item = self.table.item(row, 0)
        if not name_item:
            return
        name = name_item.text()
        if delete_staff(name):
            self._load_staff()
        else:
            QMessageBox.warning(self, "提示", "删除失败，未找到该人员")
