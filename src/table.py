#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from .libs import (
    sys, QApplication, QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout,
    QAbstractScrollArea, QFrame, QPushButton, QLabel, QComboBox, QLineEdit, QDialog, QHeaderView, Qt, QColor
)


class Table(QDialog):
    def __init__(self, columns: list = None, data: list = None, widget=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle("Select Records")
        self.setStyleSheet("background-color: #212529; color: white;")
        self.data = data
        self.widget = widget
        self.selected = []
        self.row_count, self.col_count = 0, 0
        self.columns = columns
        self.box = QVBoxLayout(self)
        self.setLayout(self.box)
        self.title = QLabel(self)
        self.title.setText("Select Records")
        self.title.setStyleSheet("font: 20px Monospace;")
        self.frame = QFrame(self)
        self.frame.setStyleSheet("background-color: #212529; color: white;")
        self.frame.box = QHBoxLayout(self.frame)
        self.frame.setLayout(self.frame.box)
        self.frame.left = QFrame(self)
        self.frame.left.box = QVBoxLayout(self.frame.left)
        self.frame.left.frames = {}
        self.frame.left.label = QLabel(self.frame.left)
        self.frame.left.label.setText("Column")
        self.frame.left.combobox = QComboBox(self.frame.left)
        self.frame.left.combobox.setStyleSheet(
            """
            QComboBox {
                background-color: #343a40;
                font: Monospace 15;
            }
            QComboBox QAbstractItemView {
              border: none;
              background-color: #212529;
              selection-background-color: blue;
            }
            """
        )
        self.frame.left.combobox.addItems(columns)
        self.frame.left.combobox.setFixedWidth(200)
        self.frame.left.box.addWidget(self.frame.left.label, alignment=Qt.AlignCenter)
        self.frame.left.box.addWidget(self.frame.left.combobox, alignment=Qt.AlignCenter)
        self.frame.box.addWidget(self.frame.left)
        self.frame.right = QFrame(self)
        self.frame.right.box = QVBoxLayout(self.frame.right)
        self.frame.right.label = QLabel(self.frame.right)
        self.frame.right.label.setText("Filter By Value")
        self.frame.right.input = QLineEdit(self.frame.right)
        self.frame.right.input.setStyleSheet("background-color: #343a40;")
        self.frame.right.input.setFixedWidth(200)
        self.frame.right.input.textEdited.connect(self.filter_data)
        self.frame.right.box.addWidget(self.frame.right.label, alignment=Qt.AlignCenter)
        self.frame.right.box.addWidget(self.frame.right.input, alignment=Qt.AlignCenter)
        self.frame.box.addWidget(self.frame.right)
        self.table = QTableWidget(self)
        self.table.setStyleSheet("background-color: #343a40;")
        self.table.itemSelectionChanged.connect(self.change_selection)
        self.table.horizontalHeader().setStyleSheet("background: #343a40;")
        self.result = QLabel(self)
        self.result.setText("Selected Records: 0")
        self.result.setStyleSheet("margin-top: 10px;")
        self.box.addWidget(self.title, alignment=Qt.AlignCenter)
        self.box.addWidget(self.frame)
        self.box.addWidget(self.table)
        self.box.addWidget(self.result, alignment=Qt.AlignCenter)
        if data:
            self.add_data()

    def add_data(self):
        self.row_count = len(self.data)
        self.col_count = len(self.columns)
        self.table.setColumnCount(self.col_count)
        self.table.setRowCount(self.row_count)
        for index, column in enumerate(self.columns):
            item = QTableWidgetItem(column)
            item.setBackground(QColor(0, 255, 0))
            self.table.setHorizontalHeaderItem(index, item)
        self.table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #343a40 }")
        self.table.verticalHeader().setStyleSheet("QHeaderView::section { background-color: #343a40 }")

        self.table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        for row in range(self.row_count):
            for col in range(self.col_count):
                item = QTableWidgetItem(str(self.data[row][col]))
                self.table.setItem(row, col, item)
                self.table.setColumnWidth(col, 90)

    def filter_data(self, text):
        col_index = self.columns.index(self.frame.left.combobox.currentText())
        for row in range(self.row_count):
            if not str(self.data[row][col_index]).startswith(text):
                self.table.setRowHidden(row, True)
            else:
                self.table.setRowHidden(row, False)

    def change_selection(self):
        self.selected = list(
            map(lambda i: self.data[i], sorted(set(item.row() for item in self.table.selectedItems())))
        )
        self.result.setText(f"Selected Records: {len(self.selected)}")
        self.widget.setText(f"Selected Records{' ' * 3}: {len(self.selected)}")
