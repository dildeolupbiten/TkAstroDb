# -*- coding: utf-8 -*-

from .frame import Frame
from .libs import Qt, QListWidget, QLabel


class Select(Frame):
    def __init__(self, items, defaults: list = None, selected: str = "", *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.__items = items
        self.list = QListWidget(self)
        self.list.setStyleSheet("background-color: #343a40; color: white;")
        self.list.setSelectionMode(3)
        self.list.addItems(items)
        self.label = QLabel(self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setText(f"Selected{selected}: 0")
        self.label.setFixedWidth(400)
        self.list.itemSelectionChanged.connect(
            lambda: self.label.setText(f"Selected{selected}: {len(self.list.selectedItems())}")
        )
        self.box.addWidget(self.list)
        self.box.addWidget(self.label, alignment=Qt.AlignCenter)
        if defaults:
            self.set(defaults)

    @property
    def items(self):
        return self.__items

    @items.setter
    def items(self, items):
        self.__items = items
        self.list.clear()
        self.list.addItems(items)

    def get(self):
        return list(map(lambda i: i.text(), self.list.selectedItems()))

    def set(self, defaults):
        for item in defaults:
            self.list.item(self.__items.index(item)).setSelected(True)
