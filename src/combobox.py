# -*- coding: utf-8 -*-

from .consts import AYANAMSHA, CALC_TYPES
from .frame import Frame, ScrollableFrame
from .libs import QComboBox, QHBoxLayout, QLabel, QFrame


class Combobox(ScrollableFrame):
    def __init__(self, values, defaults: list = None, other: dict = None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.combobox = QComboBox(self)
        self.combobox.setStyleSheet("""
        QComboBox {
            background: #343a40;
            font: Monospace 15;
        }
        QComboBox QAbstractItemView {
          border: none;
          background: #212529;
          selection-background-color: blue;
        }
        """)
        self.combobox.addItems(values)
        self.sub_frame.addWidget(self.combobox)
        if other:
            self.other_frame, self.other_combobox = self.create_frame(other)
            self.other_frame.setVisible(False)
        self.combobox.currentTextChanged.connect(lambda e: self.add_frame())
        if defaults:
            self.set(defaults)

    def set(self, defaults):
        self.combobox.setCurrentText(defaults[0])
        if len(defaults) > 1:
            self.other_combobox.setCurrentText(defaults[1])

    def get(self):
        return self.combobox.currentText()

    def get_other(self):
        return self.other_combobox.currentText()

    def add_frame(self):
        text = self.combobox.currentText()
        if text in ["Sidereal", CALC_TYPES[0]]:
            self.other_frame.setVisible(True)
        else:
            self.other_frame.setVisible(False)

    def create_frame(self, other):
        frame = QFrame(self)
        layout = QHBoxLayout()
        frame.setLayout(layout)
        label = QLabel(frame)
        label.setText(other["text"])
        combobox = QComboBox(frame)
        combobox.setStyleSheet("""
        QComboBox {
            background: #343a40;
            font: Monospace 15;
        }
        QComboBox QAbstractItemView {
          border: none;
          background: #212529;
          selection-background-color: blue;
        }
        """)
        combobox.addItems(other["values"])
        layout.addWidget(label)
        layout.addWidget(combobox)
        self.sub_frame.addWidget(frame)
        return frame, combobox
