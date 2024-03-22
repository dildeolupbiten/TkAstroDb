# -*- coding: utf-8 -*-

from .libs import QFrame, QLabel, QVBoxLayout, Qt, QScrollArea


class Frame(QFrame):
    def __init__(self, title, parent):
        super().__init__(parent=parent)
        self.box = QVBoxLayout(self)
        self.title = QLabel(self)
        self.title.setStyleSheet("font: 20px Monospace")
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setText(title)
        self.box.addWidget(self.title)


class ScrollableFrame(Frame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setFixedWidth(400)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: #343a40; color: white;")
        self.frame = QFrame(self.scroll_area)
        self.frame.setStyleSheet("background-color: #212529;")
        self.sub_frame = QVBoxLayout(self.frame)
        self.frame.setLayout(self.sub_frame)
        self.scroll_area.setWidget(self.frame)
        self.box.addWidget(self.scroll_area)
