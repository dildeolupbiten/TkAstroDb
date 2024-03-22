# -*- coding: utf-8 -*-

from .frame import Frame
from .utils import create_frames, create_widgets
from .libs import QWidget, QVBoxLayout, QHBoxLayout, QFrame


class MainCollapse(QWidget):
    def __init__(self, labels: list[str], frames):
        super().__init__()
        self.box = QVBoxLayout(self)
        self.widgets = []
        create_frames(self, frames)
        create_widgets(self, labels)


class Collapse(Frame):
    def __init__(self, labels: list[str], frames, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.frame = QFrame(parent=self)
        self.box.addWidget(self.frame)
        self.box = QHBoxLayout(self.frame)
        self.widgets = []
        create_frames(self, frames)
        create_widgets(self, labels)
        self.get = lambda: 1
