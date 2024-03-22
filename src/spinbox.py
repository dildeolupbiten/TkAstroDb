# -*- coding: utf-8 -*-

from .frame import Frame, ScrollableFrame
from .libs import Qt, QSpinBox, QFrame, QHBoxLayout, QDoubleSpinBox, QLabel


class Spinbox(ScrollableFrame):
    def __init__(
            self, spinbox: int | float,
            widgets, defaults: dict = None, step: int | float = 1, *args, **kwargs
    ):
        super().__init__(*args, **kwargs)
        self.widgets = []
        self.spinbox = QSpinBox if spinbox == int else QDoubleSpinBox
        self.create_widgets(widgets, step)
        if defaults:
            self.set(defaults)

    def set(self, defaults):
        for i in self.widgets:
            i["spinbox"].setValue(defaults[i["label"].text()])

    def get(self):
        return {i["label"].text(): i["spinbox"].value() for i in self.widgets}

    def create_widgets(self, widgets, step):
        for widget in widgets:
            self.widgets += [{"frame": QFrame(self.frame)}]
            self.widgets[-1]["layout"] = QHBoxLayout(self.widgets[-1]["frame"])
            self.widgets[-1]["label"] = QLabel(self.widgets[-1]["frame"])
            self.widgets[-1]["label"].setText(widget)
            self.widgets[-1]["label"].setFixedWidth(150)
            self.widgets[-1]["spinbox"] = self.spinbox(self.widgets[-1]["frame"])
            self.widgets[-1]["spinbox"].setSingleStep(step)
            self.widgets[-1]["spinbox"].setStyleSheet("background-color: #343a40; color: white;")
            self.widgets[-1]["layout"].addWidget(self.widgets[-1]["label"])
            self.widgets[-1]["layout"].addWidget(self.widgets[-1]["spinbox"])
            self.sub_frame.addWidget(self.widgets[-1]["frame"])
