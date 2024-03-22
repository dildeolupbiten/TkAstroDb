# -*- coding: utf-8 -*-

from .frame import Frame
from .utils import create_frames
from .libs import QVBoxLayout, QHBoxLayout, QPushButton, Qt, QFrame, QLabel, QMessageBox


class Carousel(Frame):
    def __init__(self, labels, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.labels = labels
        self.widgets = []
        self.index = 0
        self.label = None
        create_frames(
            self,
            frames={
                "frame": {"layout": QVBoxLayout, "height": 500},
                "button": {"layout": QHBoxLayout, "height": 100}
            }
        )
        self.create_widgets()
        self.toggle(0)

    def create_widgets(self):
        for label in self.labels:
            self.widgets += [{"frame": QFrame(self.box.frame["frame"])}]
            self.widgets[-1]["layout"] = QVBoxLayout(self.widgets[-1]["frame"])
            self.widgets[-1]["frame"].setVisible(False)
            self.widgets[-1]["frame"].setObjectName(label)
            self.box.frame["frame"].layout.addWidget(self.widgets[-1]["frame"])
        for text, var in zip(["<<", ">>"], [-1, 1]):
            button = QPushButton(text, self.box.frame["button"])
            button.clicked.connect(lambda e, v=var: self.toggle(v))
            self.box.frame["button"].layout.addWidget(button)
            if text == "<<":
                self.label = QLabel(self.box.frame["button"])
                self.label.setAlignment(Qt.AlignCenter)
                self.box.frame["button"].layout.addWidget(self.label)

    def toggle(self, var):
        if var and self.index + 1 != len(self.widgets) and not self.widgets[self.index]["frame"].children()[-1].get():
            passed = False
            if hasattr(self.widgets[self.index]["frame"].children()[-1], "table"):
                if getattr(self.widgets[self.index]["frame"].children()[-1], "table").selected:
                    passed = True
            if not passed:
                QMessageBox.information(self, "Warning", "You should select one option.")
                return
        self.index += var
        if self.index < 0:
            self.index = 0
            return
        elif self.index == len(self.widgets):
            self.index -= 1
            return

        self.label.setText(f"{self.index + 1}/{len(self.widgets)}")
        self.widgets[self.index - var]["frame"].setVisible(False)
        self.widgets[self.index]["frame"].setVisible(True)
