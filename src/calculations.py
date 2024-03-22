# -*- coding: utf-8 -*-

from .frame import Frame
from .consts import CALC_TYPES, SHEETS
from .utils import calculate
from .libs import (
    path, Qt, QFrame, QVBoxLayout, QHBoxLayout, QPushButton,
    QProgressBar, QLabel, QComboBox, QDoubleSpinBox, QFileDialog, QObject, QThread, pyqtSignal
)


class Combobox(QFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.box = QVBoxLayout(self)
        self.setLayout(self.box)
        self.widgets = {}
        self.create_widgets()
        self.set_widgets()

    def create_widgets(self):
        for text in ["Calculation Type", "Method", "Alpha"]:
            frame = QFrame(self)
            layout = QHBoxLayout(frame)
            frame.setLayout(layout)
            label = QLabel(frame)
            label.setText(text)
            layout.addWidget(label)
            if text != "Alpha":
                widget = QComboBox(frame)
                widget.setStyleSheet(
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
                layout.addWidget(widget)
            else:
                widget = QDoubleSpinBox(frame)
                minimum = .000001
                widget.setMinimum(minimum)
                widget.setMaximum(1)
                widget.setSingleStep(minimum)
                widget.setValue(minimum)
                widget.setDecimals(6)
                label.setText(f"Alpha ({round((1 - widget.value()) * 100, 5)} %)")
                widget.valueChanged.connect(
                    lambda e, _label=label, w=widget: _label.setText(f"Alpha ({round((1 - w.value()) * 100, 5)} %)")
                )
                widget.setStyleSheet("background-color: #343a40; color: white;")
                layout.addWidget(widget)
            self.box.addWidget(frame)
            self.widgets[text] = {"frame": frame, "widget": widget}

    def set_widgets(self):
        self.widgets["Calculation Type"]["widget"].addItems(CALC_TYPES)
        self.widgets["Method"]["widget"].addItems(["Independent", "Subcategory"])
        self.widgets["Calculation Type"]["widget"].currentTextChanged.connect(self.show_hide_alpha_method)
        for i in ["Alpha"]:
            self.widgets[i]["frame"].setVisible(False)

    def show_hide_alpha_method(self):
        current_text = self.widgets["Calculation Type"]["widget"].currentText()
        if current_text == CALC_TYPES[0]:
            self.widgets["Alpha"]["frame"].setVisible(False)
            self.widgets["Method"]["frame"].setVisible(True)
        elif current_text == CALC_TYPES[-1]:
            self.widgets["Alpha"]["frame"].setVisible(True)
            self.widgets["Method"]["frame"].setVisible(False)
        else:
            self.widgets["Alpha"]["frame"].setVisible(False)
            self.widgets["Method"]["frame"].setVisible(False)

    def get(self):
        return {
            "calc_type": self.widgets["Calculation Type"]["widget"].currentText(),
            "method": self.widgets["Method"]["widget"].currentText(),
            "alpha": self.widgets["Alpha"]["widget"].value()
        }


class Files(QFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.box = QHBoxLayout(self)
        self.files = {}
        self.setLayout(self.box)
        self.create_widgets()

    def create_widgets(self):
        for i in range(1, 3):
            frame = QFrame(self)
            frame.setStyleSheet("border: 1 solid #343a40; border-radius: 5px;")
            layout = QVBoxLayout(frame)
            frame.setLayout(layout)
            label = QLabel(frame)
            label.setAlignment(Qt.AlignCenter)
            label.setStyleSheet("border: none;")
            label.setText(f"File {i}")
            button = QPushButton("Open", frame)
            button.clicked.connect(lambda e, widget=label, index=i: self.open(widget, index))
            layout.addWidget(label)
            layout.addWidget(button)
            self.box.addWidget(frame)

    def open(self, label, index):
        filename, _ = QFileDialog.getOpenFileName(self, "Open File", ".", "XLSX Files (*.xlsx)")
        if filename:
            self.files[f"file{index}"] = filename
            filepath, filename = path.split(filename)
            label.setText(f"File {index}: {filename}")


class Worker(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal(int)

    def __init__(self, file1, file2, calc_type, method, alpha, parent=None):
        super().__init__(parent)
        self.file1 = file1
        self.file2 = file2
        self.calc_type = calc_type
        self.method = method
        self.alpha = alpha

    def run(self):
        calculate(
            file1=self.file1,
            file2=self.file2,
            calc_type=self.calc_type,
            method=self.method,
            alpha=self.alpha,
            self=self
        )


class Calculations(Frame):
    def __init__(self, msgbox, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.worker = None
        self.worker_thread = None
        self.combobox = Combobox(self)
        self.msgbox = msgbox
        self.files = Files(self)
        self.button = QPushButton("Calculate", self)
        self.button.setStyleSheet("text-align: center;")
        self.button.setFixedWidth(200)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(200)
        self.box.addWidget(self.combobox)
        self.box.addWidget(self.files)
        self.box.addWidget(self.button, alignment=Qt.AlignCenter)
        self.box.addWidget(self.progress_bar, alignment=Qt.AlignCenter)
        self.setFixedHeight(500)
        self.button.clicked.connect(self.run)

    def run(self):
        if "file1" not in self.files.files:
            self.msgbox.info("Warning", "Select File 1")
            return
        if "file2" not in self.files.files:
            self.msgbox.info("Warning", "Select File 2")
            return
        selections = self.combobox.get()

        self.button.setEnabled(False)
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.quit()
            self.worker_thread.wait()
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(SHEETS) + 2)
        self.worker_thread = QThread()
        self.worker = Worker(
            self.files.files["file1"],
            self.files.files["file2"],
            selections["calc_type"],
            selections["method"],
            selections["alpha"]
        )
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.progress)
        self.worker.finished.connect(self.finished)
        self.worker_thread.start()

    def finished(self, v):
        self.progress_bar.setValue(v)
        self.button.setEnabled(True)
        self.msgbox.info("Info", f"Calculation completed successfully.")
        self.progress_bar.setVisible(False)
        self.worker_thread.quit()
        self.worker_thread.wait()

    def progress(self, v):
        self.progress_bar.setValue(v)
