# -*- coding: utf-8 -*-

from .utils import get_selections, filter_database, get_chart_patterns
from .libs import Qt, QVBoxLayout, QPushButton, QProgressBar, QFrame, QLabel, QObject, pyqtSignal, QThread


class Worker(QObject):
    progress = pyqtSignal(int)
    finished = pyqtSignal(int)

    def __init__(self, filters, filtered, widgets, parent=None):
        super().__init__(parent)
        self.filters = filters
        self.filtered = filtered
        self.widgets = widgets

    def run(self):
        get_chart_patterns(
            filters=self.filters,
            records=self.filtered,
            version=self.widgets.version,
            adb_version=self.widgets.adb_version,
            self=self
        )


class Button(QFrame):
    def __init__(self, text, widgets, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.widgets = widgets
        self.worker = None
        self.worker_thread = None
        self.box = QVBoxLayout(self)
        self.setLayout(self.box)
        self.button = QPushButton(text, self)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.box.addWidget(self.button)
        self.box.addWidget(self.progress_bar)
        self.button.clicked.connect(self.run)
        self.setFixedWidth(200)

    def run(self):
        self.button.setEnabled(False)
        filters = get_selections(self.widgets)
        filtered = filter_database(self.widgets.database, filters)
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.quit()
            self.worker_thread.wait()
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(filtered))
        self.worker_thread = QThread()
        self.worker = Worker(filters, filtered, self.widgets)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.progress)
        self.worker.finished.connect(self.finished)
        self.worker_thread.start()

    def finished(self, v):
        self.progress_bar.setValue(v)
        self.button.setEnabled(True)
        self.widgets.msgbox.info("Info", f"Analysis completed successfully.")
        self.progress_bar.setVisible(False)
        self.worker_thread.quit()
        self.worker_thread.wait()

    def progress(self, v):
        self.progress_bar.setValue(v)
