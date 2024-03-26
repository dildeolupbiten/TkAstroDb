# -*- coding: utf-8 -*-

__version__ = "3.0.0"

from .about import About
from .consts import *
from .select import Select
from .button import Button
from .utils import from_xml
from .spinbox import Spinbox
from .carousel import Carousel
from .combobox import Combobox
from .collapse import MainCollapse, Collapse
from .calculations import Calculations
from .comparison import Comparison
from .table import Table
from .libs import (
    os, sys, path, getcwd, Qt, QApplication, QMainWindow, QLabel, QPushButton, QFrame, QHBoxLayout, QVBoxLayout,
    QMessageBox, QAction, QFileDialog, ElementTree, QObject, pyqtSignal, QThread, QTableWidget
)


class Worker(QObject):
    finished = pyqtSignal()

    def __init__(self, filename, parent=None):
        super().__init__(parent)
        self.filename = filename
        self.database = None
        self.categories = None

    def run(self):
        from_xml(self)


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.database, self.categories = [], []
        self.msgbox = QMessageBox(self)
        self.msgbox.info = lambda title, text: self.msgbox.information(self, title, text)
        self.version = __version__
        self.adb_version = None
        self.setWindowTitle("TkAstroDb")
        self.setGeometry(0, 0, 900, 900)
        self.menubar = self.menuBar()
        self.worker = None
        self.worker_thread = None
        self.on_load = False
        self.file_menu = self.menubar.addMenu("File")
        self.help_menu = self.menubar.addMenu("Help")
        self.add_action(
            menu=self.file_menu,
            label="Open Adb",
            func=self.load_database
        )
        self.add_action(
            menu=self.help_menu,
            label="About",
            func=lambda: About(version=__version__).exec_()
        )
        self.main_collapse = MainCollapse(
            labels=["Analysis", "Calculations", "Comparison"],
            frames={
                "button": {"layout": QHBoxLayout, "height": 100, "width": 900},
                "frame": {"layout": QVBoxLayout, "height": 800, "width": 900}
            }
        )
        self.analysis = Carousel(
            parent=self.main_collapse.widgets[0]["frame"],
            title="Analysis",
            labels=["Select Categories", "Rodden Ratings", "Sheets", "Optional Selections", "Analyze"]
        )
        self.calculations = Calculations(
            parent=self.main_collapse.widgets[1]["frame"],
            title="Calculations",
            msgbox=self.msgbox
        )
        self.comparison = Comparison(
            parent=self.main_collapse.widgets[2]["frame"],
            title="Comparison"
        )
        self.select_categories = Select(
            parent=self.analysis.widgets[0]["frame"],
            items=self.categories,
            title="Select Categories",
            selected=" Categories"
        )
        self.rodden_ratings = Select(
            parent=self.analysis.widgets[1]["frame"],
            items=RODDEN_RATINGS,
            title="Rodden Ratings",
            defaults=["AA"]
        )
        self.sheets = Select(
            parent=self.analysis.widgets[2]["frame"],
            items=SHEETS[1:],
            title="Sheets",
            defaults=SHEETS[1:]
        )
        self.optional_selections = Collapse(
            parent=self.analysis.widgets[3]["frame"],
            labels=OPTIONAL_SELECTIONS,
            frames={
                "button": {"layout": QVBoxLayout, "height": 400, "width": 300},
                "frame": {"layout": QHBoxLayout, "height": 400, "width": 500}
            },
            title="Optional Selections"
        )
        self.zodiac = Combobox(
            parent=self.optional_selections.widgets[0]["frame"],
            title="Zodiac",
            values=ZODIAC,
            defaults=["Tropical", "Hindu-Lahiri"],
            other={"text": "Ayanamsha", "values": AYANAMSHA}
        )
        self.house_systems = Combobox(
            parent=self.optional_selections.widgets[1]["frame"],
            title="House Systems",
            values=HOUSE_SYSTEMS,
            defaults=["Placidus"]
        )
        self.orb_factors = Spinbox(
            parent=self.optional_selections.widgets[2]["frame"],
            title="Orb Factors",
            widgets=ORB_FACTORS,
            defaults=ORB_FACTORS,
            spinbox=float,
            step=.01
        )
        self.midpoint_orb_factors = Spinbox(
            parent=self.optional_selections.widgets[3]["frame"],
            title="Midpoint Orb Factors",
            widgets=MIDPONT_ORB_FACTORS,
            defaults=MIDPONT_ORB_FACTORS,
            spinbox=float,
            step=.01
        )
        self.ignored_categories = Select(
            parent=self.optional_selections.widgets[4]["frame"],
            items=self.categories,
            title="Ignored Categories"
        )
        self.ignored_records = Select(
            parent=self.optional_selections.widgets[5]["frame"],
            items=IGNORED_RECORDS,
            title="Ignored Records",
            defaults=["Event"]
        )
        self.year_range = Spinbox(
            parent=self.optional_selections.widgets[6]["frame"],
            title="Year Range",
            widgets=["From", "To"],
            spinbox=int,
            step=1
        )
        self.latitude_range = Spinbox(
            parent=self.optional_selections.widgets[6]["frame"],
            title="Latitude Range",
            widgets=["From", "To"],
            spinbox=float,
            step=.01
        )
        self.longitude_range = Spinbox(
            parent=self.optional_selections.widgets[8]["frame"],
            title="Longitude Range",
            widgets=["From", "To"],
            spinbox=float,
            step=.01
        )
        self.button = Button(text="Analyze", parent=self.analysis.widgets[0]["frame"], widgets=self)
        self.add_widgets()
        self.select_categories.selected_records = QLabel(self.select_categories)
        self.select_categories.selected_records.setText(f"Selected Records{' ' * 3}: 0")
        self.select_categories.selected_records.setFixedWidth(400)
        self.select_categories.selected_records.setAlignment(Qt.AlignCenter)
        self.select_categories.button = QPushButton(self.select_categories)
        self.select_categories.button.setFixedWidth(250)
        self.select_categories.button.setText("Select Records Manually")
        self.select_categories.table = Table(
            columns=[],
            data=[],
            widget=self.select_categories.selected_records
        )
        self.select_categories.button.clicked.connect(lambda e: self.open_table())
        self.select_categories.box.addWidget(self.select_categories.selected_records, alignment=Qt.AlignCenter)
        self.select_categories.box.addWidget(self.select_categories.button, alignment=Qt.AlignCenter)
        self.setFixedSize(920, 940)
        if os.name == "nt":
            self.setStyleSheet(
                """
                QSpinBox, QDoubleSpinBox {
                    background-color: #343a40;
                    color: white;
                    border: none;
                }
                QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {
                    image: url("img/arrow-up.png");
                }

                QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {
                    image: url("img/arrow-down.png");
                }
                """
            )
        else:
            self.setStyleSheet(
                """
                QSpinBox, QDoubleSpinBox {
                    background-color: #343a40;
                    color: white;
                }
                """
            )

    def open_table(self):
        self.select_categories.table.exec_()

    def add_widgets(self):
        for index, widget in enumerate(
            [
                self.zodiac,
                self.house_systems,
                self.orb_factors,
                self.midpoint_orb_factors,
                self.ignored_categories,
                self.ignored_records,
                self.year_range,
                self.latitude_range,
                self.longitude_range
            ]
        ):
            if widget in [self.ignored_categories, self.ignored_records]:
                widget.list.setFixedWidth(400)
                self.optional_selections.widgets[index]["layout"].addWidget(widget, alignment=Qt.AlignCenter)
            elif widget in [self.orb_factors, self.midpoint_orb_factors]:
                for sub in widget.widgets:
                    sub["spinbox"].setFixedWidth(100)
                self.optional_selections.widgets[index]["layout"].addWidget(widget)
            else:
                self.optional_selections.widgets[index]["layout"].addWidget(widget)
        for index, widget in enumerate(
            [
                self.select_categories,
                self.rodden_ratings,
                self.sheets,
                self.optional_selections,
                self.button
            ]
        ):
            if widget == self.button:
                self.analysis.widgets[index]["layout"].addWidget(widget, alignment=Qt.AlignCenter)
            else:
                self.analysis.widgets[index]["layout"].addWidget(widget)
        self.main_collapse.widgets[0]["layout"].addWidget(self.analysis)
        self.main_collapse.widgets[1]["layout"].addWidget(self.calculations)
        self.main_collapse.widgets[2]["layout"].addWidget(self.comparison)
        self.setCentralWidget(self.main_collapse)

    def add_action(self, menu, label, func):
        action = QAction(label, self)
        action.triggered.connect(func)
        menu.addAction(action)

    def load_database(self):
        if self.on_load:
            self.msgbox.info("Warninig", "A file is currently being loaded.")
            return
        filename, _ = QFileDialog.getOpenFileName(self, "Open Adb", ".", "XML Files (*.xml)")
        if filename:
            if self.worker_thread and self.worker_thread.isRunning():
                self.worker_thread.quit()
                self.worker_thread.wait()
            self.on_load = True
            self.adb_version = path.split(filename)[-1]
            self.worker_thread = QThread()
            self.worker = Worker(filename)
            self.worker.moveToThread(self.worker_thread)
            self.worker_thread.started.connect(self.worker.run)
            self.worker.finished.connect(self.show_message_box)
            self.worker_thread.start()

    def show_message_box(self):
        self.database = self.worker.database
        self.categories = self.worker.categories
        self.select_categories.items = self.categories
        self.select_categories.table.data = [i[1:-1] for i in self.database]
        columns = [
            "Name", "Gender", "Rodden Rating", "Date", "Time",
            "Julian Date", "Latitude", "Longitude", "Place", "Country"
        ]
        self.select_categories.table.columns = columns

        self.select_categories.table.add_data()
        self.select_categories.table.frame.left.combobox.addItems(columns)
        self.select_categories.table.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.ignored_categories.items = self.categories
        self.on_load = False
        self.msgbox.info("Info", f"Loaded '{self.worker.filename}' successfully.")
        self.worker_thread.quit()
        self.worker_thread.wait()


def main():
    app = QApplication(sys.argv)
    with open(path.join(getcwd(), "qss", "style.qss"), "r") as style:
        app.setStyleSheet(style.read())
    window = App()
    window.show()
    sys.exit(app.exec_())
