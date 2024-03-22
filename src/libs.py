# -*- coding: utf-8 -*-

import os
import sys
import numpy as np
import pandas as pd
import swisseph as swe

from queue import Queue
from numpy import e, pi
from os import path, getcwd
from threading import Thread
from webbrowser import open_new
from scipy.integrate import quad
from xml.etree import ElementTree
from time import time, perf_counter
from scipy.stats import binom, norm
from datetime import datetime as dt
from warnings import filterwarnings
from matplotlib.figure import Figure
from xlsxwriter.workbook import Workbook
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar

from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt, QObject, pyqtSignal, QThread
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFrame, QListWidget, QLabel,
    QSpinBox, QScrollArea, QComboBox, QDoubleSpinBox, QProgressBar, QAction, QFileDialog,
    QMessageBox, QDialog, QCheckBox, QHeaderView, QTableWidget, QTableWidgetItem,
    QAbstractScrollArea, QLineEdit
)
