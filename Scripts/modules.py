# -*- coding: utf-8 -*-

import os
import json
import time
import shutil
import logging
import numpy as np
import pandas as pd
import tkinter as tk
import swisseph as swe
import xml.etree.ElementTree as ET

from subprocess import Popen
from threading import Thread
from scipy.stats import binom
from statistics import variance
from xlsxwriter import Workbook
from webbrowser import open_new
from urllib.error import URLError
from urllib.request import urlopen
from tkinter import ttk, PhotoImage
from datetime import datetime as dt
from matplotlib import pyplot as plt
from configparser import ConfigParser
from tkinter.ttk import Progressbar, Treeview
from tkinter.filedialog import askopenfilename
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk
)