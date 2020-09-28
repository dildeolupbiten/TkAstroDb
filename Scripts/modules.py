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

from PIL import ImageTk
from tkinter import ttk
from threading import Thread
from subprocess import Popen
from scipy.stats import binom
from xlsxwriter import Workbook
from webbrowser import open_new
from statistics import variance
from urllib.error import URLError
from urllib.request import urlopen
from datetime import datetime as dt
from configparser import ConfigParser
from tkinter.ttk import Progressbar, Treeview
