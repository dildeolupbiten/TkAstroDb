#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__version__ = "1.5.5"

import os
import sys
import ssl
import time
import shutil
import logging
import platform
import threading
import subprocess
import webbrowser
import tkinter as tk
import sqlite3 as sql
import urllib.request as urllib
import xml.etree.ElementTree as et
import tkinter.messagebox as msgbox

from math import cos, sin, radians
from tkinter.ttk import Progressbar, Treeview
from datetime import datetime as dt
    
whl_path = os.path.join(os.getcwd(), "Eph", "Whl")

    
def select_module(module_name, module_files):
    if os.name == "posix":
        os.system(f"pip3 install {module_name}")
    elif os.name == "nt":
        if sys.version_info.minor == 6:
            if platform.architecture()[0] == "32bit":
                new_path = os.path.join(whl_path, module_files[0])
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "64bit":
                new_path = os.path.join(whl_path, module_files[1])
                os.system(f"pip3 install {new_path}")
        elif sys.version_info.minor == 7:
            if platform.architecture()[0] == "32bit":
                new_path = os.path.join(whl_path, module_files[2])
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "64bit":
                new_path = os.path.join(whl_path, module_files[3])
                os.system(f"pip3 install {new_path}")
                
                
try:
    import numpy as np
except ModuleNotFoundError:
    os.system("pip3 install numpy")
    import numpy as np
try:
    from dateutil import tz
except ModuleNotFoundError:
    os.system("pip3 install python-dateutil")
    from dateutil import tz
try:
    from geopy.geocoders import Nominatim
except ModuleNotFoundError:
    os.system("pip3 install geopy")
    from geopy.geocoders import Nominatim
try:
    import shapely
except ModuleNotFoundError:            
    select_module(
        "shapely", 
        [i for i in os.listdir(whl_path) if "Shapely" in i]
    )  
try:
    from tzwhere import tzwhere
except ModuleNotFoundError:
    os.system("pip3 install tzwhere")
    from tzwhere import tzwhere
try:
    import swisseph as swe
except ModuleNotFoundError:
    select_module(
        "pyswisseph", 
        [i for i in os.listdir(whl_path) if "pyswisseph" in i]
    )
    import swisseph as swe    
try:
    from countryinfo import CountryInfo
except ModuleNotFoundError:
    os.system("pip3 install countryinfo")
    if os.name == "nt":
        import site
        sitepackages = site.getsitepackages()[-1]
        if "countryinfo" in os.listdir(sitepackages):
            path = os.path.join(sitepackages, "countryinfo")
            package_file = "countryinfo.py"
            package_path = os.path.join(path, package_file)
            script_code = []
            with open(file=package_path, mode="r+", encoding="utf-8") \
                    as read_file:
                readlines = read_file.readlines()
                for line in readlines:
                    script_code.append(line)
            if "country_info = json.load(open(file_path))" in script_code[29]:
                script_code[29] = script_code[29].replace(
                    "file_path", 
                    "file_path, encoding='utf-8'")
            with open(file=package_path, mode="w+", encoding="utf-8") \
                    as write_file:
                for line in script_code:
                    write_file.write(line)
    from countryinfo import CountryInfo   
try:
    import certifi
except ModuleNotFoundError:
    os.system("pip3 install certifi")
    import certifi
try:
    import xlrd
except ModuleNotFoundError:
    os.system("pip3 install xlrd")
    import xlrd
try:
    import xlwt
except ModuleNotFoundError:
    os.system("pip3 install xlwt")
    import xlwt
    
icon = subprocess.Popen(
    "where python", 
    stdout=subprocess.PIPE
).communicate()[0].decode("utf-8")[:-2] if os.name == "nt" else None
    

# -----------------------------sqlite3 & xml------------------------------------


connect = sql.connect("TkAstroDb.db")
cursor = connect.cursor()

col_names = "no, add_date, adb_id, name, gender, rr, date, time, " \
            "julian_date, lat, c_lat, lon, c_lon, place, country, " \
            "adb_link, category"

cursor.execute(f"CREATE TABLE IF NOT EXISTS DATA({col_names})")

database, _database, category_names, xml_files = [], [], [], []

_count_ = 0

all_categories, category_dict, _category_dict = {}, {}, {}
        
name_order = {
    j: i for i, j in enumerate(
        ["adb_export_update_190506_1154.xml",
         "adb_export_update_190309_1118.xml",
         "adb_export_update_190113_1936.xml",
         "adb_export_181128_2309.xml"])        
}

for xml in os.listdir(os.getcwd()):
    if xml.endswith("xml"):
        xml_files.append(xml)
        
logging.basicConfig(
    format="- %(levelname)s - %(asctime)s - %(message)s", 
    level=logging.INFO,
    datefmt="%d.%m.%Y %H:%M:%S"
    )    


def parse_xml():
    global database, category_dict, _database, _category_dict
    database = []
    category_dict = {}
    a = open(file="65657 records.txt", mode="w", encoding="utf-8")
    for xml_file in sorted(xml_files, key=name_order.get):
        xml_database = []
        ignored = 0
        logging.info(f"Parsing {xml_file} file...")         
        tree = et.parse(xml_file)
        root = tree.getroot()
        for _i in range(1000000):
            try:
                start_stop = False
                user_data = []
                for gender, roddenrating, bdata, adb_link, categories in \
                    zip(
                        root[_i + 2][1].findall("gender"),
                        root[_i + 2][1].findall("roddenrating"),
                        root[_i + 2][1].findall("bdata"),
                        root[_i + 2][2].findall("adb_link"),
                        root[_i + 2][3].findall("categories")
                    ):
                    _name = root[_i + 2][1][0].text
                    for _record_ in database:
                        if _name == _record_[1]:
                            ignored += 1
                            start_stop = True
                            break
                    if start_stop:
                        break
                    else:
                        sbdate_dmy = bdata[1].text
                        sbtime = bdata[2].text
                        jd_ut = bdata[2].get("jd_ut")
                        lat = bdata[3].get("slati")
                        lon = bdata[3].get("slong")
                        place = bdata[3].text
                        country = bdata[4].text
                        category = [
                            (
                                categories[_j].get("cat_id"),
                                categories[_j].text
                            )
                            for _j in range(len(categories))
                        ]
                        for cate in category:
                            if cate[0] not in category_dict.keys():
                                category_dict[cate[0]] = cate[1]
                        user_data.append(int(root[_i + 2].get("adb_id")))
                        user_data.append(_name)
                        user_data.append(gender.text)
                        user_data.append(roddenrating.text)
                        user_data.append(sbdate_dmy)
                        user_data.append(sbtime)
                        user_data.append(jd_ut)
                        user_data.append(lat)
                        user_data.append(lon)
                        user_data.append(place)
                        user_data.append(country)
                        user_data.append(adb_link.text)
                        user_data.append(category)
                        if len(user_data) != 0:                            
                            xml_database.append(user_data)
                            a.write(f"{user_data}\n\n")
                            a.flush()                 
            except IndexError:
                break
        logging.info("Parsing completed.")
        logging.info(f"{ignored} records are ignored.")
        logging.info(f"{len(xml_database)} records are inserted.")
        database.extend(xml_database)
        _database = [i for i in database]
        _category_dict = {i: j for i, j in category_dict.items()}
    logging.info(f"{len(database)} records are available.")
    a.close()
                
            
def merge_databases(db, cat_dict):
    global _count_, database, category_dict
    database = [i for i in _database]
    category_dict = {i: j for i, j in _category_dict.items()}
    reverse_category_list = {
        value: key for key, value in cat_dict.items()
    }
    for _i_ in cursor.execute("SELECT * FROM DATA"):
        _data_ = list(_i_)[2:]
        _data_.pop(8)
        _data_.pop(9)
        new_category = []
        edit_data = _data_[:12]
        if "|" in _data_[12]:
            edit_category = _data_[12].split("|")
            for _cat_ in edit_category:
                if _cat_ in cat_dict.values():
                    new_category.append(
                        (reverse_category_list[_cat_], _cat_)
                    )
                else:
                    new_category.append((str(4014 + _count_), _cat_))
                    cat_dict[str(4014 + _count_)] = _cat_
                    _count_ += 1
                    reverse_category_list = {
                        value: key for key, value in cat_dict.items()
                    }
        else:
            if _data_[12] in cat_dict.values():
                new_category.append(
                    (reverse_category_list[_data_[12]], _data_[12])
                )
            else:
                new_category.append((str(4014 + _count_), _data_[12]))
                cat_dict[str(4014 + _count_)] = _data_[12]
                reverse_category_list = {
                    value: key for key, value in cat_dict.items()
                }
                _count_ += 1
        edit_data.append(new_category)
        db.append(edit_data)
        
              
def group_categories(db, cat_dict):
    global category_names, all_categories
    category_groups = {}
    count = 0
    for _record in db:
        for _category in _record[-1]:
            if (_category[0], _category[1]) not in category_groups:
                if _category[1] is None:
                    pass
                category_groups[(_category[0], _category[1])] = []
            category_groups[(_category[0], _category[1])].append(_record)
    all_categories = category_groups
    category_names = sorted(
        [i for i in cat_dict.values() if i is not None]
    )


def load_database(*args): 
    merge_databases(*args)
    group_categories(*args)
    

parse_xml()    
load_database(database, category_dict)


# ---------------------------------tkinter--------------------------------------


canvas = None
toplevel = None


def create_toplevel():
    global toplevel
    if toplevel is not None:
        toplevel.destroy()
    toplevel = tk.Toplevel()
    toplevel.resizable(width=False, height=False)
    toplevel.title("Open Chart")
    if os.name == "nt":
        toplevel.iconbitmap(icon)
    return toplevel


def create_canvas(master_):
    global canvas
    if canvas is not None:
        canvas.destroy()
    canvas = tk.Canvas(master=master_, bg="white", width=1270, height=660)
    canvas.grid(row=0, column=1)
    return canvas
    
    
def oval_object(x, y, r, dash=True):
    if dash is True:
        dash = (1, 10)
        canvas.create_oval(
            x - r,
            y - r,
            x + r,
            y + r,
            fill="white",
            width=2,
            dash=dash
        )
    else:
        canvas.create_oval(
            x - r,
            y - r,
            x + r,
            y + r,
            fill="white",
            width=2,
        )


def line_object(x1, y1, x2, y2, width=2, fill="black"):
    canvas.create_line(x1, y1, x2, y2, width=width, fill=fill)


def text_object(x, y, _text, width=0, font="Arial", fill="black"):
    canvas.create_text(x, y, text=_text, width=width, font=font, fill=fill)
    
    
# --------------------------------swisseph--------------------------------------


swe.set_ephe_path(os.path.join(os.getcwd(), "Eph"))

PLANETS = {
    "Sun": swe.SUN,
    "Moon": swe.MOON,
    "Mercury": swe.MERCURY,
    "Venus": swe.VENUS,
    "Mars": swe.MARS,
    "Jupiter": swe.JUPITER,
    "Saturn": swe.SATURN,
    "Uranus": swe.URANUS,
    "Neptune": swe.NEPTUNE,
    "Pluto": swe.PLUTO,
    "North Node": swe.TRUE_NODE,
    "Chiron": swe.CHIRON
}

ASPECT_SYMBOLS = {
    "Conjunction": "\u260C",
    "Semi-Sextile": "\u26BA",
    "Semi-Square": "\u2220",
    "Sextile": "\u26B9",
    "Quintile": "Q",
    "Square": "\u25A1",
    "Trine": "\u25B3",
    "Sesquiquadrate": "\u26BC",
    "BiQuintile": "bQ",
    "Quincunx": "\u26BB",
    "Opposite": "\u260D",
}


signs = [
    "Aries", "Taurus", "Gemini", "Cancer", "Leo", "Virgo",
    "Libra", "Scorpio", "Sagittarius", "Capricorn", "Aquarius", "Pisces"
]

planets = [i for i, j in PLANETS.items()]

modern_rulership = {
    "Aries": "Pluto",
    "Taurus": "Venus",
    "Gemini": "Mercury",
    "Cancer": "Moon",
    "Leo": "Sun",
    "Virgo": "Mercury",
    "Libra": "Venus",
    "Scorpio": "Pluto",
    "Sagittarius": "Neptune",
    "Capricorn": "Uranus",
    "Aquarius": "Uranus",
    "Pisces": "Neptune"
}

traditional_rulership = {
    "Aries": "Mars",
    "Taurus": "Venus",
    "Gemini": "Mercury",
    "Cancer": "Moon",
    "Leo": "Sun",
    "Virgo": "Mercury",
    "Libra": "Venus",
    "Scorpio": "Mars",
    "Sagittarius": "Jupiter",
    "Capricorn": "Saturn",
    "Aquarius": "Saturn",
    "Pisces": "Jupiter"
}

houses = [f"{i + 1}" for i in range(12)]

hsys = "P"

house_systems = {
    "P": "Placidus",
    "K": "Koch",
    "O": "Porphyrius",
    "R": "Regiomontanus",
    "C": "Campanus",
    "E": "Equal",
    "W": "Whole Signs"
}


def dd_to_dms(dd):
    degree = int(dd)
    minute = int((dd - degree) * 60)
    second = round(float((dd - degree - minute / 60) * 3600))
    return f"{degree}\u00b0 {minute}\' {second}\""


def dms_to_dd(dms):
    dms = dms.replace("\u00b0", " ").replace("\'", " ").replace("\"", " ")
    degree = int(dms.split(" ")[0])
    minute = float(dms.split(" ")[1]) / 60
    second = float(dms.split(" ")[2]) / 3600
    return degree + minute + second


open_chart = False


class Chart:
    def __init__(self, julian_date: float, longitude: float, latitude: float):
        self.julian_date = julian_date
        self.longitude = longitude
        self.latitude = latitude
        self.PLANETS = PLANETS
        self._PLANETS = {i: j for i, j in self.PLANETS.items()}
        self._SIGNS = {j: 30 * i for i, j in enumerate(signs)}
        self.SIGNS = [i for i in signs]
        self.SIGN_SYMBOLS = [
            "\u2648", "\u2649", "\u264A", "\u264B", "\u264C", "\u264D",
            "\u264E", "\u264F", "\u2650", "\u2651", "\u2652", "\u2653"]
        self.SIGN_COLORS = ["red", "green", "yellow", "blue"] * 3
        self._SIGN_SYMBOLS = {
            i: j for i, j in zip(self.SIGNS, self.SIGN_SYMBOLS)
        }
        self._SIGN_COLORS = {
            i: j for i, j in zip(self.SIGNS, self.SIGN_COLORS)
        }
        self.PLANET_SYMBOLS = {
            planets[0]: "\u2299",
            planets[1]: "\u263E",
            planets[2]: "\u263F",
            planets[3]: "\u2640",
            planets[4]: "\u2642",
            planets[5]: "\u2643",
            planets[6]: "\u2644",
            planets[7]: "\u2645",
            planets[8]: "\u2646",
            planets[9]: "\u2647",
            planets[10]: "\u260A",
            planets[11]: "\u26B7"
        }
        self.ASPECT_SYMBOLS = ASPECT_SYMBOLS
        self.ASPECT_SYMBOLS["Null"] = " "
        self.ASPECTS = {i: [] for i in self._PLANETS.keys()}
        self.PLANET_INFO_FORMAT = []
        self.MODIFY_PLANET_INFO_FORMAT = []
        self.PLANET_DEGREES = []
        self.HOUSE_INFO_FORMAT = []
        self.HOUSE_SIGN = []
        self.PLANET_SIGN_HOUSE = []
        self.NEW_HOUSE_DEGREES = []
        self.MIDPOINTS = {}
        self.ORB_FACTORS = {}
        self.CONJUNCTION = []
        self.SEMI_SEXTILE = []
        self.SEMI_SQUARE = []
        self.SEXTILE = []
        self.QUINTILE = []
        self.SQUARE = []
        self.TRINE = []
        self.SESQUIQUADRATE = []
        self.BIQUINTILE = []
        self.QUINCUNX = []
        self.OPPOSITE = []
        self.NULL = []
        self.__PLANET_DEGREES = []
        self.MIDPOINT_OF_HOUSES = []
        self.__PLANET_INFO_FORMAT = []
        self.__HOUSE_INFO_FORMAT = []
        self.organize_chart_data()
        self.convert_planet_degrees()
        self.convert_house_info()
        self.convert_house_degrees()
        self.find_planet_positions()
        if open_chart is True:
            self.draw_oval_object()
            self.draw_line_object()
            self.draw_houses()
            self.draw_signs()
            self.draw_house_numbers()
            self.draw_sign_symbols()
            self.draw_planets()
            self.draw_aspects()
            self.parse_aspects()
            self.new_order_of_the_aspects()
            self.draw_house_info()
            self.draw_planet_info()
            self.draw_aspect_info()

    @staticmethod
    def convert_angle(angle):
        for i in range(12):
            if i * 30 <= angle < (i + 1) * 30:
                return angle - (30 * i), signs[i]

    def planet_pos(self, planet):
        calc = self.convert_angle(swe.calc_ut(self.julian_date, planet)[0])
        return calc[0], calc[1]

    def planet_pos_(self, planet):
        calc = self.convert_angle(swe.calc_ut(self.julian_date, planet)[0])
        return dd_to_dms(calc[0]), calc[1]

    def house_cusps(self):
        global hsys
        house = []
        asc = 0
        angle = []
        for i, j in enumerate(swe.houses(
                self.julian_date, self.latitude, self.longitude,
                bytes(hsys.encode("utf-8")))[0]):
            if i == 0:
                asc += j
            angle.append(j)
            house.append((
                f"{i + 1}",
                f"{self.convert_angle(j)[0]}",
                f"{self.convert_angle(j)[1]}"))
        return house, asc, angle

    def house_pos(self):
        return self.house_cusps()[2]

    def organize_chart_data(self):
        count = 0
        planet_info_format = []
        house_info_format = []
        for key, value in self.PLANETS.items():
            planet = self.planet_pos(value)
            planet_info = (
                key,
                planet[0],
                planet[1]
            )
            planet_info_format.append(planet_info)
            house_info = [
                self.house_cusps()[0][count][0],
                self.house_cusps()[0][count][1],
                self.house_cusps()[0][count][-1]
            ]
            house_info_format.append(house_info)
            count += 1
        self.PLANET_INFO_FORMAT = planet_info_format
        self.HOUSE_INFO_FORMAT = house_info_format

    def convert_planet_degrees(self):
        for i in self.PLANET_INFO_FORMAT:
            planet_info_1 = [i[0], i[1] + self._SIGNS[i[2]], i[2]]
            planet_info_2 = [i[0], i[1] + self._SIGNS[i[2]]]
            self.MODIFY_PLANET_INFO_FORMAT.append(planet_info_1)
            self.PLANET_DEGREES.append(planet_info_2)

    def convert_house_info(self):
        for i in self.HOUSE_INFO_FORMAT:
            house_info = [i[0], i[2]]
            self.HOUSE_SIGN.append(house_info)

    def convert_house_degrees(self):
        for i in self.HOUSE_INFO_FORMAT:
            house_info = [i[0], float(i[1]) + self._SIGNS[i[2]], i[2]]
            self.NEW_HOUSE_DEGREES.append(house_info)

    def find_planet_positions(self):
        for i, j in enumerate(self.MODIFY_PLANET_INFO_FORMAT):
            for k, m in enumerate(self.NEW_HOUSE_DEGREES):
                try:
                    if self.NEW_HOUSE_DEGREES[k][1] - \
                            self.NEW_HOUSE_DEGREES[k + 1][1] > 180 \
                            and j[1] < self.NEW_HOUSE_DEGREES[k + 1][1]:
                        next_house_degree = self.NEW_HOUSE_DEGREES[k + 1][1] \
                            + 360
                        new_planet_degree = j[1] + 360
                        if self.NEW_HOUSE_DEGREES[k][1] < \
                                new_planet_degree < next_house_degree:
                            self.PLANET_SIGN_HOUSE.append(
                                [j[0], j[2], self.HOUSE_SIGN[k][0]]
                            )
                    elif self.NEW_HOUSE_DEGREES[k][1] - \
                            self.NEW_HOUSE_DEGREES[k + 1][1] > 180 \
                            and j[1] > self.NEW_HOUSE_DEGREES[k][1]:
                        next_house_degree = self.NEW_HOUSE_DEGREES[k + 1][1] \
                            + 360
                        if self.NEW_HOUSE_DEGREES[k][1] < j[1] < \
                                next_house_degree:
                            self.PLANET_SIGN_HOUSE.append(
                                [j[0], j[2], self.HOUSE_SIGN[k][0]]
                            )
                    if self.NEW_HOUSE_DEGREES[k][1] < j[1] < \
                            self.NEW_HOUSE_DEGREES[k + 1][1]:
                        self.PLANET_SIGN_HOUSE.append(
                            [j[0], j[2], self.HOUSE_SIGN[k][0]]
                        )
                except IndexError:
                    if self.NEW_HOUSE_DEGREES[k][1] - \
                            self.NEW_HOUSE_DEGREES[0][1] > 180 \
                            and j[1] < self.NEW_HOUSE_DEGREES[0][1]:
                        next_house_degree = self.NEW_HOUSE_DEGREES[0][1] + 360
                        new_planet_degree = j[1] + 360
                        if self.NEW_HOUSE_DEGREES[k][1] < \
                                new_planet_degree < next_house_degree:
                            self.PLANET_SIGN_HOUSE.append(
                                [j[0], j[2], self.HOUSE_SIGN[k][0]]
                            )
                    elif self.NEW_HOUSE_DEGREES[k][1] - \
                            self.NEW_HOUSE_DEGREES[0][1] > 180 \
                            and j[1] > self.NEW_HOUSE_DEGREES[k][1]:
                        next_house_degree = self.NEW_HOUSE_DEGREES[0][1] + 360
                        if self.NEW_HOUSE_DEGREES[k][1] < j[1] < \
                                next_house_degree:
                            self.PLANET_SIGN_HOUSE.append(
                                [j[0], j[2], self.HOUSE_SIGN[k][0]]
                            )
                    if self.NEW_HOUSE_DEGREES[k][1] < j[1] < \
                            self.NEW_HOUSE_DEGREES[0][1]:
                        self.PLANET_SIGN_HOUSE.append(
                            [j[0], j[2], self.HOUSE_SIGN[k][0]]
                        )

    def select_rulership(self, ruler_list, rulership, ruler_type, 
                         ruler_sign, i, m):
        ruler_list.append([f"Lord {i + 1}", f"House {m[2]}"])
        for n, o in enumerate(signs):
            if rulership[o] == ruler_type:
                if [f"{i + 1}", o] in self.HOUSE_SIGN:
                    if n % 2 == 0:
                        ruler_sign[o].append(
                            [f"Lord {i + 1} is {ruler_type} (+)",
                             f"House {m[2]}"])
                    else:
                        ruler_sign[o].append(
                            [f"Lord {i + 1} is {ruler_type} (-)", 
                             f"House {m[2]}"])

    @staticmethod
    def set_lords(ruler_sign):
        lords = {}
        for keys, values in ruler_sign.items():
            for value in values:
                for i in range(12):
                    if f"Lord {i + 1} is" in value[0]:
                        lords[f"Lord {i + 1}"] = value
        return lords

    def get_chart_data(self):
        traditional_ruler, modern_ruler = [], []
        trad_ruler_sign = {i: [] for i in signs}
        mode_ruler_sign = {i: [] for i in signs}
        for i, j in enumerate(self.house_cusps()[0]):
            traditional = traditional_rulership[self.house_cusps()[0][i][2]]
            modern = modern_rulership[self.house_cusps()[0][i][2]]
            for k, m in enumerate(self.PLANET_SIGN_HOUSE):
                if traditional == m[0]:
                    self.select_rulership(ruler_list=traditional_ruler,
                                          rulership=traditional_rulership,
                                          ruler_type=traditional,
                                          ruler_sign=trad_ruler_sign, 
                                          i=i, m=m)
                if modern == m[0]:
                    self.select_rulership(ruler_list=modern_ruler,
                                          rulership=modern_rulership,
                                          ruler_type=modern,
                                          ruler_sign=mode_ruler_sign, 
                                          i=i, m=m)
        trad_lords = self.set_lords(trad_ruler_sign)
        mode_lords = self.set_lords(mode_ruler_sign)
        return self.PLANET_SIGN_HOUSE, self.HOUSE_SIGN, self.PLANET_DEGREES, \
            traditional_ruler, modern_ruler, trad_ruler_sign, \
            mode_ruler_sign, trad_lords, mode_lords

    def append_house(self, house, i, j, name=""):
        if name == "":
            house.append((
                f"House {i + 1}",
                f"{dd_to_dms(self.convert_angle(j)[0])}",
                f"{self.convert_angle(j)[1]}"))
        else:
            house.append((
                f"{name}",
                f"{dd_to_dms(self.convert_angle(j)[0])}",
                f"{self.convert_angle(j)[1]}"))

    def house_cusps_(self):
        houses_ = []
        asc = 0
        angle = []
        for i, j in enumerate(swe.houses(
                self.julian_date, self.latitude, self.longitude,
                bytes(hsys.encode("utf-8")))[0]):
            if i == 0:
                asc += j
            angle.append(j)
            if i + 1 == 1:
                self.append_house(houses_, i, j, name="Asc")
            elif i + 1 == 4:
                self.append_house(houses_, i, j, name="IC")
            elif i + 1 == 7:
                self.append_house(houses_, i, j, name="Dsc")
            elif i + 1 == 10:
                self.append_house(houses_, i, j, name="MC")
            else:
                self.append_house(houses_, i, j)
        return houses_, asc, angle

    def sign_pos(self):
        asc = self.house_cusps_()[1]
        degree = self.house_cusps_()[0][0][1].replace("'", "")\
            .replace('"', "").replace("\u00b0", "")
        end = 30 - dms_to_dd(degree) + asc
        start = end - 30
        signs_ = []
        for i, j in enumerate(self.house_cusps_()[0]):
            signs_.append(j[2])
        for i in self.SIGNS:
            if i not in signs_:
                if self.SIGNS.index(i) != 11:
                    index_1 = self.SIGNS.index(i) + 1
                    sign = self.SIGNS[index_1]
                    index_2 = signs_.index(sign)
                    signs_.insert(index_2, i)
                else:
                    index_1 = self.SIGNS.index(i) - 1
                    sign = self.SIGNS[index_1]
                    index_2 = signs_.index(sign) + 1
                    signs_.insert(index_2, i)
                sign = self.SIGNS[index_1]
                index_2 = signs_.index(sign)
                signs_.insert(index_2, i)
        for i in signs_:
            count = signs_.count(i)
            if count > 1:
                signs_.pop(signs_.index(i, 1))
        _signs_ = []
        for i, j in enumerate(signs_):
            _start = start + (i * 30)
            _end = end + (i * 30)
            if _start > 360:
                _start -= 360
            if _end > 360:
                _end -= 360
            _signs_.append(_start)
        return _signs_, signs_

    @staticmethod
    def line_components(angle, r):
        x, y = 610, 350
        x_ = x + (r * cos(radians(angle)))
        y_ = y - (r * sin(radians(angle)))
        return x_, y_

    def x_y(self, angle, r1_, r2_):
        x1, y1 = self.line_components(angle=angle, r=r1_)
        x2, y2 = self.line_components(angle=angle, r=r2_)
        return x1, y1, x2, y2

    @staticmethod
    def draw_oval_object(x=610, y=350):
        oval_object(x=x, y=y, r=260, dash=False)
        oval_object(x=x, y=y, r=210, dash=False)
        oval_object(x=x, y=y, r=165)
        oval_object(x=x, y=y, r=60, dash=False)

    @staticmethod
    def draw_line_object(x=900, y=1200):
        line_object(x1=x, y1=205, x2=y, y2=205, width=1)
        line_object(x1=x, y1=410, x2=y, y2=410, width=1)
        line_object(x1=x, y1=428, x2=x, y2=624, width=1)
        line_object(x1=x, y1=428 + (13 * 15), x2=x + 25 + (20 * 15),
                    y2=428 + (13 * 15), width=1)

    def draw_houses(self):
        for i, j in enumerate(self.house_pos()):
            _degree = j - (self.house_pos()[0] - 180)
            if _degree < 0:
                _degree += 360
            elif _degree > 360:
                _degree -= 360
            self.MIDPOINT_OF_HOUSES.append(_degree)
            x1, y1, x2, y2 = self.x_y(angle=_degree, r1_=60, r2_=210)
            if i == 0 or i == 3 or i == 6 or i == 9:
                line_object(x1, y1, x2, y2, width=4)
            else:
                line_object(x1, y1, x2, y2, width=2)

    def draw_signs(self):
        sign_pos = self.sign_pos()[0]
        for i in sign_pos:
            _degree = i - (self.house_pos()[0] - sign_pos[1])
            if _degree > 360:
                _degree -= 360
            x1, y1, x2, y2 = self.x_y(angle=_degree, r1_=210, r2_=260)
            line_object(x1, y1, x2, y2, width=2)

    def draw_house_numbers(self):
        for i, j in enumerate(self.MIDPOINT_OF_HOUSES):
            if i == 11:
                midpoint = (self.MIDPOINT_OF_HOUSES[i] +
                            self.MIDPOINT_OF_HOUSES[0]) / 2
            else:
                if self.MIDPOINT_OF_HOUSES[i] == 360:
                    midpoint = self.MIDPOINT_OF_HOUSES[i + 1] / 2
                else:
                    if self.MIDPOINT_OF_HOUSES[i + 1] == 0 or \
                            self.MIDPOINT_OF_HOUSES[i + 1] < 30:
                        midpoint = (self.MIDPOINT_OF_HOUSES[i] +
                                    self.MIDPOINT_OF_HOUSES[i + 1] + 360) / 2
                    else:
                        midpoint = (self.MIDPOINT_OF_HOUSES[i] +
                                    self.MIDPOINT_OF_HOUSES[i + 1]) / 2
            x1, y1, x2, y2 = self.x_y(angle=midpoint, r1_=60, r2_=110)
            x = (x1 + x2) / 2
            y = (y1 + y2) / 2
            text_object(x=x, y=y, _text=f"{i + 1}")

    def draw_sign_symbols(self):
        asc = self.house_cusps()[1]
        for i, j in enumerate(self.SIGNS):
            end = 30 - (asc % 30) + 180
            start = end - 30
            start += (30 * i)
            end += (30 * i)
            if start > 360:
                start -= 360
            if end > 360:
                end -= 360
            if start > 330:
                midpoint = (start + end + 360) / 2
            else:
                midpoint = (start + end) / 2
            if midpoint > 360:
                midpoint -= 360
            x1, y1, x2, y2 = self.x_y(angle=midpoint, r1_=210, r2_=260)
            x = (x1 + x2) / 2
            y = (y1 + y2) / 2
            text_object(
                x=x,
                y=y,
                _text=self._SIGN_SYMBOLS[self.sign_pos()[1][i]],
                font="Arial 25",
                fill=self._SIGN_COLORS[self.sign_pos()[1][i]]
            )

    def modify_text_object(self, planet_symbol, key, value, offset):
        x1, y1, x2, y2 = self.x_y(angle=value, r1_=210, r2_=175)
        x = ((x1 + x2) / 2) + offset
        y = ((y1 + y2) / 2) + offset
        if key == "Mars" or key == "Venus":
            if os.name == "posix":
                text_object(
                    x=x,
                    y=y,
                    _text=f"{planet_symbol}",
                    width=0,
                    font="Arial 30"
                )
            elif os.name == "nt":
                text_object(
                    x=x,
                    y=y,
                    _text=f"{planet_symbol}",
                    width=0,
                    font="Arial 20"
                )
        else:
            text_object(
                x=x,
                y=y,
                _text=f"{planet_symbol}",
                width=0,
                font="Arial 20"
            )

    def draw_planets(self):
        count = 0
        planet_info_format = []
        house_info_format = []
        for key, value in self._PLANETS.items():
            planet = self.planet_pos_(value)
            split_degree_1 = planet[0].split(" ")
            degree_1, minute_1, second_1 = split_degree_1
            planet_info = (
                self.PLANET_SYMBOLS[key],
                key,
                degree_1,
                minute_1,
                second_1,
                self._SIGN_SYMBOLS[planet[1]],
                planet[1]
            )
            planet_info_format.append(planet_info)
            split_degree_2 = self.house_cusps_()[0][count][1].split(" ")
            degree_2, minute_2, second_2 = split_degree_2
            house_info = [
                self.house_cusps()[0][count][0],
                degree_2,
                minute_2,
                second_2,
                self._SIGN_SYMBOLS[self.house_cusps_()[0][count][-1]],
                self.house_cusps_()[0][count][-1]
            ]
            house_info_format.append(house_info)
            count += 1
            convert_planet_degree = dms_to_dd(
                self.planet_pos_(value)[0].replace("'", "")
                    .replace('"', "").replace("\u00b0", ""))
            asc = self.house_cusps_()[1]
            _signs_ = []
            for i, j in enumerate(self.SIGNS):
                end = 30 - (asc % 30) + 180
                start = end - 30
                start += (30 * i)
                end += (30 * i)
                if start > 360:
                    start -= 360
                if end > 360:
                    end -= 360
                _signs_.append(start)
            planet_index = self.sign_pos()[1].index(planet[1])
            new_degree_of_planet = convert_planet_degree + _signs_[planet_index]
            self.__PLANET_DEGREES.append(new_degree_of_planet)
        self.__PLANET_INFO_FORMAT = planet_info_format
        self.__HOUSE_INFO_FORMAT = house_info_format
        count1 = 0
        count2 = 0
        for i, j in zip(["Ascendant", "Medium Coeli"], ["Asc", "Mc"]):
            self._PLANETS[i] = 16 + count1
            self.PLANET_SYMBOLS[i] = j
            self.__PLANET_DEGREES.append(self.MIDPOINT_OF_HOUSES[count2])
            count1 += 3
            count2 += 9
        planet_degrees = {
            i: j for i, j in zip(self._PLANETS, self.__PLANET_DEGREES)
        }
        drawn_signs = []
        for key, value in planet_degrees.items():
            x1, y1, x2, y2 = self.x_y(angle=value, r1_=210, r2_=205)
            line_object(x1, y1, x2, y2, width=2, fill="red")
            for _key, _value in planet_degrees.items():
                aspect = value - _value
                if 0 < aspect < 4:
                    drawn_signs.append(key)
                    drawn_signs.append(_key)
                    self.modify_text_object(
                        planet_symbol=self.PLANET_SYMBOLS[key],
                        key=key, value=value, offset=6)
                elif -4 < aspect < 0:
                    self.modify_text_object(
                        planet_symbol=self.PLANET_SYMBOLS[key],
                        key=key, value=_value, offset=-6)
        drawn_signs = set(drawn_signs)
        for key, value in planet_degrees.items():
            if key not in drawn_signs:
                self.modify_text_object(
                    planet_symbol=self.PLANET_SYMBOLS[key],
                    key=key, value=value, offset=4)

    def create_aspect(self, planet_degrees, value, color, r1_=160, r2_=165):
        x1, y1, x2, y2 = self.x_y(angle=value, r1_=r1_, r2_=r2_)
        _x1, _y1, _x2, _y2 = self.x_y(angle=planet_degrees, r1_=r1_, r2_=r2_)
        line_object(x2, y2, _x2, _y2, width=2, fill=color)

    def select_aspect(self, aspect, value, planet_degrees, key, _key):
        if 0 < aspect < conjunction or 360 - conjunction < aspect < 360:
            self.CONJUNCTION.append(
                (key, self.ASPECT_SYMBOLS["Conjunction"], _key)
            )
            self.create_aspect(planet_degrees, value, color="red")
        elif 30 - semi_sextile < aspect < 30 + semi_sextile or \
                330 - semi_sextile < aspect < 330 + semi_sextile:
            self.SEMI_SEXTILE.append(
                (key, self.ASPECT_SYMBOLS["Semi-Sextile"], _key)
            )
            self.create_aspect(planet_degrees, value, color="black")
        elif 45 - semi_square < aspect < 45 + semi_square or \
                315 - semi_square < aspect < 315 + semi_square:
            self.SEMI_SQUARE.append(
                (key, self.ASPECT_SYMBOLS["Semi-Square"], _key)
            )
            self.create_aspect(planet_degrees, value, color="black")
        elif 60 - sextile < aspect < 60 + sextile or \
                300 - sextile < aspect < 300 + sextile:
            self.SEXTILE.append(
                (key, self.ASPECT_SYMBOLS["Sextile"], _key)
            )
            self.create_aspect(planet_degrees, value, color="blue")
        elif 72 - quintile < aspect < 72 + quintile or \
                288 - quintile < aspect < 288 + quintile:
            self.QUINTILE.append(
                (key, self.ASPECT_SYMBOLS["Quintile"], _key)
            )
            self.create_aspect(planet_degrees, value, color="purple")
        elif 90 - square < aspect < 90 + square or \
                270 - square < aspect < 270 + square:
            self.SQUARE.append(
                (key, self.ASPECT_SYMBOLS["Square"], _key)
            )
            self.create_aspect(planet_degrees, value, color="red")
        elif 120 - trine < aspect < 120 + trine or \
                240 - trine < aspect < 240 + trine:
            self.TRINE.append(
                (key, self.ASPECT_SYMBOLS["Trine"], _key)
            )
            self.create_aspect(planet_degrees, value, color="blue")
        elif 135 - sesquiquadrate < aspect < 135 + sesquiquadrate or \
                225 - sesquiquadrate < aspect < 225 + sesquiquadrate:
            self.SESQUIQUADRATE.append(
                (key, self.ASPECT_SYMBOLS["Sesquiquadrate"], _key)
            )
            self.create_aspect(planet_degrees, value, color="orange")
        elif 144 - biquintile < aspect < 144 + biquintile or \
                204 - biquintile < aspect < 204 + biquintile:
            self.BIQUINTILE.append(
                (key, self.ASPECT_SYMBOLS["BiQuintile"], _key)
            )
            self.create_aspect(planet_degrees, value, color="gray")
        elif 150 - quincunx < aspect < 150 + quincunx or \
                210 - quincunx < aspect < 210 + quincunx:
            self.QUINCUNX.append(
                (key, self.ASPECT_SYMBOLS["Quincunx"], _key)
            )
            self.create_aspect(planet_degrees, value, color="pink")
        elif 180 - opposite < aspect < 180 + opposite:
            self.OPPOSITE.append(
                (key, self.ASPECT_SYMBOLS["Opposite"], _key)
            )
            self.create_aspect(planet_degrees, value, color="red")
        else:
            self.NULL.append((key, self.ASPECT_SYMBOLS["Null"], _key))

    def draw_aspects(self):
        planet_degrees = {
            i: j for i, j in zip(self._PLANETS, self.__PLANET_DEGREES)
        }
        planet_degrees["Asc"] = self.MIDPOINT_OF_HOUSES[0]
        planet_degrees["Mc"] = self.MIDPOINT_OF_HOUSES[9]
        for key, value in planet_degrees.items():
            for _key, _value in planet_degrees.items():
                aspect = abs(value - _value)
                if key != _key:
                    self.select_aspect(aspect, value, _value, key, _key)

    def select_planet(self, i):
        for key, value in self.ASPECTS.items():
            if i[0] == key:
                value.append(i)

    def parse_aspects(self):
        for i in self.CONJUNCTION:
            self.select_planet(i)
        for i in self.SEMI_SEXTILE:
            self.select_planet(i)
        for i in self.SEMI_SQUARE:
            self.select_planet(i)
        for i in self.SEXTILE:
            self.select_planet(i)
        for i in self.QUINTILE:
            self.select_planet(i)
        for i in self.SQUARE:
            self.select_planet(i)
        for i in self.TRINE:
            self.select_planet(i)
        for i in self.SESQUIQUADRATE:
            self.select_planet(i)
        for i in self.BIQUINTILE:
            self.select_planet(i)
        for i in self.QUINCUNX:
            self.select_planet(i)
        for i in self.OPPOSITE:
            self.select_planet(i)
        for i in self.NULL:
            self.select_planet(i)

    def order_aspects(self):
        for planet in self.ASPECTS:
            planet_aspects = self.ASPECTS[planet]
            new_order = []
            save_aspects = {}
            for i, j in enumerate(planet_aspects):
                new_order.append(j[2])
                save_aspects[j[2]] = j[1]
            yield [(save_aspects[i], i) for i in
                   self._PLANETS.keys() if i in new_order]

    def new_order_of_the_aspects(self):
        for i, j in zip(self.order_aspects(), self.ASPECTS):
            self.ASPECTS[j] = i
            aspect_list = []
            for m in i:
                try:
                    if list(self.ASPECTS.keys()).index(j) < \
                            list(self.ASPECTS.keys()).index(m[1]):
                        aspect_list.append(m[0])
                except ValueError:
                    aspect_list.append(m[0])
            self.ASPECTS[j] = aspect_list
        self.ASPECTS["Ascendant"] = []
        self.ASPECTS["Medium Coeli"] = []

    def modify_info(self, k, x, m, i, count1, count2):
        if k == self._SIGN_SYMBOLS["Aries"] or \
                k == self._SIGN_SYMBOLS["Leo"] or \
                k == self._SIGN_SYMBOLS["Sagittarius"]:
            text_object(
                x=x + count1 + (m * 55),
                y=count2 + (i * 16),
                font="Arial 10",
                _text=k,
                fill="red"
            )
        elif k == self._SIGN_SYMBOLS["Taurus"] or \
                k == self._SIGN_SYMBOLS["Virgo"] or \
                k == self._SIGN_SYMBOLS["Capricorn"]:
            text_object(
                x=x + count1 + (m * 55),
                y=count2 + (i * 16),
                font="Arial 10",
                _text=k,
                fill="green"
            )
        elif k == self._SIGN_SYMBOLS["Gemini"] or \
                k == self._SIGN_SYMBOLS["Libra"] or \
                k == self._SIGN_SYMBOLS["Aquarius"]:
            text_object(
                x=x + count1 + (m * 55),
                y=count2 + (i * 16),
                font="Arial 10",
                _text=k,
                fill="yellow"
            )
        elif k == self._SIGN_SYMBOLS["Cancer"] or \
                k == self._SIGN_SYMBOLS["Scorpio"] or \
                k == self._SIGN_SYMBOLS["Pisces"]:
            text_object(
                x=x + count1 + (m * 55),
                y=count2 + (i * 16),
                font="Arial 10",
                _text=k,
                fill="blue"
            )
        elif k in self.SIGNS or "\u00b0" in k or \
                "'" in k or '"' in k or k in self._PLANETS:
            pass
        else:
            text_object(
                x=x + count1 + (m * 55),
                y=count2 + (i * 16),
                font="Arial 10",
                _text=k,
                fill="black"
            )

    def draw_planet_info(self, x=900):
        planet_symbols = ""
        __planets__ = ""
        degrees = ""
        minutes = ""
        seconds = ""
        sign_symbols = ""
        __signs__ = ""
        for i, j in enumerate(self.__PLANET_INFO_FORMAT):
            planet_symbols += f"{j[0]}\n"
            __planets__ += f"{j[1]}\n"
            degrees += f"{j[2]}\n"
            minutes += f"{j[3]}\n"
            seconds += f"{j[4]}\n"
            sign_symbols += f"{j[5]}\n"
            __signs__ += f"{j[6]}\n"
            for m, k in enumerate(j):
                if k == self.PLANET_SYMBOLS["Mars"] or \
                        k == self.PLANET_SYMBOLS["Venus"]:
                    if os.name == "posix":
                        text_object(
                            x=x + 10 + (m * 60),
                            y=15 + (i * 15),
                            font="Arial 15",
                            _text=k
                        )
                    elif os.name == "nt":
                        text_object(
                            x=x + 10 + (m * 60),
                            y=15 + (i * 15),
                            font="Arial 9",
                            _text=k
                        )
                else:
                    self.modify_info(
                        k=k, x=x, m=m, i=i, count1=10, count2=15
                    )
        text_object(x=x + 240, y=110, font="Arial 10", _text=__signs__)
        text_object(x=x + 60, y=110, font="Arial 10", _text=__planets__)
        text_object(x=x + 180, y=110, font="Arial 10", _text=seconds)
        text_object(x=x + 150, y=110, font="Arial 10", _text=minutes)
        text_object(x=x + 120, y=110, font="Arial 10", _text=degrees)

    def draw_house_info(self, x=900):
        __houses__ = ""
        degrees = ""
        minutes = ""
        seconds = ""
        sign_symbols = ""
        __signs__ = ""
        for i, j in enumerate(self.__HOUSE_INFO_FORMAT):
            if i == 0:
                __houses__ += "Asc\n"
            elif i == 3:
                __houses__ += "Ic\n"
            elif i == 6:
                __houses__ += "Dsc\n"
            elif i == 9:
                __houses__ += "Mc\n"
            else:
                __houses__ += f"House {j[0]}\n"
            degrees += f"{j[1]}\n"
            minutes += f"{j[2]}\n"
            seconds += f"{j[3]}\n"
            sign_symbols += f"{j[4]}\n"
            __signs__ += f"{j[5]}\n"
            for m, k in enumerate(j):
                if k == j[0]:
                    pass
                else:
                    self.modify_info(
                        k=k, x=x, m=m, i=i, count1=65, count2=220
                    )
        text_object(x=x + 30, y=315, font="Arial 10", _text=__houses__)
        text_object(x=x + 120, y=315, font="Arial 10", _text=degrees)
        text_object(x=x + 150, y=315, font="Arial 10", _text=minutes)
        text_object(x=x + 180, y=315, font="Arial 10", _text=seconds)
        text_object(x=x + 240, y=315, font="Arial 10", _text=__signs__)

    def draw_aspect_info(self, x=900):
        for i, j in enumerate(self.ASPECTS.items()):
            if i != len(self.ASPECTS.items()) - 1:
                line_object(
                    x1=x + 25 + (i * 25),
                    y1=428 + (i * 15),
                    x2=x + 25 + (i * 25),
                    y2=624,
                    width=1
                )
                line_object(
                    x1=x,
                    y1=428 + (i * 15),
                    x2=x + 25 + (i * 25),
                    y2=428 + (i * 15),
                    width=1
                )
            if j[0] == "Mars" or j[0] == "Venus":
                if os.name == "posix":
                    text_object(
                        x=x + 15 + (i * 25),
                        y=420 + (i * 15),
                        _text=self.PLANET_SYMBOLS[j[0]],
                        font="Arial 15"
                    )
                elif os.name == "nt":
                    text_object(
                        x=x + 15 + (i * 25),
                        y=420 + (i * 15),
                        _text=self.PLANET_SYMBOLS[j[0]],
                        font="Arial 10"
                    )
            else:
                text_object(
                    x=x + 15 + (i * 25),
                    y=420 + (i * 15),
                    _text=self.PLANET_SYMBOLS[j[0]],
                    font="Arial 10"
                )
            for k, m in enumerate(j[1]):
                if m == self.ASPECT_SYMBOLS["Sextile"] or \
                        m == self.ASPECT_SYMBOLS["Trine"]:
                    text_object(
                        x=x + 15 + (i * 25),
                        y=420 + (i * 15) + ((k + 1) * 15),
                        _text=m,
                        font="Arial 10",
                        fill="blue"
                    )
                elif m == self.ASPECT_SYMBOLS["Conjunction"] or \
                        m == self.ASPECT_SYMBOLS["Square"] or \
                        m == self.ASPECT_SYMBOLS["Opposite"]:
                    text_object(
                        x=x + 15 + (i * 25),
                        y=420 + (i * 15) + ((k + 1) * 15),
                        _text=m,
                        font="Arial 10",
                        fill="red"
                    )
                elif m == self.ASPECT_SYMBOLS["Semi-Sextile"] or \
                        m == self.ASPECT_SYMBOLS["Quincunx"]:
                    text_object(
                        x=x + 15 + (i * 25),
                        y=420 + (i * 15) + ((k + 1) * 15),
                        _text=m,
                        font="Arial 10",
                        fill="green"
                    )
                else:
                    text_object(
                        x=x + 15 + (i * 25),
                        y=420 + (i * 15) + ((k + 1) * 15),
                        _text=m,
                        font="Arial 10"
                    )


# ---------------------------------tkinter--------------------------------------


selected_categories, selected_ratings, displayed_results, \
    record_categories = [], [], [], []

toplevel1, toplevel2, menu, search_menu, listbox_menu = None, None, None, \
    None, None

record, on_toplevel, = False, False

_num_ = 0

master = tk.Tk()
master.title("TkAstroDb")
master.geometry(f"{master.winfo_screenwidth()}x{master.winfo_screenheight()}")
if os.name == "nt":
    master.iconbitmap(icon)

info_var = tk.StringVar()
info_var.set("0")

top_frame = tk.Frame(master=master)
top_frame.pack()

bottom_frame = tk.Frame(master=master)
bottom_frame.pack()

y_scrollbar = tk.Scrollbar(master=bottom_frame, orient="vertical")
y_scrollbar.pack(side="right", fill="y")

columns_1 = ["Adb ID", "Name", "Gender", "Rodden Rating", "Date",
             "Hour", "Julian Date", "Latitude", "Longitude", "Place",
             "Country", "Adb Link", "Category"]


def create_treeview(_master_, columns, height=18):
    treeview_ = Treeview(
        master=_master_,
        show="headings",
        columns=[f"#{_i_ + 1}" for _i_ in range(len(columns))],
        height=height
    )
    for _i_, _j_ in enumerate(columns):
        treeview_.heading(f"#{_i_ + 1}", text=_j_)
    treeview_.pack()
    return treeview_


treeview = create_treeview(_master_=bottom_frame, columns=columns_1)

entry_button_frame = tk.Frame(master=top_frame)
entry_button_frame.grid(row=0, column=0)

search_label = tk.Label(master=entry_button_frame, 
                        text="Search A Record By Name: ", fg="red")
search_label.grid(row=0, column=0, padx=5, sticky="w", pady=5)

search_entry = tk.Entry(master=entry_button_frame)
search_entry.grid(row=0, column=1, padx=5, pady=5)

found_record = tk.Label(master=entry_button_frame, text="")
found_record.grid(row=1, column=0, padx=5, pady=5)

add_button = tk.Button(master=entry_button_frame, text="Add")


def add_command(_record_):
    global _num_
    if _record_ in displayed_results:
        pass
    else:
        treeview.insert("", _num_, values=[col for col in _record_])
        _num_ += 1
        displayed_results.append(_record_)
        info_var.set(len(displayed_results))
    add_button.grid_forget()
    found_record.configure(text="")
    search_entry.delete("0", "end")


def search_func(event, _search_entry):
    master.update()
    save_record = ""
    count = 0
    for _record_ in database:
        if _search_entry.get() == _record_[1]:
            index = database.index(_record_)
            count += 1
            found_record.configure(text=f"Record Found = {count}")
            add_button.grid(row=1, column=1, padx=5, pady=5)
            add_button.configure(
                command=lambda: add_command(_record_=database[index]))
            save_record += database[index][1]
    if save_record != _search_entry.get() or \
            save_record == _search_entry.get() == "":
        found_record.configure(text="")
        add_button.grid_forget()


def destroy_menu(event, _menu_):
    if _menu_ is not None:
        _menu_.destroy()


def button_3_on_entry(event):
    global search_menu
    if search_menu is not None:
        destroy_menu(event, search_menu)
    search_menu = tk.Menu(master=None, tearoff=False)
    search_menu.add_command(
        label="Copy", 
        command=lambda: master.focus_get().event_generate('<<Copy>>'))
    search_menu.add_command(
        label="Cut", 
        command=lambda: master.focus_get().event_generate('<<Cut>>'))
    search_menu.add_command(
        label="Paste", 
        command=lambda: master.focus_get().event_generate('<<Paste>>'))
    search_menu.add_command(
        label="Remove", 
        command=lambda: master.focus_get().event_generate('<<Clear>>'))
    search_menu.add_command(
        label="Select All", 
        command=lambda: master.focus_get().event_generate('<<SelectAll>>'))
    search_menu.post(event.x_root, event.y_root)


def select_range(event):
    event.widget.select_range("0", "end")


search_entry.bind(
    "<Control-KeyRelease-a>", 
    lambda event: select_range(event))
search_entry.bind(
    "<Button-1>", 
    lambda event: destroy_menu(event, search_menu))
search_entry.bind(
    "<Button-3>", 
    lambda event: button_3_on_entry(event))
search_entry.bind(
    "<KeyRelease>", 
    lambda event: search_func(event, search_entry))

category_label = tk.Label(master=entry_button_frame, 
                          text="Categories:", fg="red")
category_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

rrating_label = tk.Label(master=entry_button_frame, 
                         text="Rodden Rating:", fg="red")
rrating_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")


rating_or_category = None


def on_frame_configure(event, tcanvas):
    tcanvas.configure(scrollregion=tcanvas.bbox("all"))


def tbutton_command(cvar_list, tlevel, select):
    for item in cvar_list:
        if item[0].get() is True:
            select.append(item[1])
    tlevel.destroy()


def check_all_command(check_all, cvar_list, checkbutton_list):
    if check_all.get() is True:
        for var, c_button in zip(cvar_list, checkbutton_list):
            var[0].set(True)
            c_button.configure(variable=var[0])
    else:
        for var, c_button in zip(cvar_list, checkbutton_list):
            var[0].set(False)
            c_button.configure(variable=var[0])


def select_ratings():
    global selected_ratings
    selected_ratings = []
    global toplevel2
    try:
        if not toplevel2.winfo_exists():
            toplevel2 = None
    except AttributeError:
        pass
    if toplevel2 is None:
        toplevel2 = tk.Toplevel()
        toplevel2.geometry("300x250")
        toplevel2.resizable(width=False, height=False)
        toplevel2.title("Select Rodden Ratings")
        if os.name == "nt":
            toplevel2.iconbitmap(icon)
        rating_frame = tk.Frame(master=toplevel2)
        rating_frame.pack(side="top")
        button_frame = tk.Frame(master=toplevel2)
        button_frame.pack(side="bottom")
        tbutton = tk.Button(master=button_frame, text="Apply")
        tbutton.pack()
        cvar_list = []
        checkbutton_list = []
        check_all = tk.BooleanVar()
        check_uncheck_ = tk.Checkbutton(master=rating_frame,
                                        text="Check/Uncheck All",
                                        variable=check_all)
        check_all.set(False)
        check_uncheck_.grid(row=0, column=0, sticky="nw")
        for num, c in enumerate(
                ["AA", "A", "B", "C", "DD", "X", "XX", "AX"], 1):
            _rating_ = c
            cvar = tk.BooleanVar()
            cvar_list.append([cvar, _rating_])
            checkbutton = tk.Checkbutton(master=rating_frame, 
                                         text=_rating_, variable=cvar)
            checkbutton_list.append(checkbutton)
            cvar.set(False)
            checkbutton.grid(row=num, column=0, sticky="nw")
        tbutton.configure(
            command=lambda: tbutton_command(cvar_list, 
                                            toplevel2, 
                                            selected_ratings))
        check_uncheck_.configure(
            command=lambda: check_all_command(check_all, 
                                              cvar_list, 
                                              checkbutton_list))


def select_categories():
    global selected_categories, record_categories, toplevel1
    load_database(database, category_dict)
    master.update()
    selected_categories, record_categories = [], []
    try:
        if not toplevel1.winfo_exists():
            toplevel1 = None
    except AttributeError:
        pass
    if toplevel1 is None:
        toplevel1 = tk.Toplevel()
        toplevel1.title("Select Categories")
        toplevel1.resizable(width=False, height=False)
        toplevel1.update()
        if os.name == "nt":
            toplevel1.iconbitmap(icon)
        canvas_frame = tk.Frame(master=toplevel1)
        canvas_frame.pack(side="top")
        button_frame = tk.Frame(master=toplevel1)
        button_frame.pack(side="bottom")
        tcanvas = tk.Canvas(master=canvas_frame)
        tframe = tk.Frame(master=tcanvas)
        tscrollbar = tk.Scrollbar(master=canvas_frame, 
                                  orient="vertical", 
                                  command=tcanvas.yview)
        tcanvas.configure(yscrollcommand=tscrollbar.set)
        tscrollbar.pack(side="right", fill="y")
        tcanvas.pack()
        tcanvas.create_window((4, 4), window=tframe, anchor="nw")
        tframe.bind(
            "<Configure>", 
            lambda event: on_frame_configure(event, tcanvas))
        tbutton = tk.Button(master=button_frame, text="Apply")
        tbutton.pack()
        cvar_list = []
        checkbutton_list = []
        check_all = tk.BooleanVar()
        check_uncheck_ = tk.Checkbutton(master=tframe,
                                        text="Check/Uncheck All",
                                        variable=check_all)
        check_all.set(False)
        check_uncheck_.grid(row=0, column=0, sticky="nw")
        for num, _category_ in enumerate(category_names, 1):
            master.update()
            cvar = tk.BooleanVar()
            cvar_list.append([cvar, _category_])
            checkbutton = tk.Checkbutton(master=tframe, 
                                         text=_category_, 
                                         variable=cvar)
            checkbutton_list.append(checkbutton)
            cvar.set(False)
            checkbutton.grid(row=num, column=0, sticky="nw")
        if record is False:
            tbutton.configure(
                command=lambda: tbutton_command(cvar_list, 
                                                toplevel1, 
                                                selected_categories))
        else:
            tbutton.configure(
                command=lambda: tbutton_command(cvar_list, 
                                                toplevel1, 
                                                record_categories))
        check_uncheck_.configure(
            command=lambda: check_all_command(check_all, 
                                              cvar_list, 
                                              checkbutton_list))
        master.update()


category_button = tk.Button(master=entry_button_frame, text="Select", 
                            command=select_categories)
category_button.grid(row=2, column=1, padx=5, pady=5)

rating_button = tk.Button(master=entry_button_frame, text="Select", 
                          command=select_ratings)
rating_button.grid(row=3, column=1, padx=5, pady=5)


def create_checkbutton():
    check_frame = tk.Frame(master=top_frame)
    check_frame.grid(row=0, column=2)
    for i, j in enumerate(
            ("event", "human", "male", "female", "North Hemisphere", 
             "South Hemisphere")):
        _var_ = tk.StringVar()
        _var_.set(value="0")
        _checkbutton_ = tk.Checkbutton(
            master=check_frame,
            text=f"Do not display {j} charts.",
            variable=_var_)
        _checkbutton_.grid(row=i, column=2, columnspan=2, sticky="w")
        yield _var_
        yield _checkbutton_


var_checkbutton_1, display_checkbutton_1, \
    var_checkbutton_2, display_checkbutton_2, \
    var_checkbutton_3, display_checkbutton_3, \
    var_checkbutton_4, display_checkbutton_4, \
    var_checkbutton_5, display_checkbutton_5, \
    var_checkbutton_6, display_checkbutton_6 = create_checkbutton()


def insert_to_treeview(item):
    global _num_
    treeview.insert("", _num_, values=[col for col in item])
    info_var.set(len(displayed_results))
    _num_ += 1
    displayed_results.append(item)
    master.update()


def south_north_check(item):
    if var_checkbutton_5.get() == "1" and var_checkbutton_6.get() == "0":
        if "n" in item[7]:
            pass
        else:
            insert_to_treeview(item)
    elif var_checkbutton_5.get() == "0" and var_checkbutton_6.get() == "1":
        if "s" in item[7]:
            pass
        else:
            insert_to_treeview(item)
    elif var_checkbutton_5.get() == "0" and var_checkbutton_6.get() == "0":
        insert_to_treeview(item)
    elif var_checkbutton_5.get() == "1" and var_checkbutton_6.get() == "1":
        pass


def male_female_check(item):
    if var_checkbutton_3.get() == "1" and var_checkbutton_4.get() == "0":
        if item[2] == "M":
            pass
        else:
            south_north_check(item)
    elif var_checkbutton_3.get() == "0" and var_checkbutton_4.get() == "1":
        if item[2] == "F":
            pass
        else:
            south_north_check(item)
    elif var_checkbutton_3.get() == "0" and var_checkbutton_4.get() == "0":
        south_north_check(item)
    elif var_checkbutton_3.get() == "1" and var_checkbutton_4.get() == "1":
        if item[2] == "F" or item[2] == "M":
            pass
        else:
            south_north_check(item)


def display_results():
    global displayed_results
    treeview.delete(*treeview.get_children())
    displayed_results = []
    count = 0
    b = open(file="65181 records.txt", mode="w", encoding="utf-8")
    for key, value in all_categories.items():
        if key[1] in selected_categories:
            for item in value:
                if item[3] in selected_ratings:
                    if item in displayed_results:
                        pass
                    else:
                        if var_checkbutton_1.get() == "0" and \
                                var_checkbutton_2.get() == "0":
                            if item[0] == 3546 or item[0] == 68092:
                                pass
                            else:
                                male_female_check(item)
                        elif var_checkbutton_1.get() == "1" and \
                                var_checkbutton_2.get() == "0":
                            if item[2] == "N/A":
                                pass
                            elif item[0] == 3546:
                                pass
                            else:
                                male_female_check(item)
                        elif var_checkbutton_1.get() == "0" and \
                                var_checkbutton_2.get() == "1":
                            if item[2] != "N/A" or item[0] == 68092:
                                pass
                            else:
                                south_north_check(item)
                        elif var_checkbutton_1.get() == "1" and \
                                var_checkbutton_2.get() == "1":
                            pass
    for i in displayed_results:
        b.write(f"{i}\n\n")
        b.flush()
    b.close()              
    info_var.set(len(displayed_results))
    master.update()
    if len(displayed_results) == 0:
        msgbox.showinfo(title="Display Records", 
                        message="No record is inserted.")
    elif len(displayed_results) == 1:
        msgbox.showinfo(title="Display Records", 
                        message="1 record is inserted.")
    else:
        msgbox.showinfo(title="Display Records", 
                        message=f"{len(displayed_results)} "
                                "records are inserted.")
    master.update()


def button_3_remove(_treeview):
    selected = _treeview.selection()
    if not selected:
        pass
    else:
        for i in selected:
            for j in displayed_results:
                if j[0] == _treeview.item(i)["values"][0]:
                    displayed_results.remove(j)
            _treeview.delete(i)
        info_var.set(len(displayed_results))


def button_3_open_url():
    focused = treeview.focus()
    if not focused:
        pass
    else:
        for i in displayed_results:
            if i[0] == treeview.item(focused)["values"][0]:
                webbrowser.open(i[11])


def button_3_open_chart(_treeview_):
    global open_chart, toplevel, canvas
    focused = _treeview_.focus()
    master.update()
    if not focused:
        pass
    else:
        if on_toplevel is False:
            _items_ = _treeview_.item(focused)["values"]
        else:
            _items_ = _treeview_.item(focused)["values"][2:]
        julian_date = float(_items_[6])
        latitude = _items_[7]
        if "n" in latitude:
            latitude = latitude.replace("n", "\u00b0") + "'0\""
            latitude = dms_to_dd(latitude)
        elif "s" in latitude:
            latitude = latitude.replace("s", "\u00b0") + "'0\""
            latitude = -1 * dms_to_dd(latitude)

        longitude = _items_[8]
        if "e" in longitude:
            longitude = longitude.replace("e", "\u00b0") + "'0\""
            longitude = dms_to_dd(longitude)
        elif "w" in longitude:
            longitude = longitude.replace("w", "\u00b0") + "'0\""
            longitude = -1 * dms_to_dd(longitude)
        open_chart = True
        toplevel = create_toplevel()
        canvas = create_canvas(toplevel)
        aspect_list = [
            f"{key} ({value})"
            for key, value in ASPECT_SYMBOLS.items() if key != "Null"
        ]
        orb_list = [
            conjunction,
            semi_sextile,
            semi_square,
            sextile,
            quintile,
            square,
            trine,
            sesquiquadrate,
            biquintile,
            quincunx,
            opposite
        ]
        info_frame_1 = tk.Frame(master=canvas, bg="white")
        info_frame_1.place(x=10, y=7)
        for i, j in enumerate(columns_1):
            if i == 0 or i == 6 or i == 10 or i == 11 or i == 12:
                pass
            elif i == 9:
                l1 = tk.Label(
                    master=info_frame_1, 
                    text=j, 
                    bg="white", 
                    font="Arial 10"
                )
                l1.grid(row=i, column=0, sticky="w")
                l2 = tk.Label(
                    master=info_frame_1, 
                    text="      :", 
                    bg="white",
                    font="Arial 10"
                )
                l2.grid(row=i, column=1, sticky="w")
                l3 = tk.Label(
                    master=info_frame_1, 
                    text=f"{_items_[i]}, {_items_[i + 1]}", 
                    bg="white",
                    font="Arial 10"
                )
                l3.grid(row=i, column=2, sticky="w")
            else:   
                l1 = tk.Label(
                    master=info_frame_1, 
                    text=j, 
                    bg="white", 
                    font="Arial 10"
                )
                l1.grid(row=i, column=0, sticky="w")
                l2 = tk.Label(
                    master=info_frame_1, 
                    text="      :", 
                    bg="white",
                    font="Arial 10"
                )
                l2.grid(row=i, column=1)
                l3 = tk.Label(
                    master=info_frame_1, 
                    text=_items_[i], 
                    bg="white",
                    font="Arial 10"
                )
                l3.grid(row=i, column=2, sticky="w")

        info_frame_2 = tk.Frame(master=canvas, bg="white")
        if os.name == "nt":
            line_object(x1=5, x2=260, y1=183, y2=183, width=1)
            info_frame_2.place(x=10, y=185)       
        else:
            line_object(x1=5, x2=260, y1=170, y2=170, width=1)
            info_frame_2.place(x=10, y=172)
        l4 = tk.Label(
            master=info_frame_2,
            text="House System      ",
            bg="white",
            font="Arial 10"
        )
        l4.grid(row=0, column=0, sticky="w")
        l5 = tk.Label(
            master=info_frame_2,
            text=":",
            bg="white",
            font="Arial 10"
        )
        l5.grid(row=0, column=1, sticky="w")
        l6 = tk.Label(
            master=info_frame_2,
            text=house_systems[hsys],
            bg="white",
            font="Arial 10"
        )
        l6.grid(row=0, column=2, sticky="w")
        info_frame_3 = tk.Frame(master=canvas, bg="white")
        if os.name == "nt":
            line_object(x1=5, x2=260, y1=210, y2=210, width=1)
            info_frame_3.place(x=10, y=215)
        else:
            line_object(x1=5, x2=260, y1=195, y2=195, width=1)        
            info_frame_3.place(x=10, y=197)
        for i, j in enumerate(aspect_list):
            l7 = tk.Label(
                master=info_frame_3,
                text=j,
                bg="white",
                font="Arial 10"
            )
            l7.grid(row=i, column=0, sticky="w")
            l8 = tk.Label(
                master=info_frame_3,
                text=":",
                bg="white",
                font="Arial 10"
            )
            l8.grid(row=i, column=1, sticky="w")
            l9 = tk.Label(
                master=info_frame_3,
                text=orb_list[i],
                bg="white",
                font="Arial 10"
            )
            l9.grid(row=i, column=2, sticky="w")
        Chart(
            julian_date,
            longitude,
            latitude
        )
        master.update()


def destroy(event):
    if menu is not None:
        menu.destroy()


def button_3_on_treeview(event):
    global menu, on_toplevel
    on_toplevel = False
    if menu is not None:
        destroy(event)
    menu = tk.Menu(master=None, tearoff=False)
    menu.add_command(
        label="Remove", 
        command=lambda: button_3_remove(treeview)
    )
    menu.add_command(
        label="Open ADB Page", 
        command=button_3_open_url
    )
    menu.add_command(
        label="Open Chart", 
        command=lambda: button_3_open_chart(treeview)
    )
    menu.post(event.x_root, event.y_root)


def select_all_tree(event):
    children = treeview.get_children()
    treeview.selection_set(children)


def delete_all_tree(event, _treeview_):
    button_3_remove(_treeview_)


treeview.bind("<Delete>", lambda event: delete_all_tree(event, treeview))
treeview.bind("<Control-a>", lambda event: select_all_tree(event))
treeview.bind("<Button-1>", lambda event: destroy(event))
treeview.bind("<Button-3>", lambda event: button_3_on_treeview(event))

display_button = tk.Button(master=top_frame, text="Display Records", 
                           command=display_results)
display_button.grid(row=10, column=0, columnspan=4, pady=10)


def x_scrollbar(y_scrl, _master_, _treeview_):
    y_scrl.configure(command=_treeview_.yview)
    x_scrl = tk.Scrollbar(master=_master_, orient="horizontal", 
                          command=_treeview_.xview)
    x_scrl.pack(side="top", fill="x")
    _treeview_.configure(xscrollcommand=x_scrl.set, 
                         yscrollcommand=y_scrl.set)
    return x_scrl


x_scrollbar(y_scrollbar, master, treeview)

info_frame = tk.Frame(master=master)
info_frame.pack(side="top")

total_info = tk.Label(master=info_frame, text="Total = ")
total_info.grid(row=0, column=0)

info = tk.Label(master=info_frame, textvariable=info_var)
info.grid(row=0, column=1)


# --------------------------------xlrd/xlwt-------------------------------------


conjunction = 6
semi_sextile = 2
semi_square = 2
sextile = 4
quintile = 2
square = 6
trine = 6
sesquiquadrate = 2
biquintile = 2
quincunx = 2
opposite = 6

r1, r2, r3, dir2 = "", "", "", ""

_r1, _r2, _r3 = [], [], []

method, selection = False, False

style = xlwt.XFStyle()
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER
alignment.vert = xlwt.Alignment.VERT_CENTER
style.alignment = alignment


def _font_(name="Arial", bold=False):
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    return font


def get_excel_datas(sheet):
    global r1, r2, r3, _r1, _r2, _r3, selection
    datas = []
    if r1.count("/") == 2 and r2.count("/") == 2 and r3.count("/") == 2:
        r1, r2, r3 = "", "", ""
    for row in range(sheet.nrows):
        if selection == "expected_values":
            if row == 1:
                r1 += f"{sheet.cell_value(row, 5)} / "
            elif row == 2:
                r2 += f"{sheet.cell_value(row, 5)} / "
            elif row == 3:
                r3 += f"{sheet.cell_value(row, 5)} / "
            elif row == 0 or row == 4 or row == 5 \
                    or row in [i for i in range(230, 396, 15)] or row == 411 \
                    or row == 427 \
                    or row in [i for i in range(443, 609, 15)] \
                    or row in [i for i in range(624, 790, 15)]:
                pass
            else:
                for col in range(sheet.ncols):
                    if col < 13:
                        datas.append(([row, col], sheet.cell_value(row, col)))
                    elif (col == 13 and 216 < row < 398) or \
                            (col == 13 and 429 < row < 789):
                        datas.append(([row, col], sheet.cell_value(row, col)))
        elif selection == "chi-square" or selection == "effect-size" or \
                selection == "cohens_d_effect":
            if row == 1:
                _r1.append(f"{sheet.cell_value(row, 5)}")
                for i in _r1:
                    if "/" in i:
                        r1 = i
            elif row == 2:
                _r2.append(f"{sheet.cell_value(row, 5)}")
                for i in _r2:
                    if "/" in i:
                        r2 = i
            elif row == 3:
                _r3.append(f"{sheet.cell_value(row, 5)}")
                for i in _r3:
                    if "/" in i:
                        r3 = i
            elif row == 0 or row == 4 or row == 5 \
                    or row in [i for i in range(230, 396, 15)] \
                    or row == 411 or row == 427 \
                    or row in [i for i in range(443, 609, 15)] \
                    or row in [i for i in range(624, 790, 15)]:
                pass
            else:
                for col in range(sheet.ncols):
                    if col < 13:
                        datas.append(([row, col], sheet.cell_value(row, col)))
                    elif (col == 13 and 216 < row < 398) or \
                            (col == 13 and 429 < row < 789):
                        datas.append(([row, col], sheet.cell_value(row, col)))
        elif selection == "observed":
            if row in [i for i in range(6)] \
                    or row in [i for i in range(230, 396, 15)] or row == 411 \
                    or row == 427 \
                    or row in [i for i in range(443, 609, 15)] \
                    or row in [i for i in range(624, 790, 15)]:
                pass
            else:
                for col in range(sheet.ncols):
                    if col < 13:
                        datas.append(([row, col], sheet.cell_value(row, col)))
                    elif (col == 13 and 216 < row < 398) or \
                            (col == 13 and 429 < row < 789):
                        datas.append(([row, col], sheet.cell_value(row, col)))
    return datas


def write_total_horz(sheet, num=7, col=13):
    for _ in range(12):
        v = _ + num
        if col == 13:
            sheet.write(
                v + 1, col,
                xlwt.Formula(
                    f"SUM(B{v + 2};C{v + 2};D{v + 2};E{v + 2};F{v + 2};"
                    f"G{v + 2};H{v + 2};I{v + 2};J{v + 2};K{v + 2};L{v + 2};"
                    f"M{v + 2})"),
                style=style)
        elif col == 14:
            sheet.write(
                v + 1, col,
                xlwt.Formula(
                    f"SUM(C{v + 2};D{v + 2};E{v + 2};F{v + 2};G{v + 2};"
                    f"H{v + 2};I{v + 2};J{v + 2};K{v + 2};L{v + 2};M{v + 2};"
                    f"N{v + 2})"),
                style=style)


def write_total_vert(sheet, num=7, col=1):
    if col == 1:
        for __, _ in enumerate("BCDEFGHIJKLMN"):
            sheet.write(
                num + 11, __ + col,
                xlwt.Formula(
                    f"SUM({_}{num};{_}{num + 1};{_}{num + 2};{_}{num + 3};"
                    f"{_}{num + 4};{_}{num + 5};{_}{num + 6};{_}{num + 7};"
                    f"{_}{num + 8};{_}{num + 9};{_}{num + 10};{_}{num + 11})"),
                style=style)
    elif col == 2:
        for __, _ in enumerate("CDEFGHIJKLMNO"):
            sheet.write(
                num + 11, __ + col,
                xlwt.Formula(
                    f"SUM({_}{num};{_}{num + 1};{_}{num + 2};{_}{num + 3};"
                    f"{_}{num + 4};{_}{num + 5};{_}{num + 6};{_}{num + 7};"
                    f"{_}{num + 8};{_}{num + 9};{_}{num + 10};{_}{num + 11})"),
                style=style)


def modify_category_names():
    cat = selected_categories[0].split(" : ")
    for i, j in enumerate(cat):
        if "/ " in j:
            cat[i] = cat[i].replace("/ ", "_", j.count("/"))
        if "/" in j:
            cat[i] = cat[i].replace("/", "_", j.count("/"))
        if " - " in j:
            cat[i] = cat[i].replace(" - ", "_", j.count(" - "))
        if "-" in j:
            cat[i] = cat[i].replace("-", "_", j.count("-"))
        if " " in j:
            cat[i] = cat[i].replace(" ", "_", j.count(" "))
    return cat


def write_title(sheet, row, var_checkbutton, label):
    style.font = _font_(bold=True)
    sheet.write_merge(r1=row, r2=row, c1=0, c2=1, label=f"{label}:", 
                      style=style)
    style.font = _font_(bold=False)
    if var_checkbutton.get() == "0":
        sheet.write(row, 2, label="True", style=style)
    else:
        sheet.write(row, 2, label="False", style=style)


def write_title_of_total(sheet):
    try:
        dis_res = displayed_results[0][1]
        record_name = f"{dis_res.replace(' ', '_', dis_res.count(' '))}"
    except IndexError:
        record_name = ""
    style.alignment = xlwt.Alignment()
    style.font = _font_(bold=True)
    sheet.write_merge(r1=0, r2=0, c1=3, c2=4, label="Adb Version:", 
                      style=style)
    style.font = _font_(bold=False)
    sheet.write_merge(r1=0, r2=0, c1=5, c2=13, 
                      label=f"{' / '.join(xml_files).replace('.xml', '')}", 
                      style=style)
    write_title(sheet, 0, var_checkbutton_1, "Event")
    write_title(sheet, 1, var_checkbutton_2, "Human")
    write_title(sheet, 2, var_checkbutton_3, "Male")
    write_title(sheet, 3, var_checkbutton_4, "Female")
    write_title(sheet, 4, var_checkbutton_5, "North Hemisphere")
    write_title(sheet, 5, var_checkbutton_6, "South Hemisphere")
    style.font = _font_(bold=True)
    if selection == "observed":
        sheet.write_merge(r1=1, r2=1, c1=3, c2=4, 
                          label="House System:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=1, r2=1, c1=5, c2=13, 
                          label=f"{house_systems[hsys]}", style=style)
        style.font = _font_(bold=True)
        if len(displayed_results) == 1:
            sheet.write_merge(r1=2, r2=2, c1=3, c2=4, 
                              label="Rodden Rating:", style=style)
            style.font = _font_(bold=False)
            sheet.write_merge(r1=2, r2=2, c1=5, c2=13, 
                              label=f"{displayed_results[0][3]}", 
                              style=style)
        else:
            sheet.write_merge(r1=2, r2=2, c1=3, c2=4, 
                              label="Rodden Rating:", style=style)
            style.font = _font_(bold=False)
            sheet.write_merge(r1=2, r2=2, c1=5, c2=13, 
                              label=f"{'+'.join(selected_ratings)}", 
                              style=style)
        style.font = _font_(bold=True)
        if len(selected_categories) == 1:
            if len(displayed_results) == 1:
                sheet.write_merge(r1=3, r2=3, c1=3, c2=4, 
                                  label="Category:", style=style)
                style.font = _font_(bold=False)
                sheet.write_merge(r1=3, r2=3, c1=5, c2=13, 
                                  label=f"{record_name}", style=style)
            else:
                sheet.write_merge(r1=3, r2=3, c1=3, c2=4, 
                                  label="Category:", style=style)
                style.font = _font_(bold=False)
                sheet.write_merge(
                    r1=3, r2=3, c1=5, c2=13,
                    label=f"{'/'.join(modify_category_names())}",
                    style=style)
        else:
            sheet.write_merge(r1=3, r2=3, c1=3, c2=4, 
                              label="Category:", style=style)
            style.font = _font_(bold=False)
            if len(selected_categories) < 100:
                sheet.write_merge(r1=3, r2=3, c1=5, c2=13, 
                                  label=" + ".join(selected_categories), 
                                  style=style)
            else:
                sheet.write_merge(r1=3, r2=3, c1=5, c2=13, 
                                  label="Control_Group", style=style)
    elif selection == "expected_values":
        sheet.write_merge(r1=1, r2=1, c1=3, c2=4, 
                          label="House System:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=1, r2=1, c1=5, c2=13, 
                          label=f"{r1[:-3]}", style=style)
        style.font = _font_(bold=True)
        sheet.write_merge(r1=2, r2=2, c1=3, c2=4, 
                          label="Rodden Rating:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=2, r2=2, c1=5, c2=13, 
                          label=f"{r2[:-3]}", style=style)
        style.font = _font_(bold=True)
        sheet.write_merge(r1=3, r2=3, c1=3, c2=4, 
                          label="Category:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=3, r2=3, c1=5, c2=13, 
                          label=f"{r3[:-3]}", style=style)
    elif selection == "chi-square" or selection == "effect-size" or \
             selection == "cohens_d_effect":
        sheet.write_merge(r1=1, r2=1, c1=3, c2=4, 
                          label="House System:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=1, r2=1, c1=5, c2=13, 
                          label=f"{r1}", style=style)
        style.font = _font_(bold=True)
        sheet.write_merge(r1=2, r2=2, c1=3, c2=4, 
                          label="Rodden Rating:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=2, r2=2, c1=5, c2=13, 
                          label=f"{r2}", style=style)
        style.font = _font_(bold=True)
        sheet.write_merge(r1=3, r2=3, c1=3, c2=4, 
                          label="Category:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=3, r2=3, c1=5, c2=13, 
                          label=f"{r3}", style=style)
    style.alignment = alignment
    style.font = _font_(bold=True)
    sheet.write(7, 13, "Total", style=style)
    sheet.write(21, 13, "Total", style=style)
    sheet.write(35, 13, "Total", style=style)
    for i in range(14):
        if i == 12:
            sheet.write(218 + (i * 15), 13, "Total", style=style)
            sheet.write(231 + (i * 15), 0, "Total", style=style)
        elif i == 13:
            sheet.write(219 + (i * 15), 13, "Total", style=style)
            sheet.write(232 + (i * 15), 0, "Total", style=style)
        else:
            sheet.write(217 + (i * 15), 14, "Total", style=style)
            sheet.write(430 + (i * 15), 14, "Total", style=style)
            sheet.write(611 + (i * 15), 14, "Total", style=style)
            sheet.write_merge(r1=230 + (i * 15), 
                              r2=230 + (i * 15), 
                              c1=0, 
                              c2=1, 
                              label="Total", 
                              style=style)
            sheet.write_merge(r1=443 + (i * 15), 
                              r2=443 + (i * 15), 
                              c1=0, 
                              c2=1, 
                              label="Total", 
                              style=style)
            sheet.write_merge(r1=624 + (i * 15), 
                              r2=624 + (i * 15), 
                              c1=0, 
                              c2=1, 
                              label="Total", 
                              style=style)
    style.font = _font_(bold=False)


def write_total_data(sheet):
    write_total_horz(sheet, 7)
    write_total_horz(sheet, 21)
    write_total_horz(sheet, 35)
    for i in range(14):
        if i == 12:
            write_total_horz(sheet, 218 + (i * 15))
            write_total_vert(sheet, 220 + (i * 15))
        elif i == 13:
            write_total_horz(sheet, 219 + (i * 15))
            write_total_vert(sheet, 221 + (i * 15))
        else:
            write_total_horz(sheet, 217 + (i * 15), 14)
            write_total_vert(sheet, 219 + (i * 15), 2)
            write_total_horz(sheet, 430 + (i * 15), 14)
            write_total_vert(sheet, 432 + (i * 15), 2)
            write_total_horz(sheet, 611 + (i * 15), 14)
            write_total_vert(sheet, 613 + (i * 15), 2)


def save_file(file, sheet, table_name, table0="table0", table1="table1"):
    write_total_data(sheet)
    os.remove(f"{table0}")
    os.remove(f"{table1}")
    file.save(f"{table_name}.xlsx")


def special_cells(new_sheet, i):
    style.font = _font_(bold=True)
    if "Orb" in i[1]:
        style.alignment = xlwt.Alignment()
        new_sheet.write_merge(r1=i[0][0], 
                              r2=i[0][0], 
                              c1=i[0][1], 
                              c2=i[0][1] + 2,
                              label=i[1], 
                              style=style)
        style.alignment = alignment
    elif i[1] == "Traditional House Rulership":
        new_sheet.write_merge(r1=397, r2=397, c1=0, c2=13,
                              label="Traditional House Rulership", 
                              style=style)
    elif i[1] == "Modern House Rulership":
        new_sheet.write_merge(r1=413, r2=413, c1=0, c2=13,
                              label="Modern House Rulership", 
                              style=style)
    elif i[1] == "Detailed Traditional House Rulership":
        new_sheet.write_merge(r1=429, r2=429, c1=0, c2=13,
                              label="Detailed Traditional House Rulership", 
                              style=style)
    elif i[1] == "Detailed Modern House Rulership":
        new_sheet.write_merge(r1=610, r2=610, c1=0, c2=13,
                              label="Detailed Modern House Rulership", 
                              style=style)
    elif i[0][0] in [i__ for i__ in range(218, 394, 15)] and i[0][1] == 0:
        new_sheet.write_merge(r1=i[0][0], r2=i[0][0] + 11, c1=0, c2=0,
                              label=i[1], style=style)
    elif i[0][0] in [i__ for i__ in range(431, 607, 15)] and i[0][1] == 0:
        new_sheet.write_merge(r1=i[0][0], r2=i[0][0] + 11, c1=0, c2=0,
                              label=i[1], style=style)
    elif i[0][0] in [i__ for i__ in range(612, 788, 15)] and i[0][1] == 0:
        new_sheet.write_merge(r1=i[0][0], r2=i[0][0] + 11, c1=0, c2=0,
                              label=i[1], style=style)
    else:
        new_sheet.write(*i[0], i[1], style=style)
    style.font = _font_(bold=False)


def create_a_new_table():
    file_1 = "table0.xlsx"
    file_2 = "table1.xlsx"
    read_1 = xlrd.open_workbook(file_1)
    read_2 = xlrd.open_workbook(file_2)
    sheet_1 = read_1.sheet_by_name("Sheet1")
    sheet_2 = read_2.sheet_by_name("Sheet1")
    new_file = xlwt.Workbook()
    new_sheet = new_file.add_sheet("Sheet1")
    data_1 = get_excel_datas(sheet_1)
    data_2 = get_excel_datas(sheet_2)
    control_list = []
    write_title_of_total(new_sheet)
    for i in data_1:
        for j in data_2:
            if i[0] == j[0]:
                if i[1] != "" and j[1] != "":
                    if type(i[1]) == float or type(j[1]) == float:
                        new_sheet.write(*i[0], i[1] + j[1], style=style)
                    else:
                        special_cells(new_sheet, i)
                elif i[1] != "" and j[1] == "":
                    new_sheet.write(*i[0], i[1], style=style)
                elif i[1] == "" and j[1] != "":
                    new_sheet.write(*j[0], j[1], style=style)
                if i[0] not in control_list:
                    control_list.append(i[0])
                if j[0] not in control_list:
                    control_list.append(j[0])
    save_file(new_file, new_sheet, "table0", file_1, file_2)


_planets_ = {i: [] for i in planets}
_planets = {i: [] for i in planets}


def extract_aspects(planet, value, num=11):
    if len(_planets_[f"{planet}"]) == num:
        _planets[f"{planet}"].append(_planets_[f"{planet}"])
        _planets_[f"{planet}"] = []
        _planets_[f"{planet}"].append(value)
    elif len(_planets_[f"{planet}"]) < num:
        _planets_[f"{planet}"].append(value)


def search_aspect(planet_pos, sheet, row: int, aspect: int, 
                  orb: int, name: str):
    _row = row
    n_ = 1
    name_frmt = f"{name}: Orb Factor: +- {orb}"
    style.alignment = xlwt.Alignment()
    sheet.write_merge(r1=_row - 2, r2=_row - 2, c1=0, c2=2, 
                      label=name_frmt, style=style)
    style.alignment = alignment
    for i, j in enumerate(planet_pos):
        style.font = _font_(bold=True)
        sheet.write(i + row - 1, i, j[0], style=style)
        for k in range(12 - (i + 1)):
            degree_1 = planet_pos[i][1]
            degree_2 = planet_pos[k + n_][1]
            if degree_1 < 30 and degree_2 > 210:
                degree_1 += 360
            elif 30 < degree_1 < 60 and degree_2 > 240:
                degree_1 += 360
            elif 60 < degree_1 < 90 and degree_2 > 270:
                degree_1 += 360
            elif 90 < degree_1 < 120 and degree_2 > 300:
                degree_1 += 360
            elif 120 < degree_1 < 150 and degree_2 > 330:
                degree_1 += 360
            elif 150 < degree_1 < 180 and 330 < degree_2 < 360:
                if degree_2 - degree_1 > 180:
                    degree_1 += 360
            elif degree_2 < 30 and degree_1 > 210:
                degree_2 += 360
            elif 30 < degree_2 < 60 and degree_1 > 240:
                degree_2 += 360
            elif 60 < degree_2 < 90 and degree_1 > 270:
                degree_2 += 360
            elif 90 < degree_2 < 120 and degree_1 > 300:
                degree_2 += 360
            elif 120 < degree_2 < 150 and degree_1 > 330:
                degree_2 += 360
            elif 150 < degree_2 < 180 and 330 < degree_1 < 360:
                if degree_1 - degree_2 > 180:
                    degree_2 += 360
            style.font = _font_(bold=False)
            if aspect - orb <= abs(degree_1 - degree_2) <= aspect + orb:
                sheet.write(k + _row, i, 1, style=style)
                for num, planet in enumerate(planets):
                    if j[0] == planet:
                        extract_aspects(planet, 1, 11 - num)
            else:
                sheet.write(k + _row, i, 0, style=style)
                for num, planet in enumerate(planets):
                    if j[0] == planet:
                        extract_aspects(planet, 0, 11 - num)
        n_ += 1
        _row += 1
    style.font = _font_(bold=True)


def sum_aspects(planet):
    for i, j in enumerate(planet):
        planet[i] = np.array(j)
    new_list = list(planet[0] + planet[1] + planet[2] + planet[3] + 
                    planet[4] + planet[5] + planet[6] + planet[7] + 
                    planet[8] + planet[9] + planet[10])
    return [float(i) for i in new_list]


def detailed_traditional(sheet, row, label, rulership):
    sheet.write_merge(r1=row, r2=row, c1=0, c2=13, label=label, style=style)
    item0 = 0
    item1 = 0
    lords = []
    for i, j in enumerate(rulership):
        if i % 2 == 0:
            lords.append(f"{j[1]} (+)")
        else:
            lords.append(f"{j[1]} (-)")
    for i in range(12):
        item2 = 0
        sheet.write_merge(
            r1=row + 2 + item0 + item2, 
            r2=row + 2 + item0 + item2 + 11, 
            c1=0, 
            c2=0,
            label=f"Lord {item1 + 1}\nis", 
            style=style)
        for lord in lords:
            sheet.write_merge(
                r1=row + 1 + item0, 
                r2=row + 1 + item0, 
                c1=item2 + 2, 
                c2=item2 + 2,
                label=f"House {item2 + 1}", 
                style=style)
            sheet.write(
                row + 2 + item0 + item2, 1,
                label=lord, style=style)
            item2 += 1
        item1 += 1
        item0 += 15


def detailed_values(sheet, lords, rulership, _row1=431, _row2=443):
    num = 0
    modify_rulership = {}
    for i, j in rulership.items():
        if num % 2 == 0:
            modify_rulership[i] = f"{j} (+)"
        else:
            modify_rulership[i] = f"{j} (-)"
        num += 1
    for keys, values in lords.items():
        step = 0
        for i, j in enumerate(signs):
            lord_name = f"Lord {i + 1}"
            if keys == lord_name:
                rows = [i for i in range(_row1 + step, _row2 + step, 1)]
                for _keys, _values in modify_rulership.items():
                    if _values in values[0]:
                        row = rows[signs.index(_keys)]
                        for k in range(12):
                            house = f"House {k + 1}"
                            if house == values[1]:
                                col = k + 2
                                sheet.write(row, col, 1, style=style)
                            else:
                                col = k + 2
                                sheet.write(row, col, 0, style=style)
                    else:
                        row = rows[signs.index(_keys)]
                        for _ in range(12):
                            sheet.write(row, _ + 2, 0, style=style)
            else:
                step += 15


def write_datas_to_excel(get_datas):
    global _planets_, _planets
    _planets_ = {i: [] for i in planets}
    _planets = {i: [] for i in planets}
    file = xlwt.Workbook()
    sheet = file.add_sheet("Sheet1")
    planet_info, house_info, planet_pos, traditional_ruler, modern_ruler, \
        trad_ruler_sign, mode_ruler_sign, trad_lords, mode_lords = get_datas
    style.font = _font_(bold=True)
    for i, j in enumerate(signs):
        sheet.write(7, i + 1, j, style=style)
        sheet.write(21, i + 1, j, style=style)
    for i, j in enumerate(planet_info):
        sheet.write(i + 8, 0, j[0], style=style)
        sheet.write(i + 36, 0, j[0], style=style)
    for i in range(12):
        sheet.write(i + 22, 0, f"Cusp {i + 1}", style=style)
        sheet.write(35, i + 1, f"House {i + 1}", style=style)
    style.font = _font_(bold=False)
    for i, j in enumerate(planet_info):
        for k, m in enumerate(signs):
            if j[1] == m:
                sheet.write(i + 8, k + 1, 1, style=style)
            else:
                sheet.write(i + 8, k + 1, 0, style=style)
    for i, j in enumerate(house_info):
        for k, m in enumerate(signs):
            if j[1] == m:
                sheet.write(i + 22, k + 1, 1, style=style)
            else:
                sheet.write(i + 22, k + 1, 0, style=style)
    __planets__ = [[] for _ in planets]
    for i, j in enumerate(planet_info):
        for k, m in enumerate(house_info):
            if j[-1] == m[0]:
                for _ in range(12):
                    if j[0] == planets[_]:
                        __planets__[_].append(j)
                sheet.write(i + 36, k + 1, 1, style=style)
            else:
                sheet.write(i + 36, k + 1, 0, style=style)
    style.font = _font_(bold=True)
    search_aspect(planet_pos, sheet, 51, 0, conjunction, "Conjunction")
    search_aspect(planet_pos, sheet, 65, 30, semi_sextile, "Semi-Sextile")
    search_aspect(planet_pos, sheet, 79, 45, semi_square, "Semi-Square")
    search_aspect(planet_pos, sheet, 93, 60, sextile, "Sextile")
    search_aspect(planet_pos, sheet, 107, 72, quintile, "Quintile")
    search_aspect(planet_pos, sheet, 121, 90, square, "Square")
    search_aspect(planet_pos, sheet, 135, 120, trine, "Trine")
    search_aspect(planet_pos, sheet, 149, 135, sesquiquadrate, "Sesquiquadrate")
    search_aspect(planet_pos, sheet, 163, 144, biquintile, "BiQuintile")
    search_aspect(planet_pos, sheet, 177, 150, quincunx, "Quincunx")
    search_aspect(planet_pos, sheet, 191, 180, opposite, "Opposite")
    style.alignment = xlwt.Alignment()
    sheet.write_merge(r1=203, r2=203, c1=0, c2=2, 
                      label="All Aspects", style=style)
    style.alignment = alignment
    for i, j in enumerate(planets):
        sheet.write(204 + i, i, j, style=style)
    style.font = _font_(bold=False)
    for i in planets:
        _planets[f"{i}"].append(_planets_[f"{i}"])
    for keys, values in _planets.items():
        try:
            _planets[f"{keys}"] = sum_aspects(values)
        except IndexError:
            _planets[f"{keys}"] = []
    col = 0
    n = 0
    for keys, values in _planets.items():
        row = 0
        for value in values:
            sheet.write(205 + row + n, col, value, style=style)
            row += 1
        col += 1
        n += 1
    new_order_of_planets = []
    for i in __planets__:
        house_group = [[] for _ in range(12)]
        for m in i:
            for _ in range(12):
                if m[2] == f"{_ + 1}":
                    house_group[_].append(m[1])
        new_order_of_planets.append(house_group)
    count = 0
    style.font = _font_(bold=True)
    for _ in new_order_of_planets:
        for o, p in enumerate(signs):
            sheet.write(217 + count, o + 2, p, style=style)
        count += 15
    count = 0
    for i in planets:
        sheet.write_merge(r1=218 + count, r2=218 + count + 11, c1=0, c2=0,
                          label=f"{i}\nin", style=style)
        for j, k in enumerate(houses):
            sheet.write(217 + count + (j + 1), 1,
                        label=f"House {k}", style=style)
        count += 15
    count = 0
    style.font = _font_(bold=False)
    for i in new_order_of_planets:
        for j, k in enumerate(i):
            for num, sign in enumerate(signs):
                if sign in k:
                    sheet.write(217 + count + (j + 1), num + 2, 1, 
                                style=style)
                else:
                    sheet.write(217 + count + (j + 1), num + 2, 0, 
                                style=style)
        count += 15
    style.font = _font_(bold=True)
    sheet.write_merge(r1=397, r2=397, c1=0, c2=13, 
                      label="Traditional House Rulership", style=style)
    for i, j in enumerate(traditional_ruler):
        style.font = _font_(bold=True)
        sheet.write(398, i + 1, f"House {i + 1}", style=style)
        sheet.write(399 + i, 0, j[0], style=style)
        for k, m in enumerate(houses):
            style.font = _font_(bold=False)
            if j[1] == f"House {m}":
                sheet.write(399 + i, k + 1, 1, style=style)
            else:
                sheet.write(399 + i, k + 1, 0, style=style)
    style.font = _font_(bold=True)
    sheet.write_merge(r1=413, r2=413, c1=0, c2=13, 
                      label="Modern House Rulership", style=style)
    for i, j in enumerate(modern_ruler):
        style.font = _font_(bold=True)
        sheet.write(414, i + 1, f"House {i + 1}", style=style)
        sheet.write(415 + i, 0, j[0], style=style)
        for k, m in enumerate(houses):
            style.font = _font_(bold=False)
            if j[1] == f"House {m}":
                sheet.write(415 + i, k + 1, 1, style=style)
            else:
                sheet.write(415 + i, k + 1, 0, style=style)
    style.font = _font_(bold=True)
    detailed_traditional(sheet, 429, 
                         "Detailed Traditional House Rulership",
                         traditional_rulership.items())
    detailed_traditional(sheet, 610, 
                         "Detailed Modern House Rulership",
                         modern_rulership.items())
    style.font = _font_(bold=False)
    detailed_values(sheet, trad_lords, 
                    traditional_rulership, 431, 443)
    detailed_values(sheet, mode_lords, 
                    modern_rulership, 612, 624)
    count = 0
    for i in os.listdir(os.getcwd()):
        if i.startswith("table"):
            count += 1
    write_title_of_total(sheet)
    write_total_data(sheet)
    file.save(f"table{count}.xlsx")
    if count == 1:
        create_a_new_table()


def change_dir(cat, dir1, orb_factor, sub_dir):
    return os.path.join(
        *cat,
        dir1,
        f"ORB_{'_'.join(orb_factor)}",
        f"{house_systems[hsys]}",
        sub_dir
    )


def check_dir_names(cat, dir1, orb_factor, criteria):
    if var_checkbutton_5.get() == "0" and var_checkbutton_6.get() == "0":
        return change_dir(cat, dir1, orb_factor, f"{criteria}")
    elif var_checkbutton_5.get() == "0" and var_checkbutton_6.get() == "1":
        return change_dir(cat, 
                          dir1, 
                          orb_factor, 
                          os.path.join(f"{criteria}", "North"))
    elif var_checkbutton_5.get() == "1" and var_checkbutton_6.get() == "0":
        return change_dir(cat, 
                          dir1, 
                          orb_factor, 
                          os.path.join(f"{criteria}", "South"))


def dir_names(cat, dir1, orb_factor):
    if len(displayed_results) == 1:
        return os.path.join(*cat, 
                            dir1, 
                            f"ORB_{'_'.join(orb_factor)}", 
                            f"{house_systems[hsys]}")
    if var_checkbutton_1.get() == "0" and \
            var_checkbutton_2.get() == "1":
        return check_dir_names(cat, 
                               dir1, 
                               orb_factor, 
                               "Event")
    elif var_checkbutton_1.get() == "1" and \
            var_checkbutton_2.get() == "0":
        if var_checkbutton_3.get() == "0" and \
                var_checkbutton_4.get() == "1":
            return check_dir_names(cat, 
                                   dir1, 
                                   orb_factor, 
                                   "Male")
        elif var_checkbutton_3.get() == "1" and \
                var_checkbutton_4.get() == "0":
            return check_dir_names(cat, 
                                   dir1, 
                                   orb_factor, 
                                   "Female")
        elif var_checkbutton_3.get() == "0" and \
                var_checkbutton_4.get() == "0":
            return check_dir_names(cat, 
                                   dir1, 
                                   orb_factor, 
                                   "Human")
    elif var_checkbutton_1.get() == "0" and \
            var_checkbutton_2.get() == "0":
        if var_checkbutton_3.get() == "0" and \
                var_checkbutton_4.get() == "1":
            return check_dir_names(cat, 
                                   dir1, 
                                   orb_factor, 
                                   "Male+Event")
        elif var_checkbutton_3.get() == "1" and \
                var_checkbutton_4.get() == "0":
            return check_dir_names(cat, 
                                   dir1, 
                                   orb_factor, 
                                   "Female+Event")
        elif var_checkbutton_3.get() == "0" and \
                var_checkbutton_4.get() == "0":
            return check_dir_names(cat, 
                                   dir1, 
                                   orb_factor, 
                                   "Human+Event")
        elif var_checkbutton_3.get() == "1" and \
                var_checkbutton_4.get() == "1":
            return check_dir_names(cat, 
                                   dir1, 
                                   orb_factor, 
                                   "Event")


def find_observed_values():
    global selection, dir2, open_chart
    selection = "observed"
    if len(displayed_results) == 0:
        msgbox.showinfo(
            title="Find Observed Values", 
            message="0 records selected."
        )
    else:
        __size__ = len(displayed_results)
        __received__ = 0
        __now__ = time.time()
        pframe = tk.Frame(master=master)
        pbar = Progressbar(master=pframe, orient="horizontal", length=200, 
                           mode="determinate")
        pstring = tk.StringVar()
        plabel = tk.Label(master=pframe, textvariable=pstring)
        pframe.pack()
        pbar.pack(side="left")
        plabel.pack(side="left")
        orb_factor = [
            conjunction, semi_sextile, semi_square, sextile, quintile, square,
            trine, sesquiquadrate, biquintile, quincunx, opposite
        ]
        orb_factor = [str(i) for i in orb_factor]
        log = open(file="output.log", mode="w", encoding="utf-8")
        log.write(
            f"Adb Version: {' / '.join(xml_files).replace('.xml', '')}\n"
        )
        log.write(f"House System: {house_systems[hsys]}\n")
        if len(displayed_results) == 1:
            log.write(f"Rodden Rating: {displayed_results[0][3]}\n")
        else:
            log.write(f"Rodden Rating: {'+'.join(selected_ratings)}\n")
        log.write(f"Orb Factor: {'_'.join(orb_factor)}\n")
        if len(selected_categories) == 1:
            if len(displayed_results) == 1:
                dis_res = displayed_results[0][1]
                cnt = dis_res.count(' ')
                dis_res = dis_res.replace(' ', '_', cnt)
                log.write(f"Category: {dis_res}\n\n")
            else:
                mod_cat_names = '/'.join(modify_category_names())
                log.write(f"Category: {mod_cat_names}\n\n")
        elif len(selected_categories) > 1:
            if len(selected_categories) < 100:
                log.write(f"Category: {' + '.join(selected_categories)}\n\n")
            else:                
                log.write("Category: Control_Group\n\n")
            criterias = ["Event", "Human", "Male", "Female",
                         "North Hemisphere", "South Hemisphere"]
            checkbuttons = [var_checkbutton_1, var_checkbutton_2,
                            var_checkbutton_3, var_checkbutton_4,
                            var_checkbutton_5, var_checkbutton_6]
            for i in range(6):
                if checkbuttons[i].get() == "0":
                    log.write(f"{criterias[i]}: True\n")
                else:
                    log.write(f"{criterias[i]}: False\n")
        log.write(f"|{str(dt.now())[:-7]}| Process started.\n\n")
        log.flush()
        for records in displayed_results:
            julian_date = float(records[6])
            latitude = records[7]
            longitude = records[8]
            if type(longitude) != float:
                if "e" in longitude:
                    longitude = float(longitude.replace("e", "."))
                elif "w" in longitude:
                    longitude = -1 * float(longitude.replace("w", "."))
            if type(latitude) != float:
                if "n" in latitude:
                    latitude = float(latitude.replace("n", "."))
                elif "s" in latitude:
                    latitude = -1 * float(latitude.replace("s", "."))
            try:
                open_chart = False
                chart = Chart(julian_date, longitude, latitude)
                write_datas_to_excel(chart.get_chart_data())
            except BaseException as err:
                log.write(f"|{str(dt.now())[:-7]}| Error Type: {err}\n"
                          f"{' ' * 22}Record: {records}\n\n")
                log.flush()
            __received__ += 1
            if __received__ != __size__:
                pbar["value"] = __received__
                pbar["maximum"] = __size__
                pstring.set("{} %, {} minutes remaining.".format(
                    int(100 * __received__ / __size__),
                    round(
                        (int(__size__ /
                             (__received__ / (time.time() - __now__)))
                         - int(time.time() - __now__)) / 60)))
            else:
                pframe.destroy()
                pbar.destroy()
                plabel.destroy()
                try:
                    os.rename("table0.xlsx", "observed_values.xlsx")
                    dir1 = f"RR_{'+'.join(selected_ratings)}"
                    if len(selected_categories) == 1:
                        if len(displayed_results) == 1:
                            dis_res = displayed_results[0][1]
                            cnt = dis_res.count(" ")
                            dis_res = dis_res.replace(" ", "_", cnt)
                            cat = [dis_res]
                            dir1 = f"RR_{displayed_results[0][3]}"
                        else:
                            cat = modify_category_names()
                        dir2 = dir_names(cat, dir1, orb_factor)
                    elif len(selected_categories) > 1:
                        if len(selected_categories) < 5:
                            cats = []
                            for i in selected_categories:
                                split_new_name = i.split(" : ")
                                new_name = ""
                                for j, k in enumerate(split_new_name):
                                    if j != len(split_new_name) - 1:
                                        new_name += f"{k[0]}_"
                                    else:
                                        new_name += k 
                                cats.append(new_name)                              
                            cat = ["_+_".join(cats)]
                        else:
                            cat = ["Control_Group"]
                        dir2 = dir_names(cat, dir1, orb_factor)
                    try:
                        os.makedirs(dir2)
                        shutil.move(
                            src=os.path.join(os.getcwd(),
                                             "observed_values.xlsx"),
                            dst=os.path.join(os.getcwd(),
                                             dir2,
                                             "observed_values.xlsx"))
                    except FileExistsError as err:
                        log.write(
                            f"|{str(dt.now())[:-7]}| "
                            f"Error Type: {err}\n\n")
                        log.flush()
                except FileNotFoundError:
                    pass
                master.update()
                log.write(f"|{str(dt.now())[:-7]}| Process finished.")
                log.close()
                shutil.move(
                    src=os.path.join(os.getcwd(), "output.log"),
                    dst=os.path.join(os.getcwd(), dir2, "output.log"))
                msgbox.showinfo(title="Find Observed Values",
                                message="Process finished successfully.")


def sum_of_row(table):
    total = 0
    for item in table:
        if item[0][0] == 8 and type(item[1]) == float:
            total += item[1]
    return total


def calculate(file_name_1, file_name_2, table_name, msg_title):
    global selection
    selection = table_name
    file_name_1 = f"{file_name_1}.xlsx"
    file_name_2 = f"{file_name_2}.xlsx"
    read_file_1 = None
    read_file_2 = None
    try:
        read_file_1 = xlrd.open_workbook(file_name_1)
    except FileNotFoundError:
        msgbox.showinfo(title=f"Find {msg_title} Values",
                        message=f"No such file or directory: "
                                f"'{file_name_1}'")
    try:
        read_file_2 = xlrd.open_workbook(file_name_2)
    except FileNotFoundError:
        msgbox.showinfo(title=f"Find {msg_title} Values",
                        message=f"No such file or directory: "
                                f"'{file_name_2}'")
    if read_file_1 is not None and read_file_2 is not None:
        sheet_1 = read_file_1.sheet_by_name("Sheet1")
        sheet_2 = read_file_2.sheet_by_name("Sheet1")
        new_file = xlwt.Workbook()
        new_sheet = new_file.add_sheet("Sheet1")
        data_1 = get_excel_datas(sheet_1)
        data_2 = get_excel_datas(sheet_2)
        control_list = []
        ratio = sum_of_row(data_2) / sum_of_row(data_1)
        sum_of_all = sum_of_row(data_2) + sum_of_row(data_1)
        write_title_of_total(new_sheet)
        for i in data_1:
            for j in data_2:
                if i[0] == j[0]:
                    if i[1] != "" and j[1] != "":
                        if type(i[1]) == float or type(j[1]) == float:
                            try:
                                if selection == "expected_values":
                                    if method is False:
                                        # Sjoerd's method
                                        new_sheet.write(
                                            *i[0],
                                            i[1] * ratio,
                                            style=style)
                                    elif method is True:
                                        # Flavia's method
                                        new_sheet.write(
                                            *i[0],
                                            sum_of_row(data_2) *
                                            (i[1] + j[1]) / sum_of_all,
                                            style=style)
                                elif selection == "chi-square":
                                    new_sheet.write(
                                        *i[0], 
                                        (i[1] - j[1]) ** 2 / j[1], 
                                        style=style)
                                elif selection == "effect-size":
                                    new_sheet.write(
                                        *i[0],
                                        i[1] / j[1],
                                        style=style)
                                elif selection == "cohens_d_effect":
                                    new_sheet.write(
                                        *i[0], 
                                        (i[1] - j[1]) / (((i[1] * 
                                        (1 - (i[1] / sum_of_row(data_1)))) ** 2
                                         + (j[1] * (1 - 
                                         (j[1] / sum_of_row(data_2))) ** 2)) 
                                         / 2) ** 0.5, 
                                        style=style)                                    
                            except ZeroDivisionError:
                                new_sheet.write(*i[0], 0, style=style)
                        else:
                            special_cells(new_sheet, i)
                    elif i[1] != "" and j[1] == "":
                        new_sheet.write(*i[0], i[1], style=style)
                    elif i[1] == "" and j[1] != "":
                        new_sheet.write(*j[0], j[1], style=style)
                    if i[0] not in control_list:
                        control_list.append(i[0])
                    if j[0] not in control_list:
                        control_list.append(j[0])
        save_file(new_file, new_sheet, table_name, file_name_1, file_name_2)
        master.update()
        msgbox.showinfo(title=f"Find {msg_title} Values", 
                        message="Process finished successfully.")


# -------------------------menu variables & functions---------------------------


add_or_edit = False
edit_or_search = False
treeviews = []
items = []

freq_frmt = [0, 2000, 100]


def func1():
    global r1, r2, r3, _r1, _r2, _r3
    r1, r2, r3 = "", "", ""
    _r1, _r2, _r3 = [], [], []
    t1 = threading.Thread(target=find_observed_values)
    t1.start()


def func2():
    global r1, r2, r3, _r1, _r2, _r3
    r1, r2, r3 = "", "", ""
    _r1, _r2, _r3 = [], [], []
    t2 = threading.Thread(
        target=calculate,
        args=("control_group", "observed_values",
              "expected_values", "Expected")
    )
    t2.start()


def func3():
    global r1, r2, r3, _r1, _r2, _r3
    r1, r2, r3 = "", "", ""
    _r1, _r2, _r3 = [], [], []
    t3 = threading.Thread(
        target=calculate,
        args=("observed_values", "expected_values",
              "chi-square", "Chi-Square")
    )
    t3.start()


def func4():
    global r1, r2, r3, _r1, _r2, _r3
    r1, r2, r3 = "", "", ""
    _r1, _r2, _r3 = [], [], []
    t4 = threading.Thread(
        target=calculate,
        args=("observed_values", "expected_values",
              "effect-size", "Effect Size")
    )
    t4.start()
    
    
def func5():
    global r1, r2, r3, _r1, _r2, _r3
    r1, r2, r3 = "", "", ""
    _r1, _r2, _r3 = [], [], []
    t5 = threading.Thread(
        target=calculate,
        args=("observed_values", "expected_values",
              "cohens_d_effect", "Cohen's d effect")
    )
    t5.start()


def set_method_to_false():
    global method
    method = False
    func2()


def set_method_to_true():
    global method
    method = True
    func2()


def change_orb_factors(parent, orb_entries):
    global conjunction, semi_sextile, semi_square, sextile
    global quintile, square, trine, sesquiquadrate, biquintile
    global quincunx, opposite
    conjunction = int(orb_entries[0].get())
    semi_sextile = int(orb_entries[1].get())
    semi_square = int(orb_entries[2].get())
    sextile = int(orb_entries[3].get())
    quintile = int(orb_entries[4].get())
    square = int(orb_entries[5].get())
    trine = int(orb_entries[6].get())
    sesquiquadrate = int(orb_entries[7].get())
    biquintile = int(orb_entries[8].get())
    quincunx = int(orb_entries[9].get())
    opposite = int(orb_entries[10].get())
    parent.destroy()


def choose_orb_factor():
    toplevel3 = tk.Toplevel()
    toplevel3.title("Orb Factor")
    toplevel3.resizable(width=False, height=False)
    if os.name == "nt":
        toplevel3.iconbitmap(icon)
    aspects = ["Conjunction", "Semi-Sextile", "Semi-Square",
               "Sextile", "Quintile", "Square", "Trine",
               "Sesquiquadrate", "BiQuintile", "Quincunx", "Opposite"]
    default_orbs = [conjunction, semi_sextile, semi_square, sextile,
                    quintile, square, trine, sesquiquadrate, biquintile,
                    quincunx, opposite]
    orb_entries = []
    for i, j in enumerate(aspects):
        aspect_label = tk.Label(master=toplevel3, text=f"{j}")
        aspect_label.grid(row=i, column=0, sticky="w")
        equal_to = tk.Label(master=toplevel3, text="=")
        equal_to.grid(row=i, column=1, sticky="e")
        orb_entry = tk.Entry(master=toplevel3, width=5)
        orb_entry.grid(row=i, column=2)
        orb_entry.insert(0, default_orbs[i])
        orb_entries.append(orb_entry)
    apply_button = tk.Button(
        master=toplevel3,
        text="Apply",
        command=lambda: change_orb_factors(parent=toplevel3,
                                           orb_entries=orb_entries))
    apply_button.grid(row=11, column=0, columnspan=3)


def change_hsys(parent, checkbuttons, _house_systems_):
    global hsys
    if checkbuttons[_house_systems_[0]][1].get() == "1":
        hsys = "P"
    elif checkbuttons[_house_systems_[1]][1].get() == "1":
        hsys = "K"
    elif checkbuttons[_house_systems_[2]][1].get() == "1":
        hsys = "O"
    elif checkbuttons[_house_systems_[3]][1].get() == "1":
        hsys = "R"
    elif checkbuttons[_house_systems_[4]][1].get() == "1":
        hsys = "C"
    elif checkbuttons[_house_systems_[5]][1].get() == "1":
        hsys = "E"
    elif checkbuttons[_house_systems_[6]][1].get() == "1":
        hsys = "W"
    parent.destroy()


def check_uncheck(checkbuttons, _house_systems_, j):
    for i in _house_systems_:
        if i != j:
            checkbuttons[i][1].set("0")
            checkbuttons[i][0].configure(variable=checkbuttons[i][1])


def create_hsys_checkbuttons():
    toplevel4 = tk.Toplevel()
    toplevel4.title("House System")
    toplevel4.geometry("200x200")
    toplevel4.resizable(width=False, height=False)
    if os.name == "nt":
        toplevel4.iconbitmap(icon)
    hsys_frame = tk.Frame(master=toplevel4)
    hsys_frame.pack(side="top")
    button_frame = tk.Frame(master=toplevel4)
    button_frame.pack(side="bottom")
    _house_systems_ = [values for keys, values in house_systems.items()]
    checkbuttons = {}
    for i, j in enumerate(_house_systems_):
        _var_ = tk.StringVar()
        if j == house_systems[hsys]:
            _var_.set(value="1")
        else:
            _var_.set(value="0")
        _checkbutton_ = tk.Checkbutton(
            master=hsys_frame,
            text=j,
            variable=_var_)
        _checkbutton_.grid(row=i, column=0, sticky="w")
        checkbuttons[j] = [_checkbutton_, _var_]
    checkbuttons["Placidus"][0].configure(
        command=lambda: check_uncheck(checkbuttons,
                                      _house_systems_,
                                      "Placidus"))
    checkbuttons["Koch"][0].configure(
        command=lambda: check_uncheck(checkbuttons,
                                      _house_systems_,
                                      "Koch"))
    checkbuttons["Porphyrius"][0].configure(
        command=lambda: check_uncheck(checkbuttons,
                                      _house_systems_,
                                      "Porphyrius"))
    checkbuttons["Regiomontanus"][0].configure(
        command=lambda: check_uncheck(checkbuttons,
                                      _house_systems_,
                                      "Regiomontanus"))
    checkbuttons["Campanus"][0].configure(
        command=lambda: check_uncheck(checkbuttons,
                                      _house_systems_,
                                      "Campanus"))
    checkbuttons["Equal"][0].configure(
        command=lambda: check_uncheck(checkbuttons,
                                      _house_systems_,
                                      "Equal"))
    checkbuttons["Whole Signs"][0].configure(
        command=lambda: check_uncheck(checkbuttons,
                                      _house_systems_,
                                      "Whole Signs"))
    apply_button = tk.Button(master=button_frame, text="Apply",
                             command=lambda: change_hsys(
                                 toplevel4,
                                 checkbuttons,
                                 _house_systems_))
    apply_button.pack()


def export_link():
    if len(displayed_results) > 0:
        with open(file="links.txt", mode="w", encoding="utf-8") as f:
            for i, j in enumerate(displayed_results):
                f.write(f"{i + 1}. {j[11]}\n")
        msgbox.showinfo(
            title="Export Links",
            message=f"{len(displayed_results)} links were exported.")
    else:
        msgbox.showinfo(
            title="Export Links",
            message="Please select and display records.")
        master.update()


def export_lat_frequency():
    latitude_freq_north = {
        f"{i}\u00b0 - {i + 1}\u00b0": [] for i in range(90)
    }
    latitude_freq_south = {
        f"{-i}\u00b0 - {-i - 1}\u00b0": [] for i in range(90)
    }
    latitudes = []
    if len(displayed_results) > 0:
        for item in displayed_results:
            latitude = item[7]
            if "n" in latitude:
                latitude = latitude.replace("n", "\u00b0") + "'0\""
                latitude = dms_to_dd(latitude)
            elif "s" in latitude:
                latitude = latitude.replace("s", "\u00b0") + "'0\""
                latitude = -1 * dms_to_dd(latitude)
            latitudes.append(latitude)
        for i in latitudes:
            for j in range(90):
                if j <= i < j + 1:
                    latitude_freq_north[
                        f"{j}\u00b0 - {j + 1}\u00b0"].append(i)
                elif -j - 1 <= i < -j:
                    latitude_freq_south[
                        f"{-j}\u00b0 - {-j - 1}\u00b0"].append(i)
        edit_latitude_freq_north = {
            keys: len(values)
            for keys, values in latitude_freq_north.items()
            if len(values) != 0
        }
        edit_latitude_freq_south = {
            keys: len(values)
            for keys, values in latitude_freq_south.items()
            if len(values) != 0
        }
        with open(file="latitude-frequency.txt", mode="w",
                  encoding="utf-8") as f:
            f.write("Latitude Intervals\n\n")
            for i, j in edit_latitude_freq_south.items():
                f.write(f"{i} = {j}\n")
            for i, j in edit_latitude_freq_north.items():
                f.write(f"{i} = {j}\n")
            f.write(f"\nMean Latitude = "
                    f"{dd_to_dms(sum(latitudes) / len(latitudes))}\n")
            f.write(f"\nTotal = {len(displayed_results)}")
            msgbox.showinfo(
                title="Export Latitude Frequency",
                message=f"{len(displayed_results)} "
                f"records were exported.")
            master.update()
    else:
        msgbox.showinfo(
            title="Export Latitude Frequency",
            message="Please select and display records.")
        master.update()


def year_frequency_command(parent, date_entries, years):
    min_, max_, step_ = date_entries[:]
    min_, max_, step_ = int(min_.get()), int(max_.get()), \
                        int(step_.get())
    freq_frmt[0], freq_frmt[1], freq_frmt[2] = min_, max_, step_
    if len(displayed_results) > 0:
        with open(file="year-frequency.txt", mode="w",
                  encoding="utf-8") as f:
            year_dict = {}
            count = 0
            for i in range(min_, max_ + 1, step_):
                year_dict[
                    (min_ + (count * step_),
                     min_ + (count * step_) + step_)
                ] = []
                count += 1
            for i in years:
                for keys, values in year_dict.items():
                    if keys[0] <= i < keys[1]:
                        year_dict[keys[0], keys[1]] += i,
            for keys, values in year_dict.items():
                f.write(f"{keys} = {len(values)}\n")
            f.write(f"Total = {len(displayed_results)}")
            parent.destroy()
            msgbox.showinfo(
                title="Export Year Frequency",
                message=f"{len(displayed_results)} "
                f"records were exported.")
    else:
        parent.destroy()
        msgbox.showinfo(
            title="Export Year Frequency",
            message="Please select and display records.")
        master.update()


def export_year_frequency():
    toplevel5 = tk.Toplevel()
    toplevel5.title("Year Frequency")
    toplevel5.geometry("200x100")
    toplevel5.resizable(width=False, height=False)
    if os.name == "nt":
        toplevel5.iconbitmap(icon)
    t5frame = tk.Frame(master=toplevel5)
    t5frame.pack()
    date_entries = []
    years = [int(i[4].split(" ")[2]) for i in displayed_results]
    if len(years) != 0:
        freq_frmt[0], freq_frmt[1], freq_frmt[2] = \
            min(years), max(years), 100
    for i, j in enumerate(("Minimum", "Maximum", "Step")):
        date_label = tk.Label(master=t5frame, text=j)
        date_label.grid(row=i, column=0, sticky="w")
        date_entry = tk.Entry(master=t5frame, width=5)
        date_entry.grid(row=i, column=1, sticky="w")
        date_entry.insert("1", f"{freq_frmt[i]}")
        date_entries.append(date_entry)
    apply_button = tk.Button(
        master=t5frame,
        text="Apply",
        command=lambda: year_frequency_command(
            toplevel5,
            date_entries,
            years)
    )
    apply_button.grid(row=3, column=0, columnspan=3)


def callback(event, url):
    webbrowser.open_new(url)


def url_open(args, **kwargs):
    return urllib.urlopen(args, cafile=certifi.where(), **kwargs)


def from_local_to_utc(year, month, day, hour, minute,
                      _lat, _lon):
    nominatim = Nominatim()
    nominatim.urlopen = url_open
    location = nominatim.reverse([_lat, _lon])[0]
    tzw = tzwhere.tzwhere()
    timezone = tzw.tzNameAt(_lat, _lon)
    local_zone = tz.gettz(timezone)
    utc_zone = tz.gettz("UTC")
    global_time = dt.strptime(
        f"{year}-{month}-{day} {hour}:{minute}:00",
        "%Y-%m-%d %H:%M:%S")
    local_time = global_time.replace(tzinfo=local_zone)
    utc_time = local_time.astimezone(utc_zone)
    if location.split(", ")[0].isnumeric():
        loc = location.split(", ")[2]
    else:
        loc = location.split(", ")[0]
    return utc_time.hour, utc_time.minute, utc_time.second, \
           loc, location.split(", ")[-1]


def julday(year, month, day, hour, minute, second):
    time1 = dt.strptime(f"{day}.{month}.{year}", "%d.%m.%Y")
    time2 = dt.strptime("15.10.1582", "%d.%m.%Y")

    def _julday_(calendar):
        julday_ = swe.julday(
            year,
            month,
            day,
            hour + (minute / 60) + (second / 3600),
            calendar
        )

        deltat = swe.deltat(julday_)
        return {
            "JD": round(julday_ + deltat, 6),
            "TT": round(deltat * 86400, 1)
        }

    if (time2 - time1).days > 0:
        return _julday_(swe.JUL_CAL)
    elif (time2 - time1).days < 0:
        return _julday_(swe.GREG_CAL)


def add(cat_entry, listbox, list_box):
    global record_categories
    for rec in record_categories:
        if rec not in listbox.get("0", "end"):
            listbox.insert("end", "".join(rec))
            list_box.append("".join(rec))
            master.update()
    if cat_entry.get() != "":
        if cat_entry.get() not in listbox.get("0", "end"):
            listbox.insert("end", cat_entry.get())
            list_box.append(cat_entry.get())
    record_categories = []
    cat_entry.delete("0", "end")


def cat_cmd():
    global record
    record = True
    select_categories()
    record = False


def delete(lb):
    for item in lb.curselection()[::-1]:
        lb.delete(item)


def button_3_on_listbox(event, lb):
    global listbox_menu
    if listbox_menu is not None:
        destroy_menu(event, listbox_menu)
    listbox_menu = tk.Menu(master=None, tearoff=False)
    listbox_menu.add_command(
        label="Remove", command=lambda: delete(lb))
    listbox_menu.post(event.x_root, event.y_root)


def select_set(event):
    event.widget.select_set("0", "end")


def delete_(event, lb):
    delete(lb)


def widget(_entries_, _listboxes_, _option_menu_, list_box,
           _frame, text, width, row1, col1, row2, col2):
    label = tk.Label(master=_frame, text=text, fg="red")
    if text == "Gender":
        label.grid(row=row1, column=col1, columnspan=3)
        menu_var = tk.StringVar()
        gender_menu = tk.OptionMenu(_frame,
                                    menu_var, "M", "F", "N/A")
        gender_menu.grid(row=row2, column=col2)
        _option_menu_.append(menu_var)
    elif text == "Add Category":
        label.grid(row=row1, column=col1, columnspan=3)
        cat_button = tk.Button(master=_frame, text="Select",
                               width=10, command=cat_cmd)
        cat_button.grid(row=1, column=0, columnspan=3)
        cat_entry = tk.Entry(master=_frame)
        cat_entry.grid(row=2, column=0, columnspan=3)
        cat_entry.bind(
            "<Control-KeyRelease-a>",
            lambda event: select_range(event))
        cat_entry.bind(
            "<Button-1>",
            lambda event: destroy_menu(event, search_menu))
        cat_entry.bind(
            "<Button-3>",
            lambda event: button_3_on_entry(event))
        _option_menu_.append(cat_entry)
        lbox_frame = tk.Frame(master=_frame)
        lbox_frame.grid(row=4, column=0, columnspan=3)
        listbox = tk.Listbox(master=lbox_frame, width=50,
                             selectmode="extended")
        listbox.pack(side="left")
        listbox.bind(
            "<Delete>",
            lambda event: delete_(event, listbox))
        listbox.bind(
            "<Control-a>",
            lambda event: select_set(event))
        listbox.bind(
            "<Button-1>",
            lambda event: destroy_menu(event, listbox_menu))
        listbox.bind(
            "<Button-3>",
            lambda event: button_3_on_listbox(event, listbox))
        lbox_y_sbar = tk.Scrollbar(master=lbox_frame,
                                   orient="vertical",
                                   command=listbox.yview)
        lbox_y_sbar.pack(side="left", fill="y")
        listbox.configure(yscrollcommand=lbox_y_sbar.set)
        _add_button = tk.Button(
            master=_frame,
            text="Add",
            command=lambda: add(cat_entry, listbox, list_box))
        _add_button.grid(row=3, column=0, columnspan=3)
        _listboxes_.append(listbox)
    elif text == "Rodden Rating":
        label.grid(row=row1, column=col1)
        menu_var = tk.StringVar()
        rodden_menu = tk.OptionMenu(
            _frame,
            menu_var,
            "AA",
            "A",
            "B",
            "C",
            "DD",
            "X",
            "XX",
            "AX")
        rodden_menu.grid(row=row2, column=col2)
        _option_menu_.append(menu_var)
    else:
        entry = tk.Entry(master=_frame, width=width)
        entry.grid(row=row2, column=col2)
        entry.bind(
            "<Control-KeyRelease-a>",
            lambda event: select_range(event))
        entry.bind(
            "<Button-1>",
            lambda event: destroy_menu(event, search_menu))
        entry.bind(
            "<Button-3>",
            lambda event: button_3_on_entry(event))
        _entries_.append(entry)
        label.grid(row=row1, column=col1)
    master.update()


def widgets(entries, listboxes, option_menu, list_box, frames):
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[0],
        "Name", 28, 0, 0, 1, 0)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[1],
        "Gender", 28, 0, 0, 1, 0)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[2],
        "Rodden Rating", 28, 0, 0, 1, 0)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[3],
        "Day", 2, 0, 0, 1, 0)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[3],
        "Month", 2, 0, 1, 1, 1)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[3],
        "Year", 4, 0, 2, 1, 2)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[4],
        "Hour", 2, 0, 0, 1, 0)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[4],
        "Minute", 2, 0, 1, 1, 1)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[5],
        "Latitude", 10, 0, 0, 1, 0)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[5],
        "Longitude", 10, 0, 1, 1, 1)
    widget(
        entries,
        listboxes,
        option_menu,
        list_box,
        frames[6],
        "Add Category", 28, 0, 0, 1, 0)


def create_frames(toplevel6):
    for i in range(8):
        frame = tk.Frame(toplevel6, bd=1, relief="sunken")
        frame.grid(row=i, column=0, pady=4, padx=4)
        master.update()
        yield frame


def get_record_data(toplevel6, _treeview_, entries, option_menu,
                    listboxes, list_box, data):
    global add_or_edit, edit_or_search
    name = entries[0].get()
    if option_menu[0].get() == "M":
        _gender = "M"
    elif option_menu[0].get() == "F":
        _gender = "F"
    else:
        _gender = "N/A"
    rr = option_menu[1].get()
    day = entries[1].get()
    month = entries[2].get()
    year = entries[3].get()
    hour = entries[4].get()
    minute = entries[5].get()
    _latitude_ = float(entries[6].get())
    _longitude_ = float(entries[7].get())
    try:
        date = dt.strptime(f"{year} {month} {day}", "%Y %m %d")
        _country_ = ""
        try:
            utc_hour, utc_minute, utc_second, _place, country_ = \
                from_local_to_utc(year, month, day, hour, minute,
                                  _latitude_, _longitude_)
            jd = julday(int(year), int(month), int(day), int(utc_hour),
                        int(utc_minute), int(utc_second))["JD"]
            latitude, longitude = "", ""
            if _latitude_ < 0:
                _latitude_ *= -1
                _degree = int(_latitude_)
                _minute = int((_latitude_ - _degree) * 60)
                latitude += f"{_degree}s{_minute}"
            elif _latitude_ > 0:
                _degree = int(_latitude_)
                _minute = int((_latitude_ - _degree) * 60)
                latitude += f"{_degree}n{_minute}"
            if _longitude_ < 0:
                _longitude_ *= -1
                _degree = int(_longitude_)
                _minute = int((_longitude_ - _degree) * 60)
                latitude += f"{_degree}w{_minute}"
            elif _longitude_ > 0:
                _degree = int(_longitude_)
                _minute = int((_longitude_ - _degree) * 60)
                longitude += f"{_degree}e{_minute}"
            _country = CountryInfo()
            for keys, values in _country.all().items():
                for k, v in values.items():
                    if k == "nativeName":
                        if country_ in v or v in country_:
                            _country_ += values["name"]
            if not all([name, _gender, rr, list_box]):
                msgbox.showinfo(
                    title="Create New Record",
                    message="Fill the empty fields.")
                master.update()
            else:
                now = dt.now()
                now_frmt = now.strftime("%d %B %Y %H:%M")
                select_from_data = [
                    list(_) for _ in cursor.execute("SELECT * FROM DATA")
                ]
                if add_or_edit is True:
                    no = data[0]
                else:
                    no = len(select_from_data) + 1
                lb = listboxes[0].get("0", "end")
                if len(list_box) > 1:
                    _record_data = [
                        no, now_frmt, "-", name, _gender, rr,
                        date.strftime("%d %B %Y"), f"{hour}:{minute}",
                        jd, latitude, _latitude_, longitude, _longitude_,
                        _place, _country_, "-", "|".join(lb)]
                else:
                    _record_data = [
                        no, now_frmt, "-", name, _gender, rr,
                        date.strftime("%d %B %Y"), f"{hour}:{minute}",
                        jd, latitude, _latitude_, longitude, _longitude_,
                        _place, _country_, "-", "".join(lb)]
                toplevel6.destroy()
                master.update()
                if _record_data not in select_from_data:
                    if add_or_edit is True:
                        names = col_names.split(", ")
                        if edit_or_search is False:
                            focused = _treeview_.focus()
                            _treeview_.delete(focused)
                        else:
                            _treeview_.delete(items[0])  
                        for j, k in enumerate(names):
                            if j < 3:
                                pass
                            else:
                                cursor.execute(
                                    f"UPDATE DATA SET {names[j]} = ? "
                                    f"WHERE no = ?",
                                    (_record_data[j], no))
                        modify = _record_data[:10] + \
                                 [_record_data[11]] + _record_data[13:]
                        for i in cursor.execute("SELECT * FROM DATA"):
                            if modify[0] == i[0]:
                                modify[1] = i[1]
                        try:
                            _treeview_.insert("", no - 1, values=modify)
                        except:
                            pass
                        msgbox.showinfo(
                            title="Edit New Record",
                            message="Record is edited.")
                    else:
                        cursor.execute(
                            f"INSERT INTO DATA VALUES("
                            f"{', '.join('?' * len(_record_data))})",
                            _record_data)
                        modify = _record_data
                        modify.pop(10)
                        modify.pop(11)
                        try:
                            _treeview_.insert("", no - 1, values=modify)
                        except:
                            pass
                        msgbox.showinfo(
                            title="Add New Record",
                            message="Record is added.")
                    master.update()
                    connect.commit()
                else:
                    msg = "This record is also stored in the database."
                    msgbox.showinfo(title="Add New Record",
                                    message=f"Error: {msg}")
                    master.update()
        except BaseException as err:
            msgbox.showinfo(title="Add New Record",
                            message=f"Error: {err}")
            master.update()
    except ValueError as err:
        msgbox.showinfo(title="Add New Record",
                        message=f"Error: {err}")
        master.update()


def record_panel(text):
    toplevel6 = tk.Toplevel()
    toplevel6.title(text)
    toplevel6.resizable(width=False, height=False)
    if os.name == "nt":
        toplevel6.iconbitmap(icon)
    frames = [i for i in create_frames(toplevel6)]
    entries = []
    listboxes = []
    list_box = []
    option_menu = []
    widgets(entries, listboxes, option_menu, list_box, frames)
    return toplevel6, frames, entries, listboxes, list_box, option_menu


def add_record():
    global add_or_edit
    add_or_edit = False
    toplevel6, frames, entries, listboxes, list_box, option_menu = \
        record_panel("Add Record")
    if len(treeviews) != 0:
        add_record_button = tk.Button(
            master=frames[7],
            text="Apply",
            command=lambda: get_record_data(
                toplevel6,
                treeviews[0],
                entries,
                option_menu,
                listboxes,
                list_box,
                None))
    else:
        add_record_button = tk.Button(
            master=frames[7],
            text="Apply",
            command=lambda: get_record_data(
                toplevel6,
                None,
                entries,
                option_menu,
                listboxes,
                list_box,
                None))
    add_record_button.grid(row=0, column=0)


def create_panel(entries, data, listboxes, list_box,
                 option_menu, frames, toplevel6, _treeview_):
    entries[0].insert("end", data[3])
    date = dt.strptime(f"{data[6]} {data[7]}", "%d %B %Y %H:%M")
    date_frmt = date.strftime("%d %m %Y %H %M")
    for i, j in enumerate(cursor.execute("SELECT * FROM DATA")):
        if data[0] == j[0]:
            entries[6].insert("end", j[10])
            entries[7].insert("end", j[12])
            master.update()
    for i, j in enumerate(date_frmt.split(" ")):
        entries[i + 1].insert("end", j)
        master.update()
    if "|" in data[-1]:
        for i in data[-1].split("|"):
            listboxes[0].insert("end", i)
            list_box.append(i)
            master.update()
    else:
        listboxes[0].insert("end", data[-1])
        list_box.append(data[-1])
        master.update()
    option_menu[0].set(data[4])
    option_menu[1].set(data[5])
    master.update()
    add_record_button = tk.Button(
        master=frames[7],
        text="Apply",
        command=lambda: get_record_data(
            toplevel6,
            _treeview_,
            entries,
            option_menu,
            listboxes,
            list_box,
            data))
    add_record_button.grid(row=0, column=0)


def edit_record(_treeview_):
    global add_or_edit
    add_or_edit = True
    focused = _treeview_.focus()
    if not focused:
        pass
    else:
        toplevel6, frames, entries, listboxes, list_box, option_menu = \
            record_panel("Edit Record")
        data = _treeview_.item(focused)["values"]
        create_panel(entries, data, listboxes, list_box, option_menu,
                     frames, toplevel6, _treeview_)
                     

def delete_record(_treeview_):
    focused = _treeview_.focus()
    no = _treeview_.item(focused)["values"][0]
    cursor.execute("DELETE FROM DATA WHERE no = ?", (no,))
    connect.commit()
    for i in _treeview_.get_children():
        _treeview_.delete(i)
    content = [i for i in cursor.execute("SELECT * FROM DATA")]
    for i, j in enumerate(content):
        cursor.execute("UPDATE DATA SET no = ? WHERE no = ?",
                       (i + 1, j[0]))
        connect.commit()
        modify = [k for k in j]
        modify[0] = i + 1
        modify.pop(10)
        modify.pop(11)
        _treeview_.insert("", i, values=modify)
        master.update()


def button_3_on_treeview_(event, _treeview_):
    global menu, on_toplevel
    on_toplevel = True
    if menu is not None:
        destroy(event)
    menu = tk.Menu(master=None, tearoff=False)
    menu.add_command(
        label="Edit", command=lambda: edit_record(_treeview_))
    menu.add_command(
        label="Delete", command=lambda: delete_record(_treeview_))
    menu.add_command(
        label="Open Chart",
        command=lambda: button_3_open_chart(_treeview_)
    )
    menu.post(event.x_root, event.y_root)


def search_record(event, search_entry_, _treeview_):
    global edit_or_search
    master.update()
    for _record_ in database:
        if search_entry_.get() == _record_[1]:
            for item in _treeview_.get_children():
                if _treeview_.item(item)["values"][3] == _record_[1]:
                    toplevel6, frames, entries, listboxes, list_box, \
                        option_menu = record_panel(text="Edit Record")
                    data = _treeview_.item(item)["values"]
                    items.append(item)
                    edit_or_search = True
                    create_panel(
                        entries, data, listboxes, list_box,
                        option_menu, frames, toplevel6, _treeview_
                    )
                    break
            break


def edit_and_delete():
    global add_or_edit, on_toplevel
    add_or_edit = True
    master.update()
    toplevel7 = tk.Toplevel()
    toplevel7.title("Edit Records")
    toplevel7.geometry("800x600")
    toplevel7.resizable(width=False, height=False)
    toplevel7.iconbitmap(icon)
    search_label_ = tk.Label(master=toplevel7,
                             text="Search A Record By Name", fg="red")
    search_label_.pack()
    search_entry_ = tk.Entry(master=toplevel7)
    search_entry_.pack()
    columns_2 = ["No", "Add Date"] + columns_1
    y_scrollbar_2 = tk.Scrollbar(master=toplevel7, orient="vertical")
    y_scrollbar_2.pack(side="right", fill="y")
    _treeview_ = create_treeview(_master_=toplevel7, columns=columns_2,
                                 height=26)
    if _treeview_ not in treeviews:
        treeviews.append(_treeview_)
    x_scrollbar(y_scrl=y_scrollbar_2, _master_=toplevel7,
                _treeview_=_treeview_)
    for i, j in enumerate(cursor.execute("SELECT * FROM DATA")):
        modify = [col for col in j]
        modify.pop(10)
        modify.pop(11)
        _treeview_.insert("", i, values=modify)
        master.update()
    _treeview_.bind(
        "<Button-3>",
        lambda event: button_3_on_treeview_(event, _treeview_))
    _treeview_.bind(
        "<Button-1>",
        lambda event: destroy(event))
    search_entry_.bind(
        "<KeyRelease>",
        lambda event: search_record(event, search_entry_, _treeview_))
    master.update()
       

def about():
    toplevel8 = tk.Toplevel()
    toplevel8.title("About TkAstroDb")
    if os.name == "nt":
        toplevel8.iconbitmap(icon)
    name = "TkAstroDb"
    version, _version = "Version:", __version__
    build_date, _build_date = "Built Date:", "21.12.2018"
    update_date, _update_date = "Update Date:", \
        dt.strftime(
            dt.fromtimestamp(os.stat(sys.argv[0]).st_mtime), 
            "%d.%m.%Y"
        )
    developed_by, _developed_by = "Developed By:", \
                                  "Tanberk Celalettin Kutlu"
    thanks_to, _thanks_to = "Special Thanks To:", \
                            "Alois Treindl, Flavia Minghetti, Sjoerd Visser"
    contact, _contact = "Contact:", "tckutlu@gmail.com"
    github, _github = "GitHub:", \
                      "https://github.com/dildeolupbiten/TkAstroDb"
    tframe1 = tk.Frame(master=toplevel8, bd="2", relief="groove")
    tframe1.pack(fill="both")
    tframe2 = tk.Frame(master=toplevel8)
    tframe2.pack(fill="both")
    tlabel_title = tk.Label(master=tframe1, text=name, font="Arial 25")
    tlabel_title.pack()
    for i, j in enumerate((version, build_date, update_date,
                           developed_by, thanks_to, contact, github)):
        tlabel_info_1 = tk.Label(master=tframe2, text=j,
                                 font="Arial 12", fg="red")
        tlabel_info_1.grid(row=i, column=0, sticky="w")
    for i, j in enumerate((_version, _build_date, _update_date,
                           _developed_by, _thanks_to, _contact, _github)):
        if j == _github:
            tlabel_info_2 = tk.Label(master=tframe2, text=j,
                                     font="Arial 12", fg="blue",
                                     cursor="hand2")
            url1 = "https://github.com/dildeolupbiten/TkAstroDb"
            tlabel_info_2.bind(
                "<Button-1>",
                lambda event: callback(event, url1))
        elif j == _contact:
            tlabel_info_2 = tk.Label(master=tframe2, text=j,
                                     font="Arial 12", fg="blue",
                                     cursor="hand2")
            url2 = "mailto://tckutlu@gmail.com"
            tlabel_info_2.bind(
                "<Button-1>",
                lambda event: callback(event, url2))
        else:
            tlabel_info_2 = tk.Label(master=tframe2, text=j,
                                     font="Arial 12")
        tlabel_info_2.grid(row=i, column=1, sticky="w")


def update():
    url_1 = "https://raw.githubusercontent.com/dildeolupbiten/" \
            "TkAstroDb/master/TkAstroDb.py"
    url_2 = "https://raw.githubusercontent.com/dildeolupbiten/" \
            "TkAstroDb/master/README.md"
    data_1 = urllib.urlopen(url=url_1,
                            context=ssl.SSLContext(ssl.PROTOCOL_SSLv23))
    data_2 = urllib.urlopen(url=url_2,
                            context=ssl.SSLContext(ssl.PROTOCOL_SSLv23))
    with open(file="TkAstroDb.py", mode="r", encoding="utf-8") as f:
        var_1 = [i.decode("utf-8") for i in data_1]
        var_2 = [i.decode("utf-8") for i in data_2]
        var_3 = [i for i in f]
        if var_1 == var_3:
            msgbox.showinfo(title="Update",
                            message="Program is up-to-date.")
        else:
            with open(file="README.md", mode="w", encoding="utf-8") \
                    as g:
                for i in var_2:
                    g.write(i)
                    g.flush()
            with open(file="TkAstroDb.py", mode="w", encoding="utf-8") \
                    as h:
                for i in var_1:
                    h.write(i)
                    h.flush()
                msgbox.showinfo(title="Update",
                                message="Program is updated.")              
                if os.name == "posix":
                    subprocess.Popen(["python3", "TkAstroDb.py"])
                    import signal
                    os.kill(os.getpid(), signal.SIGKILL)
                elif os.name == "nt":
                    subprocess.Popen(["python", "TkAstroDb.py"])
                    os.system(f"TASKKILL /F /PID {os.getpid()}")


# ------------------------------------menu--------------------------------------


menubar = tk.Menu(master=master)
master.configure(menu=menubar)

calculations_menu = tk.Menu(master=menubar, tearoff=False)
export_menu = tk.Menu(master=menubar, tearoff=False)
options_menu = tk.Menu(master=menubar, tearoff=False)
records_menu = tk.Menu(master=menubar, tearoff=False)
help_menu = tk.Menu(master=menubar, tearoff=False)

menubar.add_cascade(label="Calculations", menu=calculations_menu)
menubar.add_cascade(label="Export", menu=export_menu)
menubar.add_cascade(label="Options", menu=options_menu)
menubar.add_cascade(label="Records", menu=records_menu)
menubar.add_cascade(label="Help", menu=help_menu)

method_menu = tk.Menu(master=calculations_menu, tearoff=False)

calculations_menu.add_command(label="Find Observed Values",
                              command=func1)
calculations_menu.add_cascade(label="Find Expected Values",
                              menu=method_menu)
calculations_menu.add_command(label="Find Chi-Square Values",
                              command=func3)
calculations_menu.add_command(label="Find Effect Size Values",
                              command=func4)
calculations_menu.add_command(label="Find Cohen's D Effect Size Values",
                              command=func5)

method_menu.add_command(label="Flavia's method",
                        command=set_method_to_true)
method_menu.add_command(label="Sjoerd's method",
                        command=set_method_to_false)

export_menu.add_command(label="Adb Links",
                        command=export_link)
export_menu.add_command(label="Latitude Frequency",
                        command=export_lat_frequency)
export_menu.add_command(label="Year Frequency",
                        command=export_year_frequency)

options_menu.add_command(label="House System",
                         command=create_hsys_checkbuttons)
options_menu.add_command(label="Orb Factor",
                         command=choose_orb_factor)

records_menu.add_command(label="Add New Record",
                         command=add_record)
records_menu.add_command(label="Edit & Delete Records",
                         command=edit_and_delete)

help_menu.add_command(label="About",
                      command=about)
help_menu.add_command(label="Check for Updates",
                      command=update)


# -----------------------------------run----------------------------------------


if __name__ == "__main__":
    t0 = threading.Thread(target=master.mainloop)
    t0.run()
