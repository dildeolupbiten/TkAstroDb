#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__version__ = "1.2.3"

import os
import sys
import ssl
import time
import shutil
import threading
import traceback
import numpy as np
import webbrowser
import tkinter as tk
import xml.etree.ElementTree
import urllib.request as urllib
import tkinter.messagebox as msgbox

from datetime import datetime
from tkinter.ttk import Progressbar, Treeview

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
try:
    import swisseph as swe
except ModuleNotFoundError:
    if os.name == "posix":
        os.system("pip3 install pyswisseph")
    elif os.name == "nt":
        import platform
        path = os.path.join(os.getcwd(), "Eph", "Whl")
        if sys.version_info.minor == 6:
            if platform.architecture()[0] == "64bit":
                new_path = os.path.join(path, "pyswisseph-2.5.1.post0-cp36-cp36m-win_amd64.whl")
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "32bit":
                new_path = os.path.join(path, "pyswisseph-2.5.1.post0-cp36-cp36m-win32.whl")
                os.system(f"pip3 install {new_path}")
        elif sys.version_info.minor == 7:
            if platform.architecture()[0] == "64bit":
                new_path = os.path.join(path, "pyswisseph-2.5.1.post0-cp37-cp37m-win_amd64.whl")
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "32bit":
                new_path = os.path.join(path, "pyswisseph-2.5.1.post0-cp37-cp37m-win32.whl")
                os.system(f"pip3 install {new_path}")
    import swisseph as swe


# --------------------------------------------- xml ---------------------------------------------


xml_file = ""

database, all_categories, category_names = [], [], []

for _i in os.listdir(os.getcwd()):
    if _i.endswith("xml"):
        xml_file += _i

tree = xml.etree.ElementTree.parse(f"{xml_file}")
root = tree.getroot()

for _i in range(1000000):
    try:
        user_data = []
        for gender, roddenrating, bdata, adb_link, categories in zip(
                root[_i + 2][1].findall("gender"),
                root[_i + 2][1].findall("roddenrating"),
                root[_i + 2][1].findall("bdata"),
                root[_i + 2][2].findall("adb_link"),
                root[_i + 2][3].findall("categories")):
            _name = root[_i + 2][1][0].text
            sbdate_dmy = bdata[1].text
            sbtime = bdata[2].text
            jd_ut = bdata[2].get("jd_ut")
            lat = bdata[3].get("slati")
            lon = bdata[3].get("slong")
            place = bdata[3].text
            country = bdata[4].text
            sctr = bdata[4].get("sctr")
            category = [
                (categories[_j].get("cat_id"), categories[_j].text)
                for _j in range(len(categories))]
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
            user_data.append(sctr)
            user_data.append(adb_link.text)
            user_data.append(category)
        database.append(user_data)
    except IndexError:
        break

for _i in range(5000):
    _records_ = []
    category_groups = {}
    category_name = ""
    for j_ in database:
        for _k in j_[13]:
            if _k[0] == f"{_i}":
                _records_.append(j_)
                category_name = _k[1]
                if category_name is None:
                    category_name = "No Category Name"
    category_groups[(_i, category_name)] = _records_
    if not _records_:
        pass
    else:
        category_names.append(category_name)
        all_categories.append(category_groups)

category_names = sorted(category_names)


# --------------------------------------------- swisseph ---------------------------------------------


swe.set_ephe_path(os.path.join(os.getcwd(), "Eph"))

signs = [
    "Aries", "Taurus", "Gemini", "Cancer", "Leo", "Virgo",
    "Libra", "Scorpio", "Sagittarius", "Capricorn", "Aquarius", "Pisces"
]

planets = [
    "Sun", "Moon", "Mercury", "Venus", "Mars", "Jupiter", 
    "Saturn", "Uranus", "Neptune", "Pluto", "North Node", "Chiron"
]

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


class Chart:
    def __init__(self, julian_date: float, longitude: float, latitude: float):
        self.julian_date = julian_date
        self.longitude = longitude
        self.latitude = latitude

        self.SIGNS = {j: 30 * i for i, j in enumerate(signs)}
        self.PLANETS = {
            planets[0]: swe.SUN,
            planets[1]: swe.MOON,
            planets[2]: swe.MERCURY,
            planets[3]: swe.VENUS,
            planets[4]: swe.MARS,
            planets[5]: swe.JUPITER,
            planets[6]: swe.SATURN,
            planets[7]: swe.URANUS,
            planets[8]: swe.NEPTUNE,
            planets[9]: swe.PLUTO,
            planets[10]: swe.TRUE_NODE,
            planets[11]: swe.CHIRON
        }

        self.PLANET_INFO_FORMAT = []
        self.MODIFY_PLANET_INFO_FORMAT = []
        self.PLANET_DEGREES = []
        self.HOUSE_INFO_FORMAT = []
        self.HOUSE_SIGN = []
        self.PLANET_SIGN_HOUSE = []
        self.NEW_HOUSE_DEGREES = []

        self.organize_chart_data()
        self.convert_planet_degrees()
        self.convert_house_info()
        self.convert_house_degrees()
        self.find_planet_positions()

    @staticmethod
    def dd_to_dms(dd):
        degree = int(dd)
        minute = int((dd - degree) * 60)
        second = round(float((dd - degree - minute / 60) * 3600))
        return f"{degree}\u00b0 {minute}\' {second}\""

    @staticmethod
    def dms_to_dd(dms):
        dms = dms.replace("\u00b0", "").replace("\'", "").replace("\"", "")
        degree = int(dms.split(" ")[0])
        minute = float(dms.split(" ")[1]) / 60
        second = float(dms.split(" ")[2]) / 3600
        return degree + minute + second

    @staticmethod
    def convert_angle(angle):
        if 0 <= angle < 30:
            return angle, signs[0]
        elif 30 <= angle < 60:
            return angle - 30, signs[1]
        elif 60 <= angle < 90:
            return angle - 60, signs[2]
        elif 90 <= angle < 120:
            return angle - 90, signs[3]
        elif 120 <= angle < 150:
            return angle - 120, signs[4]
        elif 150 <= angle < 180:
            return angle - 150, signs[5]
        elif 180 <= angle < 210:
            return angle - 180, signs[6]
        elif 210 <= angle < 240:
            return angle - 210, signs[7]
        elif 240 <= angle < 270:
            return angle - 240, signs[8]
        elif 270 <= angle < 300:
            return angle - 270, signs[9]
        elif 300 <= angle < 330:
            return angle - 300, signs[10]
        elif 330 <= angle < 360:
            return angle - 330, signs[11]

    def planet_pos(self, planet):
        calc = self.convert_angle(swe.calc_ut(self.julian_date, planet)[0])
        return calc[0], calc[1]

    def house_cusps(self):
        global hsys
        house = []
        asc = 0
        angle = []
        for i, j in enumerate(swe.houses(
                self.julian_date, self.latitude, self.longitude, bytes(hsys.encode("utf-8")))[0]):
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
            planet_info_1 = [i[0], i[1] + self.SIGNS[i[2]], i[2]]
            planet_info_2 = [i[0], i[1] + self.SIGNS[i[2]]]
            self.MODIFY_PLANET_INFO_FORMAT.append(planet_info_1)
            self.PLANET_DEGREES.append(planet_info_2)

    def convert_house_info(self):
        for i in self.HOUSE_INFO_FORMAT:
            house_info = [i[0], i[2]]
            self.HOUSE_SIGN.append(house_info)

    def convert_house_degrees(self):
        for i in self.HOUSE_INFO_FORMAT:
            house_info = [i[0], float(i[1]) + self.SIGNS[i[2]], i[2]]
            self.NEW_HOUSE_DEGREES.append(house_info)

    def find_planet_positions(self):
        for i, j in enumerate(self.MODIFY_PLANET_INFO_FORMAT):
            for k, m in enumerate(self.NEW_HOUSE_DEGREES):
                try:
                    if self.NEW_HOUSE_DEGREES[k][1] - self.NEW_HOUSE_DEGREES[k + 1][1] > 180 \
                            and j[1] < self.NEW_HOUSE_DEGREES[k + 1][1]:
                        next_house_degree = self.NEW_HOUSE_DEGREES[k + 1][1] + 360
                        new_planet_degree = j[1] + 360
                        if self.NEW_HOUSE_DEGREES[k][1] < new_planet_degree < next_house_degree:
                            self.PLANET_SIGN_HOUSE.append([j[0], j[2], self.HOUSE_SIGN[k][0]])
                    elif self.NEW_HOUSE_DEGREES[k][1] - self.NEW_HOUSE_DEGREES[k + 1][1] > 180 \
                            and j[1] > self.NEW_HOUSE_DEGREES[k][1]:
                        next_house_degree = self.NEW_HOUSE_DEGREES[k + 1][1] + 360
                        if self.NEW_HOUSE_DEGREES[k][1] < j[1] < next_house_degree:
                            self.PLANET_SIGN_HOUSE.append([j[0], j[2], self.HOUSE_SIGN[k][0]])
                    if self.NEW_HOUSE_DEGREES[k][1] < j[1] < self.NEW_HOUSE_DEGREES[k + 1][1]:
                        self.PLANET_SIGN_HOUSE.append([j[0], j[2], self.HOUSE_SIGN[k][0]])
                except IndexError:
                    if self.NEW_HOUSE_DEGREES[k][1] - self.NEW_HOUSE_DEGREES[0][1] > 180 \
                            and j[1] < self.NEW_HOUSE_DEGREES[0][1]:
                        next_house_degree = self.NEW_HOUSE_DEGREES[0][1] + 360
                        new_planet_degree = j[1] + 360
                        if self.NEW_HOUSE_DEGREES[k][1] < new_planet_degree < next_house_degree:
                            self.PLANET_SIGN_HOUSE.append([j[0], j[2], self.HOUSE_SIGN[k][0]])
                    elif self.NEW_HOUSE_DEGREES[k][1] - self.NEW_HOUSE_DEGREES[0][1] > 180 \
                            and j[1] > self.NEW_HOUSE_DEGREES[k][1]:
                        next_house_degree = self.NEW_HOUSE_DEGREES[0][1] + 360
                        if self.NEW_HOUSE_DEGREES[k][1] < j[1] < next_house_degree:
                            self.PLANET_SIGN_HOUSE.append([j[0], j[2], self.HOUSE_SIGN[k][0]])
                    if self.NEW_HOUSE_DEGREES[k][1] < j[1] < self.NEW_HOUSE_DEGREES[0][1]:
                        self.PLANET_SIGN_HOUSE.append([j[0], j[2], self.HOUSE_SIGN[k][0]])

    def get_chart_data(self):
        traditional_ruler, modern_ruler = [], []
        for i, j in enumerate(self.house_cusps()[0]):
            traditional = traditional_rulership[self.house_cusps()[0][i][2]]
            modern = modern_rulership[self.house_cusps()[0][i][2]]
            for k, m in enumerate(self.PLANET_SIGN_HOUSE):
                if traditional == m[0]:
                    traditional_ruler.append([f"Lord {i + 1}", f"House {m[2]}"])
                if m[0] in modern:
                    modern_ruler.append([f"Lord {i + 1}", f"House {m[2]}"])
        return self.PLANET_SIGN_HOUSE, self.HOUSE_SIGN, self.PLANET_DEGREES, traditional_ruler, modern_ruler


# --------------------------------------------- tkinter ---------------------------------------------


selected_categories, selected_ratings, displayed_results = [], [], []

toplevel1, toplevel2, menu, search_menu = None, None, None, None

_num_ = 0

master = tk.Tk()
master.title("TkAstroDb")
master.geometry(f"{master.winfo_screenwidth()}x{master.winfo_screenheight()}")

info_var = tk.StringVar()
info_var.set("0")

top_frame = tk.Frame(master=master)
top_frame.pack()

bottom_frame = tk.Frame(master=master)
bottom_frame.pack()

y_scrollbar = tk.Scrollbar(master=bottom_frame, orient="vertical")
y_scrollbar.pack(side="right", fill="y")

columns = ["Adb ID", "Name", "Gender", "Rodden Rating", "Date",
           "Hour", "Julian Date", "Latitude", "Longitude", "Place",
           "Country", "Country Code", "Adb Link", "Category"]

treeview = Treeview(master=bottom_frame, show="headings",
                    columns=[f"#{_i_ + 1}" for _i_ in range(len(columns))], height=18)
for _i_, _j_ in enumerate(columns):
    treeview.heading(f"#{_i_ + 1}", text=_j_)
treeview.pack()

entry_button_frame = tk.Frame(master=top_frame)
entry_button_frame.grid(row=0, column=0)

search_label = tk.Label(master=entry_button_frame, text="Search A Record By Name: ", fg="red")
search_label.grid(row=0, column=0, padx=5, sticky="w", pady=5)

search_entry = tk.Entry(master=entry_button_frame)
search_entry.grid(row=0, column=1, padx=5, pady=5)

found_record = tk.Label(master=entry_button_frame, text="")
found_record.grid(row=1, column=0, padx=5, pady=5)

add_button = tk.Button(master=entry_button_frame, text="Add")


def add_command(record):
    global _num_
    if record in displayed_results:
        pass
    else:
        treeview.insert("", _num_, values=[col for col in record])
        _num_ += 1
        displayed_results.append(record)
        info_var.set(len(displayed_results))
    add_button.grid_forget()
    found_record.configure(text="")
    search_entry.delete("0", "end")


def search_func(event):
    master.update()
    save_record = ""
    count = 0
    for record in database:
        if search_entry.get() == record[1]:
            index = database.index(record)
            count += 1
            found_record.configure(text=f"Record Found = {count}")
            add_button.grid(row=1, column=1, padx=5, pady=5)
            add_button.configure(command=lambda: add_command(record=database[index]))
            save_record += database[index][1]
    if save_record != search_entry.get() or save_record == search_entry.get() == "":
        found_record.configure(text="")
        add_button.grid_forget()


def destroy_entry(event):
    if search_menu is not None:
        search_menu.destroy()


def button_3_on_entry(event):
    global search_menu
    if search_menu is not None:
        destroy_entry(event)
    search_menu = tk.Menu(master=None, tearoff=False)
    search_menu.add_command(
        label="Copy", command=lambda: master.focus_get().event_generate('<<Copy>>'))
    search_menu.add_command(
        label="Cut", command=lambda: master.focus_get().event_generate('<<Cut>>'))
    search_menu.add_command(
        label="Paste", command=lambda: master.focus_get().event_generate('<<Paste>>'))
    search_menu.add_command(
        label="Remove", command=lambda: master.focus_get().event_generate('<<Clear>>'))
    search_menu.add_command(
        label="Select All", command=lambda: master.focus_get().event_generate('<<SelectAll>>'))
    search_menu.post(event.x_root, event.y_root)


search_entry.bind("<Button-1>", lambda event: destroy_entry(event))
search_entry.bind("<Button-3>", lambda event: button_3_on_entry(event))
search_entry.bind("<KeyRelease>", search_func)

category_label = tk.Label(master=entry_button_frame, text="Categories:", fg="red")
category_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

rrating_label = tk.Label(master=entry_button_frame, text="Rodden Rating:", fg="red")
rrating_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")


def on_frame_configure(event, tcanvas):
    tcanvas.configure(scrollregion=tcanvas.bbox("all"))


def tbutton_command(cvar_list, toplevel, select):
    for item in cvar_list:
        if item[0].get() is True:
            select.append(item[1])
    toplevel.destroy()


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
        rating_frame = tk.Frame(master=toplevel2)
        rating_frame.pack(side="top")
        button_frame = tk.Frame(master=toplevel2)
        button_frame.pack(side="bottom")
        tbutton = tk.Button(master=button_frame, text="Apply")
        tbutton.pack()
        cvar_list = []
        checkbutton_list = []
        check_all = tk.BooleanVar()
        check_uncheck = tk.Checkbutton(master=rating_frame, text="Check/Uncheck All", variable=check_all)
        check_all.set(False)
        check_uncheck.grid(row=0, column=0, sticky="nw")
        for num, c in enumerate(["AA", "A", "B", "C", "DD", "X", "XX", "AX"], 1):
            _rating_ = c
            cvar = tk.BooleanVar()
            cvar_list.append([cvar, _rating_])
            checkbutton = tk.Checkbutton(master=rating_frame, text=_rating_, variable=cvar)
            checkbutton_list.append(checkbutton)
            cvar.set(False)
            checkbutton.grid(row=num, column=0, sticky="nw")
        tbutton.configure(command=lambda: tbutton_command(cvar_list, toplevel2, selected_ratings))
        check_uncheck.configure(command=lambda: check_all_command(check_all, cvar_list, checkbutton_list))


def select_categories():
    global selected_categories
    selected_categories = []
    global toplevel1
    try:
        if not toplevel1.winfo_exists():
            toplevel1 = None
    except AttributeError:
        pass
    if toplevel1 is None:
        toplevel1 = tk.Toplevel()
        toplevel1.title("Select Categories")
        toplevel1.resizable(width=False, height=False)
        canvas_frame = tk.Frame(master=toplevel1)
        canvas_frame.pack(side="top")
        button_frame = tk.Frame(master=toplevel1)
        button_frame.pack(side="bottom")
        tcanvas = tk.Canvas(master=canvas_frame)
        tframe = tk.Frame(master=tcanvas)
        tscrollbar = tk.Scrollbar(master=canvas_frame, orient="vertical", command=tcanvas.yview)
        tcanvas.configure(yscrollcommand=tscrollbar.set)
        tscrollbar.pack(side="right", fill="y")
        tcanvas.pack()
        tcanvas.create_window((4, 4), window=tframe, anchor="nw")
        tframe.bind("<Configure>", lambda event: on_frame_configure(event, tcanvas))
        tbutton = tk.Button(master=button_frame, text="Apply")
        tbutton.pack()
        cvar_list = []
        checkbutton_list = []
        check_all = tk.BooleanVar()
        check_uncheck = tk.Checkbutton(master=tframe, text="Check/Uncheck All", variable=check_all)
        check_all.set(False)
        check_uncheck.grid(row=0, column=0, sticky="nw")
        for num, _category_ in enumerate(category_names, 1):
            cvar = tk.BooleanVar()
            cvar_list.append([cvar, _category_])
            checkbutton = tk.Checkbutton(master=tframe, text=_category_, variable=cvar)
            checkbutton_list.append(checkbutton)
            cvar.set(False)
            checkbutton.grid(row=num, column=0, sticky="nw")
        tbutton.configure(command=lambda: tbutton_command(cvar_list,  toplevel1, selected_categories))
        check_uncheck.configure(command=lambda: check_all_command(check_all, cvar_list, checkbutton_list))


category_button = tk.Button(master=entry_button_frame, text="Select", command=select_categories)
category_button.grid(row=2, column=1, padx=5, pady=5)

rating_button = tk.Button(master=entry_button_frame, text="Select", command=select_ratings)
rating_button.grid(row=3, column=1, padx=5, pady=5)


def create_checkbutton():
    check_frame = tk.Frame(master=top_frame)
    check_frame.grid(row=0, column=2)
    for i, j in enumerate(("event", "human", "male", "female", "North Hemisphere", "South Hemisphere")):
        _var_ = tk.StringVar()
        _var_.set(value="0")
        _checkbutton_ = tk.Checkbutton(
            master=check_frame,
            text=f"Do not display {j} charts.",
            variable=_var_)
        _checkbutton_.grid(row=i, column=2, columnspan=2, sticky="w")
        yield _var_
        yield _checkbutton_
    

var_checkbutton_1, display_checkbutton_1, var_checkbutton_2, display_checkbutton_2, \
    var_checkbutton_3, display_checkbutton_3, var_checkbutton_4, display_checkbutton_4, \
    var_checkbutton_5, display_checkbutton_5, var_checkbutton_6, display_checkbutton_6 = create_checkbutton()


def insert_to_treeview(control_items, item):
    global _num_
    control_items.append(item)
    master.update()
    treeview.insert("", _num_, values=[col for col in item])
    info_var.set(len(displayed_results))
    _num_ += 1
    displayed_results.append(item)


def south_north_check(control_items, item):
    if var_checkbutton_5.get() == "1" and var_checkbutton_6.get() == "0":
        if "n" in item[7]:
            pass
        else:
            insert_to_treeview(control_items, item)
    elif var_checkbutton_5.get() == "0" and var_checkbutton_6.get() == "1":
        if "s" in item[7]:
            pass
        else:
            insert_to_treeview(control_items, item)
    elif var_checkbutton_5.get() == "0" and var_checkbutton_6.get() == "0":
        insert_to_treeview(control_items, item)
    elif var_checkbutton_5.get() == "1" and var_checkbutton_6.get() == "1":
        pass


def male_female_check(control_items, item):
    if var_checkbutton_3.get() == "1" and var_checkbutton_4.get() == "0":
        if item[2] == "M":
            pass
        else:
            south_north_check(control_items, item)
    elif var_checkbutton_3.get() == "0" and var_checkbutton_4.get() == "1":
        if item[2] == "F":
            pass
        else:
            south_north_check(control_items, item)
    elif var_checkbutton_3.get() == "0" and var_checkbutton_4.get() == "0":
        south_north_check(control_items, item)
    elif var_checkbutton_3.get() == "1" and var_checkbutton_4.get() == "1":
        if item[2] == "F" or item[2] == "M":
            pass
        else:
            south_north_check(control_items, item)


def display_results():
    global displayed_results
    treeview.delete(*treeview.get_children())
    displayed_results = []
    control_items = []
    for c in all_categories:
        _category_ = list(c.keys())[0][1]
        if _category_ in selected_categories:
            items = list(c.values())[0]
            for item in items:
                if item[3] in selected_ratings:
                    if item in control_items:
                        pass
                    else:
                        if var_checkbutton_1.get() == "0" and var_checkbutton_2.get() == "0":
                            if item[0] == 3546 or item[0] == 68092:
                                pass
                            else:
                                male_female_check(control_items, item)
                        elif var_checkbutton_1.get() == "1" and var_checkbutton_2.get() == "0":
                            if item[2] == "N/A":
                                pass
                            elif item[0] == 3546:
                                pass
                            else:
                                male_female_check(control_items, item)
                        elif var_checkbutton_1.get() == "0" and var_checkbutton_2.get() == "1":
                            if item[2] != "N/A" or item[0] == 68092:
                                pass
                            else:
                                south_north_check(control_items, item)
                        elif var_checkbutton_1.get() == "1" and var_checkbutton_2.get() == "1":
                            pass
    info_var.set(len(displayed_results))
    master.update()
    if len(displayed_results) == 0:
        msgbox.showinfo(title="Display Records", message="No record is inserted.")
    elif len(displayed_results) == 1:
        msgbox.showinfo(title="Display Records", message="1 record is inserted.")
    else:
        msgbox.showinfo(title="Display Records", message=f"{len(displayed_results)} records are inserted.")
    

def button_3_remove():
    focused = treeview.focus()
    if not focused:
        pass
    else:
        for i in displayed_results:
            if i[0] == treeview.item(focused)["values"][0]:
                displayed_results.remove(i)
        treeview.delete(focused)
        info_var.set(len(displayed_results))


def button_3_open_url():
    focused = treeview.focus()
    if not focused:
        pass
    else:
        for i in displayed_results:
            if i[0] == treeview.item(focused)["values"][0]:
                webbrowser.open(i[12])


def destroy(event):
    if menu is not None:
        menu.destroy()


def button_3_on_treeview(event):
    global menu
    if menu is not None:
        destroy(event)
    menu = tk.Menu(master=None, tearoff=False)
    menu.add_command(
        label="Remove", command=button_3_remove)
    menu.add_command(
        label="Open ADB Page", command=button_3_open_url)
    menu.post(event.x_root, event.y_root)


treeview.bind("<Button-1>", lambda event: destroy(event))
treeview.bind("<Button-3>", lambda event: button_3_on_treeview(event))

display_button = tk.Button(master=top_frame, text="Display Records", command=display_results)
display_button.grid(row=10, column=0, columnspan=4, pady=10)

y_scrollbar.configure(command=treeview.yview)
x_scrollbar = tk.Scrollbar(master=master, orient="horizontal", command=treeview.xview)
x_scrollbar.pack(side="top", fill="x")

treeview.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)

info_frame = tk.Frame(master=master)
info_frame.pack(side="top")

total_info = tk.Label(master=info_frame, text="Total = ")
total_info.grid(row=0, column=0)

info = tk.Label(master=info_frame, textvariable=info_var)
info.grid(row=0, column=1)


# --------------------------------------------- xlrd/xlwt ---------------------------------------------


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
        if selection == "expected":
            if row == 1:
                r1 += f"{sheet.cell_value(row, 5)} / "
            elif row == 2:
                r2 += f"{sheet.cell_value(row, 5)} / "
            elif row == 3:
                r3 += f"{sheet.cell_value(row, 5)} / "
            elif row == 0 or row == 4 or row == 5 or row == 230 or row == 245 or row == 260 \
                    or row == 275 or row == 290 or row == 305 or row == 320 \
                    or row == 335 or row == 350 or row == 365 or row == 380 or row == 395 or row == 411 \
                    or row == 427:
                pass
            else:
                for col in range(sheet.ncols):
                    if col < 13:
                        datas.append(([row, col], sheet.cell_value(row, col)))
                    elif col == 13 and 216 < row < 398:
                        datas.append(([row, col], sheet.cell_value(row, col)))
        elif selection == "chisquare" or selection == "effectsize":
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
            elif row == 0 or row == 4 or row == 5 or row == 230 or row == 245 or row == 260 \
                    or row == 275 or row == 290 or row == 305 or row == 320 \
                    or row == 335 or row == 350 or row == 365 or row == 380 or row == 395 or row == 411 \
                    or row == 427:
                pass
            else:
                for col in range(sheet.ncols):
                    if col < 13:
                        datas.append(([row, col], sheet.cell_value(row, col)))
                    elif col == 13 and 216 < row < 398:
                        datas.append(([row, col], sheet.cell_value(row, col)))
        elif selection == "observed":  
            if row == 0 or row == 1 or row == 2 or row == 3 or row == 4 or row == 5 \
                    or row == 230 or row == 245 or row == 260 \
                    or row == 275 or row == 290 or row == 305 or row == 320 \
                    or row == 335 or row == 350 or row == 365 or row == 380 or row == 395 or row == 411 \
                    or row == 427:
                pass
            else:
                for col in range(sheet.ncols):
                    if col < 13:
                        datas.append(([row, col], sheet.cell_value(row, col)))
                    elif col == 13 and 216 < row < 398:
                        datas.append(([row, col], sheet.cell_value(row, col)))
    return datas


def write_total_horz(sheet, num=7, col=13):
    for _ in range(12):
        v = _ + num
        if col == 13:
            sheet.write(
                v + 1, col,
                xlwt.Formula(
                    f"SUM(B{v + 2};C{v + 2};D{v + 2};E{v + 2};F{v + 2};G{v + 2};H{v + 2};\
                    I{v + 2};J{v + 2};K{v + 2};L{v + 2};M{v + 2})"),
                style=style)
        elif col == 14:
            sheet.write(
                v + 1, col,
                xlwt.Formula(
                    f"SUM(C{v + 2};D{v + 2};E{v + 2};F{v + 2};G{v + 2};H{v + 2};I{v + 2};\
                    J{v + 2};K{v + 2};L{v + 2};M{v + 2};N{v + 2})"),
                style=style)


def write_total_vert(sheet, num=7, col=1):
    if col == 1:
        for __, _ in enumerate("BCDEFGHIJKLMN"):
            sheet.write(
                num + 11, __ + col,
                xlwt.Formula(
                    f"SUM({_}{num};{_}{num + 1};{_}{num + 2};{_}{num + 3};{_}{num + 4};{_}{num + 5};\
                            {_}{num + 6};{_}{num + 7};{_}{num + 8};{_}{num + 9};{_}{num + 10};{_}{num + 11})"),
                style=style)
    elif col == 2:
        for __, _ in enumerate("CDEFGHIJKLMNO"):
            sheet.write(
                num + 11, __ + col,
                xlwt.Formula(
                    f"SUM({_}{num};{_}{num + 1};{_}{num + 2};{_}{num + 3};{_}{num + 4};{_}{num + 5};\
                            {_}{num + 6};{_}{num + 7};{_}{num + 8};{_}{num + 9};{_}{num + 10};{_}{num + 11})"),
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


def write_title_of_total(sheet):
    try:
        record_name = f"{displayed_results[0][1].replace(' ', '_', displayed_results[0][1].count(' '))}"
    except IndexError:
        record_name = ""
    style.alignment = xlwt.Alignment()

    style.font = _font_(bold=True)
    sheet.write_merge(r1=0, r2=0, c1=3, c2=4, label="Adb Version:", style=style)
    style.font = _font_(bold=False)
    sheet.write_merge(r1=0, r2=0, c1=5, c2=13, label=f"{xml_file.replace('.xml', '')}", style=style)

    style.font = _font_(bold=True)
    sheet.write_merge(r1=0, r2=0, c1=0, c2=1, label="Event:", style=style)
    style.font = _font_(bold=False)
    if var_checkbutton_1.get() == "0":
        sheet.write(0, 2, label="True", style=style)
    else:
        sheet.write(0, 2, label="False", style=style)

    style.font = _font_(bold=True)
    sheet.write_merge(r1=1, r2=1, c1=0, c2=1, label="Human:", style=style)
    style.font = _font_(bold=False)
    if var_checkbutton_2.get() == "0":
        sheet.write(1, 2, label="True", style=style)
    else:
        sheet.write(1, 2, label="False", style=style)

    style.font = _font_(bold=True)
    sheet.write_merge(r1=2, r2=2, c1=0, c2=1, label="Male:", style=style)
    style.font = _font_(bold=False)
    if var_checkbutton_3.get() == "0":
        sheet.write(2, 2, label="True", style=style)
    else:
        sheet.write(2, 2, label="False", style=style)

    style.font = _font_(bold=True)
    sheet.write_merge(r1=3, r2=3, c1=0, c2=1, label="Female:", style=style)
    style.font = _font_(bold=False)
    if var_checkbutton_4.get() == "0":
        sheet.write(3, 2, label="True", style=style)
    else:
        sheet.write(3, 2, label="False", style=style)

    style.font = _font_(bold=True)
    sheet.write_merge(r1=4, r2=4, c1=0, c2=1, label="North Hemisphere:", style=style)
    style.font = _font_(bold=False)
    if var_checkbutton_5.get() == "0":
        sheet.write(4, 2, label="True", style=style)
    else:
        sheet.write(4, 2, label="False", style=style)

    style.font = _font_(bold=True)
    sheet.write_merge(r1=5, r2=5, c1=0, c2=1, label="South Hemisphere:", style=style)
    style.font = _font_(bold=False)
    if var_checkbutton_6.get() == "0":
        sheet.write(5, 2, label="True", style=style)
    else:
        sheet.write(5, 2, label="False", style=style)

    style.font = _font_(bold=True)
    if selection == "observed":
        sheet.write_merge(r1=1, r2=1, c1=3, c2=4, label="House System:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=1, r2=1, c1=5, c2=13, label=f"{house_systems[hsys]}", style=style)
        style.font = _font_(bold=True)
        if len(displayed_results) == 1:
            sheet.write_merge(r1=2, r2=2, c1=3, c2=4, label="Rodden Rating:", style=style)
            style.font = _font_(bold=False)
            sheet.write_merge(r1=2, r2=2, c1=5, c2=13, label=f"{displayed_results[0][3]}", style=style)
        else:
            sheet.write_merge(r1=2, r2=2, c1=3, c2=4, label="Rodden Rating:", style=style)
            style.font = _font_(bold=False)
            sheet.write_merge(r1=2, r2=2, c1=5, c2=13, label=f"{'+'.join(selected_ratings)}", style=style)
        style.font = _font_(bold=True)
        if len(selected_categories) == 1:
            if len(displayed_results) == 1:
                sheet.write_merge(r1=3, r2=3, c1=3, c2=4, label="Category:", style=style)
                style.font = _font_(bold=False)
                sheet.write_merge(r1=3, r2=3, c1=5, c2=13, label=f"{record_name}", style=style)
            else:
                sheet.write_merge(r1=3, r2=3, c1=3, c2=4, label="Category:", style=style)
                style.font = _font_(bold=False)
                sheet.write_merge(r1=3, r2=3, c1=5, c2=13,
                                  label=f"{'/'.join(modify_category_names())}", style=style)
        else:
            sheet.write_merge(r1=3, r2=3, c1=3, c2=4, label="Category:", style=style)
            style.font = _font_(bold=False)
            sheet.write_merge(r1=3, r2=3, c1=5, c2=13, label="Control_Group", style=style)
    elif selection == "expected":
        sheet.write_merge(r1=1, r2=1, c1=3, c2=4, label="House System:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=1, r2=1, c1=5, c2=13, label=f"{r1[:-3]}", style=style)
        style.font = _font_(bold=True)
        sheet.write_merge(r1=2, r2=2, c1=3, c2=4, label="Rodden Rating:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=2, r2=2, c1=5, c2=13, label=f"{r2[:-3]}", style=style)
        style.font = _font_(bold=True)
        sheet.write_merge(r1=3, r2=3, c1=3, c2=4, label="Category:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=3, r2=3, c1=5, c2=13, label=f"{r3[:-3]}", style=style)
    elif selection == "chisquare" or selection == "effectsize":
        sheet.write_merge(r1=1, r2=1, c1=3, c2=4, label="House System:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=1, r2=1, c1=5, c2=13, label=f"{r1}", style=style)
        style.font = _font_(bold=True)
        sheet.write_merge(r1=2, r2=2, c1=3, c2=4, label="Rodden Rating:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=2, r2=2, c1=5, c2=13, label=f"{r2}", style=style)
        style.font = _font_(bold=True)
        sheet.write_merge(r1=3, r2=3, c1=3, c2=4, label="Category:", style=style)
        style.font = _font_(bold=False)
        sheet.write_merge(r1=3, r2=3, c1=5, c2=13, label=f"{r3}", style=style)

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
            sheet.write_merge(r1=230 + (i * 15), r2=230 + (i * 15), c1=0, c2=1, label="Total", style=style)
    style.font = _font_(bold=False)


def write_total_data(sheet):
    write_total_horz(sheet=sheet, num=7)
    write_total_horz(sheet=sheet, num=21)
    write_total_horz(sheet=sheet, num=35)
    for i in range(14):
        if i == 12:
            write_total_horz(sheet=sheet, num=218 + (i * 15))
            write_total_vert(sheet=sheet, num=220 + (i * 15))
        elif i == 13:
            write_total_horz(sheet=sheet, num=219 + (i * 15))
            write_total_vert(sheet=sheet, num=221 + (i * 15))
        else:
            write_total_horz(sheet=sheet, num=217 + (i * 15), col=14)
            write_total_vert(sheet=sheet, num=219 + (i * 15), col=2)


def save_file(file, sheet, table_name, table0="table0", table1="table1"):
    write_total_data(sheet=sheet)
    os.remove(f"{table0}.xlsx")
    os.remove(f"{table1}.xlsx")
    file.save(f"{table_name}.xlsx")


def special_cells(new_sheet, i):
    style.font = _font_(bold=True)
    if "Orb" in i[1]:
        style.alignment = xlwt.Alignment()
        new_sheet.write_merge(r1=i[0][0], r2=i[0][0], c1=i[0][1], c2=i[0][1] + 2,
                              label=i[1], style=style)
        style.alignment = alignment
    elif i[1] == "Traditional House Rulership":
        new_sheet.write_merge(r1=397, r2=397, c1=0, c2=13,
                              label="Traditional House Rulership", style=style)
    elif i[1] == "Modern House Rulership":
        new_sheet.write_merge(r1=413, r2=413, c1=0, c2=13,
                              label="Modern House Rulership", style=style)
    elif "in House" in i[1]:
        new_sheet.write_merge(r1=i[0][0], r2=i[0][0], c1=i[0][1], c2=i[0][1] + 1,
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
    save_file(new_file, new_sheet, "table0")


_planets_ = {i: [] for i in planets}
_planets = {i: [] for i in planets}


def extract_aspects(planet, value, num=11):
    if len(_planets_[f"{planet}"]) == num:
        _planets[f"{planet}"].append(_planets_[f"{planet}"])
        _planets_[f"{planet}"] = []
        _planets_[f"{planet}"].append(value)
    elif len(_planets_[f"{planet}"]) < num:
        _planets_[f"{planet}"].append(value)


def search_aspect(planet_pos, sheet, row: int, aspect: int, orb: int, name: str):
    _row = row
    n_ = 1
    name_frmt = f"{name}: Orb Factor: +- {orb}"
    style.alignment = xlwt.Alignment()
    sheet.write_merge(r1=_row - 2, r2=_row - 2, c1=0, c2=2, label=name_frmt, style=style)
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
                if j[0] == "Sun":
                    extract_aspects(planet="Sun", value=1, num=11)
                elif j[0] == "Moon":
                    extract_aspects(planet="Moon", value=1, num=10)
                elif j[0] == "Mercury":
                    extract_aspects(planet="Mercury", value=1, num=9)
                elif j[0] == "Venus":
                    extract_aspects(planet="Venus", value=1, num=8)
                elif j[0] == "Mars":
                    extract_aspects(planet="Mars", value=1, num=7)
                elif j[0] == "Jupiter":
                    extract_aspects(planet="Jupiter", value=1, num=6)
                elif j[0] == "Saturn":
                    extract_aspects(planet="Saturn", value=1, num=5)
                elif j[0] == "Uranus":
                    extract_aspects(planet="Uranus", value=1, num=4)
                elif j[0] == "Neptune":
                    extract_aspects(planet="Neptune", value=1, num=3)
                elif j[0] == "Pluto":
                    extract_aspects( planet="Pluto", value=1, num=2)
                elif j[0] == "North Node":
                    extract_aspects(planet="North Node", value=1, num=1)
                elif j[0] == "Chiron":
                    extract_aspects(planet="Chiron", value=1, num=0)
            else: 
                sheet.write(k + _row, i, 0, style=style)
                if j[0] == "Sun":
                    extract_aspects(planet="Sun", value=0, num=11)
                elif j[0] == "Moon":
                    extract_aspects(planet="Moon", value=0, num=10)
                elif j[0] == "Mercury":
                    extract_aspects(planet="Mercury", value=0, num=9)
                elif j[0] == "Venus":
                    extract_aspects(planet="Venus", value=0, num=8)
                elif j[0] == "Mars":
                    extract_aspects(planet="Mars", value=0, num=7)
                elif j[0] == "Jupiter":
                    extract_aspects(planet="Jupiter", value=0, num=6)
                elif j[0] == "Saturn":
                    extract_aspects(planet="Saturn", value=0, num=5)
                elif j[0] == "Uranus":
                    extract_aspects(planet="Uranus", value=0, num=4)
                elif j[0] == "Neptune":
                    extract_aspects(planet="Neptune", value=0, num=3)
                elif j[0] == "Pluto":
                    extract_aspects(planet="Pluto", value=0, num=2)
                elif j[0] == "North Node":
                    extract_aspects(planet="North Node", value=0, num=1)
                elif j[0] == "Chiron":
                    extract_aspects(planet="Chiron", value=0, num=0)
        n_ += 1
        _row += 1
    style.font = _font_(bold=True)


def sum_aspects(planet):
    for i, j in enumerate(planet):
        planet[i] = np.array(j)
    new_list = list(planet[0] + planet[1] + planet[2] + planet[3] + planet[4] + planet[5] + planet[6] +
                    planet[7] + planet[8] + planet[9] + planet[10])
    return [float(i) for i in new_list]


def write_datas_to_excel(get_datas):
    global _planets_, _planets
    _planets_ = {i: [] for i in planets}
    _planets = {i: [] for i in planets}
    file = xlwt.Workbook()
    sheet = file.add_sheet("Sheet1")
    planet_info, house_info, planet_pos, traditional_ruler, modern_ruler = get_datas
    style.font = _font_(bold=True)
    for i, j in enumerate(signs):
        sheet.write(7, i + 1, j, style=style)
        sheet.write(21, i + 1, j, style=style)
    for i, j in enumerate(planet_info):
        sheet.write(i + 8, 0, j[0], style=style)
        sheet.write(i + 36, 0, j[0], style=style)
    for i in range(12):
        sheet.write(i + 22, 0, f"House {i + 1}", style=style)
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
                if j[0] == planets[0]:
                    __planets__[0].append(j)
                elif j[0] == planets[1]:
                    __planets__[1].append(j)
                elif j[0] == planets[2]:
                    __planets__[2].append(j)
                elif j[0] == planets[3]:
                    __planets__[3].append(j)
                elif j[0] == planets[4]:
                    __planets__[4].append(j)
                elif j[0] == planets[5]:
                    __planets__[5].append(j)
                elif j[0] == planets[6]:
                    __planets__[6].append(j)
                elif j[0] == planets[7]:
                    __planets__[7].append(j)
                elif j[0] == planets[8]:
                    __planets__[8].append(j)
                elif j[0] == planets[9]:
                    __planets__[9].append(j)
                elif j[0] == planets[10]:
                    __planets__[10].append(j)
                elif j[0] == planets[11]:
                    __planets__[11].append(j)
                sheet.write(i + 36, k + 1, 1, style=style)
            else:
                sheet.write(i + 36, k + 1, 0, style=style)
    style.font = _font_(bold=True)
    search_aspect(planet_pos, sheet, row=51, aspect=0, orb=conjunction, name="Conjunction")
    search_aspect(planet_pos, sheet, row=65, aspect=30, orb=semi_sextile, name="Semi-Sextile")
    search_aspect(planet_pos, sheet, row=79, aspect=45, orb=semi_square, name="Semi-Square")
    search_aspect(planet_pos, sheet, row=93, aspect=60, orb=sextile, name="Sextile")
    search_aspect(planet_pos, sheet, row=107, aspect=72, orb=quintile, name="Quintile")
    search_aspect(planet_pos, sheet, row=121, aspect=90, orb=square, name="Square")
    search_aspect(planet_pos, sheet, row=135, aspect=120, orb=trine, name="Trine")
    search_aspect(planet_pos, sheet, row=149, aspect=135, orb=sesquiquadrate, name="Sesquiquadrate")
    search_aspect(planet_pos, sheet, row=163, aspect=144, orb=biquintile, name="BiQuintile")
    search_aspect(planet_pos, sheet, row=177, aspect=150, orb=quincunx, name="Quincunx")
    search_aspect(planet_pos, sheet, row=191, aspect=180, orb=opposite, name="Opposite")
    style.alignment = xlwt.Alignment()
    sheet.write_merge(r1=203, r2=203, c1=0, c2=2, label="All Aspects", style=style)
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
            if m[2] == "1":
                house_group[0].append(m[1])
            elif m[2] == "2":
                house_group[1].append(m[1])
            elif m[2] == "3":
                house_group[2].append(m[1])
            elif m[2] == "4":
                house_group[3].append(m[1])
            elif m[2] == "5":
                house_group[4].append(m[1])
            elif m[2] == "6":
                house_group[5].append(m[1])
            elif m[2] == "7":
                house_group[6].append(m[1])
            elif m[2] == "8":
                house_group[7].append(m[1])
            elif m[2] == "9":
                house_group[8].append(m[1])
            elif m[2] == "10":
                house_group[9].append(m[1])
            elif m[2] == "11":
                house_group[10].append(m[1])
            elif m[2] == "12":
                house_group[11].append(m[1])
        new_order_of_planets.append(house_group)
    count = 0
    style.font = _font_(bold=True)
    for _ in new_order_of_planets:
        for o, p in enumerate(signs):
            sheet.write(217 + count, o + 2, p, style=style)
        count += 15
    count = 0
    for i in planets:
        for j, k in enumerate(houses):
            form = f"{i} in House {k}"
            sheet.write_merge(r1=217 + count + (j + 1), r2=217 + count + (j + 1), c1=0, c2=1,
                              label=form, style=style)
        count += 15
    count = 0
    style.font = _font_(bold=False)
    for i in new_order_of_planets:
        for j, k in enumerate(i):
            if "Aries" in k:
                sheet.write(217 + count + (j + 1), 2, 1, style=style)
            elif "Taurus" in k:
                sheet.write(217 + count + (j + 1), 3, 1, style=style)
            elif "Gemini" in k:
                sheet.write(217 + count + (j + 1), 4, 1, style=style)
            elif "Cancer" in k:
                sheet.write(217 + count + (j + 1), 5, 1, style=style)
            elif "Leo" in k:
                sheet.write(217 + count + (j + 1), 6, 1, style=style)
            elif "Virgo" in k:
                sheet.write(217 + count + (j + 1), 7, 1, style=style)
            elif "Libra" in k:
                sheet.write(217 + count + (j + 1), 8, 1, style=style)
            elif "Scorpio" in k:
                sheet.write(217 + count + (j + 1), 9, 1, style=style)
            elif "Sagittarius" in k:
                sheet.write(217 + count + (j + 1), 10, 1, style=style)
            elif "Capricorn" in k:
                sheet.write(217 + count + (j + 1), 11, 1, style=style)
            elif "Aquarius" in k:
                sheet.write(217 + count + (j + 1), 12, 1, style=style)
            elif "Pisces" in k:
                sheet.write(217 + count + (j + 1), 13, 1, style=style)
            else:
                for m in range(12):
                    sheet.write(217 + count + (j + 1), m + 2, 0, style=style)
        count += 15
    count = 0
    style.font = _font_(bold=True)
    sheet.write_merge(r1=397, r2=397, c1=0, c2=13, label="Traditional House Rulership", style=style)
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
    sheet.write_merge(r1=413, r2=413, c1=0, c2=13, label="Modern House Rulership", style=style)
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
        return change_dir(cat, dir1, orb_factor, sub_dir=f"{criteria}")
    elif var_checkbutton_5.get() == "0" and var_checkbutton_6.get() == "1":
        return change_dir(cat, dir1, orb_factor, sub_dir=os.path.join(f"{criteria}", "North"))
    elif var_checkbutton_5.get() == "1" and var_checkbutton_6.get() == "0":
        return change_dir(cat, dir1, orb_factor, sub_dir=os.path.join(f"{criteria}", "South"))


def dir_names(cat, dir1, orb_factor):
    if len(displayed_results) == 1:
        return os.path.join(*cat, dir1, f"ORB_{'_'.join(orb_factor)}", f"{house_systems[hsys]}")
    if var_checkbutton_1.get() == "0" and var_checkbutton_2.get() == "1":
        return check_dir_names(cat, dir1, orb_factor, criteria="Event")
    elif var_checkbutton_1.get() == "1" and var_checkbutton_2.get() == "0":
        if var_checkbutton_3.get() == "0" and var_checkbutton_4.get() == "1":
            return check_dir_names(cat, dir1, orb_factor, criteria="Male")
        elif var_checkbutton_3.get() == "1" and var_checkbutton_4.get() == "0":
            return check_dir_names(cat, dir1, orb_factor, criteria="Female")
        elif var_checkbutton_3.get() == "0" and var_checkbutton_4.get() == "0":
            return check_dir_names(cat, dir1, orb_factor, criteria="Human")
    elif var_checkbutton_1.get() == "0" and var_checkbutton_2.get() == "0":
        if var_checkbutton_3.get() == "0" and var_checkbutton_4.get() == "1":
            return check_dir_names(cat, dir1, orb_factor, criteria="Male+Event")
        elif var_checkbutton_3.get() == "1" and var_checkbutton_4.get() == "0":
            return check_dir_names(cat, dir1, orb_factor, criteria="Female+Event")
        elif var_checkbutton_3.get() == "0" and var_checkbutton_4.get() == "0":
            return check_dir_names(cat, dir1, orb_factor, criteria="Human+Event")
        elif var_checkbutton_3.get() == "1" and var_checkbutton_4.get() == "1":
            return check_dir_names(cat, dir1, orb_factor, criteria="Event")


def find_observed_values():
    global selection, dir2
    selection = "observed"
    if len(displayed_results) == 0:
        msgbox.showinfo(title="Find Observed Values", message=f"{len(displayed_results)} records selected.")
    else:
        __size__ = len(displayed_results)
        __received__ = 0
        __now__ = time.time()
        pframe = tk.Frame(master=master)
        pbar = Progressbar(master=pframe, orient="horizontal", length=200, mode="determinate")
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
        with open("output.log", "w", encoding="utf-8") as log:
            log.write(f"Adb Version: {xml_file.replace('.xml', '')}\n")
            log.write(f"House System: {house_systems[hsys]}\n")
            if len(displayed_results) == 1:
                log.write(f"Rodden Rating: {displayed_results[0][3]}\n")
            else:
                log.write(f"Rodden Rating: {'+'.join(selected_ratings)}\n")
            log.write(f"Orb Factor: {'_'.join(orb_factor)}\n")
            if len(selected_categories) == 1:
                if len(displayed_results) == 1:
                    log.write(f"Category: \
{displayed_results[0][1].replace(' ','_', displayed_results[0][1].count(' '))}\n\n")
                else:
                    log.write(f"Category: {'/'.join(modify_category_names())}\n\n")
            elif len(selected_categories) > 1:
                log.write(f"Category: Control_Group\n\n")
                if var_checkbutton_1.get() == "0":
                    log.write(f"Event: True\n")
                else:
                    log.write(f"Event: False\n")
                if var_checkbutton_2.get() == "0":
                    log.write(f"Human: True\n")
                else:
                    log.write(f"Human: False\n")
                if var_checkbutton_3.get() == "0":
                    log.write(f"Male: True\n")
                else:
                    log.write(f"Male: False\n")
                if var_checkbutton_4.get() == "0":
                    log.write(f"Female: True\n")
                else:
                    log.write(f"Female: False\n")
                if var_checkbutton_5.get() == "0":
                    log.write(f"North Hemisphere: True\n")
                else:
                    log.write(f"North Hemisphere: False\n")
                if var_checkbutton_6.get() == "0":
                    log.write(f"South Hemisphere: True\n\n")
                else:
                    log.write(f"South Hemisphere: False\n\n")
            log.write(f"|{str(datetime.now())[:-7]}| Process started.\n\n")
            log.flush()
            for records in displayed_results:
                julian_date = float(records[6])
                longitude = records[8]
                if "e" in longitude:
                    longitude = float(longitude.replace("e", "."))
                elif "w" in longitude:
                    longitude = -1 * float(longitude.replace("w", "."))
                latitude = records[7]
                if "n" in latitude:
                    latitude = float(latitude.replace("n", "."))
                elif "s" in latitude:
                    latitude = -1 * float(latitude.replace("s", "."))
                try:
                    chart = Chart(julian_date, longitude, latitude)
                    write_datas_to_excel(chart.get_chart_data())
                except BaseException as err:
                    traceback.print_exc(file=sys.stdout)
                    log.write(f"|{str(datetime.now())[:-7]}| Error Type: {err}\n{' ' * 22}Record: {records}\n\n")
                    log.flush()
                __received__ += 1
                if __received__ != __size__:
                    pbar["value"] = __received__
                    pbar["maximum"] = __size__
                    pstring.set("{} %, {} minutes remaining.".format(
                        int(100 * __received__ / __size__),
                        round((int(__size__ / (__received__ / (time.time() - __now__))) - int(
                            time.time() - __now__)) / 60)))
                else:
                    pframe.destroy()
                    pbar.destroy()
                    plabel.destroy()
                    try:
                        os.rename("table0.xlsx", "observed_values.xlsx")
                        dir1 = f"RR_{'+'.join(selected_ratings)}"
                        if len(selected_categories) == 1:
                            if len(displayed_results) == 1:
                                cat = [displayed_results[0][1].replace(" ", "_", displayed_results[0][1].count(" "))]
                                dir1 = f"RR_{displayed_results[0][3]}"
                            else:
                                cat = modify_category_names()
                            dir2 = dir_names(cat, dir1, orb_factor)
                        elif len(selected_categories) > 1:
                            cat = ["Control_Group"]
                            dir2 = dir_names(cat, dir1, orb_factor)    
                        try:
                            os.makedirs(dir2)
                            shutil.move(src=os.path.join(os.getcwd(), "observed_values.xlsx"),
                                        dst=os.path.join(os.getcwd(), dir2, "observed_values.xlsx"))
                        except FileExistsError as err:
                            log.write(f"|{str(datetime.now())[:-7]}| Error Type: {err}\n\n")
                            log.flush()
                    except FileNotFoundError:
                        pass
                    master.update()
                    log.write(f"|{str(datetime.now())[:-7]}| Process finished.")
                    shutil.move(src=os.path.join(os.getcwd(), "output.log"),
                                dst=os.path.join(os.getcwd(), dir2, "output.log"))
                    msgbox.showinfo(title="Find Observed Values", message="Process finished successfully.")


def sum_of_row(table):
    total = 0
    for item in table:
        if item[0][0] == 8 and type(item[1]) == float:
            total += item[1]
    return total


def find_expected_values():
    global selection
    selection = "expected"
    file_name_1 = "control_group.xlsx"
    file_name_2 = "observed_values.xlsx"
    read_file_1 = None
    read_file_2 = None
    try:
        read_file_1 = xlrd.open_workbook(file_name_1)
    except FileNotFoundError:
        msgbox.showinfo(title="Find Expected Values", message="No such file or directory: 'control_group.xlsx'")
    try:
        read_file_2 = xlrd.open_workbook(file_name_2)
    except FileNotFoundError:
        msgbox.showinfo(title="Find Expected Values", message="No such file or directory: 'observed_values.xlsx'")
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
                            if method is False:
                                # Sjoerd Visser's method
                                new_sheet.write(*i[0], i[1] * ratio, style=style)
                            elif method is True:
                                # Flavia Minghetti's method
                                new_sheet.write(
                                    *i[0],
                                    sum_of_row(data_2) * (i[1] + j[1]) / sum_of_all,
                                    style=style
                                )
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
        save_file(new_file, new_sheet, "expected_values", "control_group", "observed_values")
        master.update()
        msgbox.showinfo(title="Find Expected Values", message="Process finished successfully.")


def find_chi_square_values():
    global selection
    selection = "chisquare"
    file_name_1 = "expected_values.xlsx"
    file_name_2 = "observed_values.xlsx"
    read_file_1 = None
    read_file_2 = None
    try:
        read_file_1 = xlrd.open_workbook(file_name_1)
    except FileNotFoundError:
        msgbox.showinfo(title="Find Chi-Square Values", message="No such file or directory: 'expected_values.xlsx'")
    try:
        read_file_2 = xlrd.open_workbook(file_name_2)
    except FileNotFoundError:
        msgbox.showinfo(title="Find Chi-Square Values", message="No such file or directory: 'observed_values.xlsx'")
    if read_file_1 is not None and read_file_2 is not None:
        sheet_1 = read_file_1.sheet_by_name("Sheet1")
        sheet_2 = read_file_2.sheet_by_name("Sheet1")
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
                            try:
                                new_sheet.write(*i[0], (i[1] - j[1]) ** 2 / i[1], style=style)
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
        save_file(new_file, new_sheet, "chi-square", "expected_values", "observed_values")
        master.update()
        msgbox.showinfo(title="Find Chi-Square Values", message="Process finished successfully.")


def find_effect_size_values():
    global selection
    selection = "effectsize"
    file_name_1 = "observed_values.xlsx"
    file_name_2 = "expected_values.xlsx"
    read_file_1 = None
    read_file_2 = None
    try:
        read_file_1 = xlrd.open_workbook(file_name_1)
    except FileNotFoundError:
        msgbox.showinfo(title="Find Effect Size Values", message="No such file or directory: 'observed_values.xlsx'")
    try:
        read_file_2 = xlrd.open_workbook(file_name_2)
    except FileNotFoundError:
        msgbox.showinfo(title="Find Effect Size Values", message="No such file or directory: 'expected_values.xlsx'")
    if read_file_1 is not None and read_file_2 is not None:
        sheet_1 = read_file_1.sheet_by_name("Sheet1")
        sheet_2 = read_file_2.sheet_by_name("Sheet1")
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
                            try:
                                new_sheet.write(*i[0], i[1] / j[1], style=style)
                            except ZeroDivisionError:
                                new_sheet.write(*i[0], 0, style=style)
                        else:
                            style.font = _font_(bold=True)
                            special_cells(new_sheet, i)
                            style.font = _font_(bold=False)
                    elif i[1] != "" and j[1] == "":
                        new_sheet.write(*i[0], i[1], style=style)
                    elif i[1] == "" and j[1] != "":
                        new_sheet.write(*j[0], j[1], style=style)
                    if i[0] not in control_list:
                        control_list.append(i[0])
                    if j[0] not in control_list:
                        control_list.append(j[0])
        save_file(new_file, new_sheet, "effect-size", "observed_values", "expected_values")
        master.update()
        msgbox.showinfo(title="Find Effect Size Values", message="Process finished successfully.")


# --------------------------------------------- main ---------------------------------------------


def main():
    freq_frmt = [0, 2000, 100]

    def func1():
        t1 = threading.Thread(target=find_observed_values)
        t1.start()

    def func2():
        t2 = threading.Thread(target=find_expected_values)
        t2.start()

    def func3():
        t3 = threading.Thread(target=find_chi_square_values)
        t3.start()

    def func4():
        t4 = threading.Thread(target=find_effect_size_values)
        t4.start()

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
        global quintile, square, trine, sesquiquadrate, biquintile, quincunx, opposite
        conjunction = int(orb_entries[0].get())
        semi_sextile = int(orb_entries[1].get())
        semi_square = int(orb_entries[2].get())
        sextile = int(orb_entries[3].get())
        quintile = int(orb_entries[4].get())
        square = int(orb_entries[5].get())
        trine = int(orb_entries[6].get())
        sesquiquadrate = int(orb_entries[7].get())
        biquintile = int(orb_entries[7].get())
        quincunx = int(orb_entries[9].get())
        opposite = int(orb_entries[10].get())
        parent.destroy()

    def choose_orb_factor():
        toplevel3 = tk.Toplevel()
        toplevel3.title("Orb Factor")
        toplevel3.resizable(width=False, height=False)
        aspects = ["Conjunction", "Semi-Sextile", "Semi-Square", "Sextile", "Quintile",
                   "Square", "Trine", "Sesquiquadrate", "BiQuintile", "Quincunx", "Opposite"]
        default_orbs = [conjunction, semi_sextile, semi_square, sextile, quintile,
                        square, trine, sesquiquadrate, biquintile, quincunx, opposite]
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
        apply_button = tk.Button(master=toplevel3, text="Apply",
                                 command=lambda: change_orb_factors(parent=toplevel3, orb_entries=orb_entries))
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
        hsys_frame = tk.Frame(master=toplevel4)
        hsys_frame.pack(side="top")
        button_frame = tk.Frame(master=toplevel4)
        button_frame.pack(side="bottom")
        _house_systems_ = [values for keys, values in house_systems.items()]
        checkbuttons = dict()
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
            command=lambda: check_uncheck(checkbuttons, _house_systems_, "Placidus"))
        checkbuttons["Koch"][0].configure(
            command=lambda: check_uncheck(checkbuttons, _house_systems_, "Koch"))
        checkbuttons["Porphyrius"][0].configure(
            command=lambda: check_uncheck(checkbuttons, _house_systems_, "Porphyrius"))
        checkbuttons["Regiomontanus"][0].configure(
            command=lambda: check_uncheck(checkbuttons, _house_systems_, "Regiomontanus"))
        checkbuttons["Campanus"][0].configure(
            command=lambda: check_uncheck(checkbuttons, _house_systems_, "Campanus"))
        checkbuttons["Equal"][0].configure(
            command=lambda: check_uncheck(checkbuttons, _house_systems_, "Equal"))
        checkbuttons["Whole Signs"][0].configure(
            command=lambda: check_uncheck(checkbuttons, _house_systems_, "Whole Signs"))
        apply_button = tk.Button(master=button_frame, text="Apply",
                                 command=lambda: change_hsys(
                                     parent=toplevel4,
                                     checkbuttons=checkbuttons,
                                     _house_systems_=_house_systems_))
        apply_button.pack()

    def export_link():
        with open("links.txt", "w", encoding="utf-8") as f:
            for i, j in enumerate(displayed_results):
                f.write(f"{i + 1}. {j[12]}\n")
        msgbox.showinfo(title="Export Links", message=f"{len(displayed_results)} links were exported.")

    def year_frequency_command(parent, date_entries, years):
        min_, max_, step_ = date_entries[:]
        min_, max_, step_ = int(min_.get()), int(max_.get()), int(step_.get())
        freq_frmt[0], freq_frmt[1], freq_frmt[2] = min_, max_, step_
        with open("year-frequency.txt", "w", encoding="utf-8") as f:
            year_dict = dict()
            count = 0
            for i in range(min_, max_, step_):
                year_dict[(min_ + (count * step_), min_ + (count * step_) + step_)] = []
                count += 1
            for i in years:
                for keys, values in year_dict.items():
                    if keys[0] < i < keys[1]:
                        year_dict[keys[0], keys[1]] += i,
            for keys, values in year_dict.items():
                f.write(f"{keys} = {len(values)}\n")
            f.write(f"Total = {len(displayed_results)}")
            parent.destroy()
            msgbox.showinfo(title="Export Year Frequency",
                            message=f"{len(displayed_results)} dates were exported.")

    def export_year_frequency():
        toplevel5 = tk.Toplevel()
        toplevel5.title("Year Frequency")
        toplevel5.geometry("200x100")
        toplevel5.resizable(width=False, height=False)
        t5frame = tk.Frame(master=toplevel5)
        t5frame.pack()
        date_entries = []
        years = [int(i[4].split(" ")[2]) for i in displayed_results]
        if len(years) != 0:
            freq_frmt[0], freq_frmt[1], freq_frmt[2] = min(years), max(years), 100
        for i, j in enumerate(("Minimum", "Maximum", "Step")):
            date_label = tk.Label(master=t5frame, text=j)
            date_label.grid(row=i, column=0, sticky="w")
            date_entry = tk.Entry(master=t5frame, width=5)
            date_entry.grid(row=i, column=1, sticky="w")
            date_entry.insert("1", f"{freq_frmt[i]}")
            date_entries.append(date_entry)
        apply_button = tk.Button(master=t5frame, text="Apply",
                                 command=lambda: year_frequency_command(
                                     parent=toplevel5, date_entries=date_entries, years=years))
        apply_button.grid(row=3, column=0, columnspan=3)

    def callback(event, url):
        webbrowser.open_new(url)

    def about():
        toplevel5 = tk.Toplevel()
        toplevel5.title("About TkAstroDb")
        name = "TkAstroDb"
        version, _version = "Version:", __version__
        build_date, _build_date = "Built Date:", "21 December 2018"
        update_date, _update_date = "Update Date:", "14 January 2019"
        developed_by, _developed_by = "Developed By:", "Tanberk Celalettin Kutlu"
        thanks_to, _thanks_to = "Special Thanks To:", "Alois Treindl, Flavia Minghetti, Sjoerd Visser"
        contact, _contact = "Contact:", "tckutlu@gmail.com"
        github, _github = "GitHub:", "https://github.com/dildeolupbiten/TkAstroDb"
        tframe1 = tk.Frame(master=toplevel5, bd="2", relief="groove")
        tframe1.pack(fill="both")
        tframe2 = tk.Frame(master=toplevel5)
        tframe2.pack(fill="both")
        tlabel_title = tk.Label(master=tframe1, text=name, font="Arial 25")
        tlabel_title.pack()
        for i, j in enumerate((version, build_date, update_date, developed_by, thanks_to, contact, github)):
            tlabel_info_1 = tk.Label(master=tframe2, text=j, font="Arial 12", fg="red")
            tlabel_info_1.grid(row=i, column=0, sticky="w")
        for i, j in enumerate((_version, _build_date, _update_date, _developed_by, _thanks_to, _contact, _github)):
            if j == _github:
                tlabel_info_2 = tk.Label(master=tframe2, text=j, font="Arial 12", fg="blue", cursor="hand2")
                url1 = "https://github.com/dildeolupbiten/TkAstroDb"
                tlabel_info_2.bind("<Button-1>", lambda event: callback(event, url1))
            elif j == _contact:
                tlabel_info_2 = tk.Label(master=tframe2, text=j, font="Arial 12", fg="blue", cursor="hand2")
                url2 = "mailto://tckutlu@gmail.com"
                tlabel_info_2.bind("<Button-1>", lambda event: callback(event, url2))
            else:
                tlabel_info_2 = tk.Label(master=tframe2, text=j, font="Arial 12")
            tlabel_info_2.grid(row=i, column=1, sticky="w")

    def update():
        url_1 = "https://raw.githubusercontent.com/dildeolupbiten/TkAstroDb/master/TkAstroDb.py"
        url_2 = "https://raw.githubusercontent.com/dildeolupbiten/TkAstroDb/master/README.md"
        data_1 = urllib.urlopen(url=url_1, context=ssl.SSLContext(ssl.PROTOCOL_SSLv23))
        data_2 = urllib.urlopen(url=url_2, context=ssl.SSLContext(ssl.PROTOCOL_SSLv23))
        with open("TkAstroDb.py", "r", encoding="utf-8") as f:
            var_1 = [i.decode("utf-8") for i in data_1]
            var_2 = [i.decode("utf-8") for i in data_2]
            var_3 = [i for i in f]
            if var_1 == var_3:
                msgbox.showinfo(title="Update", message="Program is up-to-date.")
            else:
                with open("README.md", "w", encoding="utf-8") as g:
                    for i in var_2:
                        g.write(i)
                        g.flush()
                with open("TkAstroDb.py", "w", encoding="utf-8") as h:
                    for i in var_1:
                        h.write(i)
                        h.flush()
                    msgbox.showinfo(title="Update", message="Program is updated.")
                    if os.name == "posix":
                        import signal      
                        os.kill(os.getpid(), signal.SIGKILL)
                    elif os.name == "nt":
                        os.system(f"TASKKILL /F /PID {os.getpid()}")
              
    menubar = tk.Menu(master=master)
    master.configure(menu=menubar)

    calculations_menu = tk.Menu(master=menubar, tearoff=False)
    export_menu = tk.Menu(master=menubar, tearoff=False)
    options_menu = tk.Menu(master=menubar, tearoff=False)
    help_menu = tk.Menu(master=menubar, tearoff=False)

    menubar.add_cascade(label="Calculations", menu=calculations_menu)
    menubar.add_cascade(label="Export", menu=export_menu)
    menubar.add_cascade(label="Options", menu=options_menu)
    menubar.add_cascade(label="Help", menu=help_menu)

    method_menu = tk.Menu(master=calculations_menu, tearoff=False)

    calculations_menu.add_command(label="Find Observed Values", command=func1)
    calculations_menu.add_cascade(label="Find Expected Values", menu=method_menu)
    calculations_menu.add_command(label="Find Chi-Square Values", command=func3)
    calculations_menu.add_command(label="Find Effect Size Values", command=func4)

    method_menu.add_command(label="Flavia's method", command=set_method_to_true)
    method_menu.add_command(label="Sjoerd's method", command=set_method_to_false)

    export_menu.add_command(label="Adb Links", command=export_link)
    export_menu.add_command(label="Year Frequency", command=export_year_frequency)

    options_menu.add_command(label="House System", command=create_hsys_checkbuttons)
    options_menu.add_command(label="Orb Factor", command=choose_orb_factor)

    help_menu.add_command(label="About", command=about)
    help_menu.add_command(label="Check for Updates", command=update)

    t0 = threading.Thread(target=master.mainloop)
    t0.run()


if __name__ == "__main__":
    main()
