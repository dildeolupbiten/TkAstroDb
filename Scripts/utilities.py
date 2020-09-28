# -*- coding: utf-8 -*-

from .messagebox import MsgBox
from .constants import PLANETS, SIGNS
from .modules import (
    dt, os, json, time, Popen, ImageTk, urlopen, URLError, ConfigParser
)


def create_image_files(path):
    return {
        i[:-4]: {
            "path": os.path.join(os.getcwd(), path, i),
            "img": ImageTk.PhotoImage(
                file=os.path.join(os.getcwd(), path, i)
            )
        }
        for i in sorted(os.listdir(os.path.join(os.getcwd(), path)))
    }


def check_update(icons):
    update = False
    for d in ["Scripts", "Algorithms"]:
        try:
            scripts = json.load(
                urlopen(
                    url=f"https://api.github.com/repos/dildeolupbiten/"
                        f"TkAstroDb/contents/{d}?ref=master"
                )
            )
        except URLError:
            MsgBox(
                title="Info",
                message="Couldn't connect to server.",
                level="info",
                icons=icons
            )
            return
        for i in scripts:
            file = urlopen(i["download_url"]).read().decode()
            if i["name"] not in os.listdir(d):
                update = True
                with open(f"{d}/{i['name']}", "w", encoding="utf-8") as f:
                    f.write(file)
            else:
                with open(f"{d}/{i['name']}", "r", encoding="utf-8") as f:
                    if file != f.read():
                        update = True
                        with open(
                                f"{d}/{i['name']}",
                                "w",
                                encoding="utf-8"
                        ) as g:
                            g.write(file)
    if update:
        MsgBox(
            title="Info",
            message="Program is updated.",
            level="info",
            icons=icons
        )
        if os.name == "posix":
            Popen(["python3", "run.py"])
            import signal
            os.kill(os.getpid(), signal.SIGKILL)
        elif os.name == "nt":
            Popen(["python", "run.py"])
            os.system(f"TASKKILL /F /PID {os.getpid()}")
    else:
        MsgBox(
            title="Info",
            message="Program is up-to-date.",
            level="info",
            icons=icons
        )


def load_defaults():
    if os.path.exists("defaults.ini"):
        return
    config = ConfigParser()
    with open("defaults.ini", "w") as f:
        config["HOUSE SYSTEM"] = {"selected": "Placidus"}
        config["ORB FACTORS"] = {
            "Conjunction": 6,
            "Semi-Sextile": 2,
            "Semi-Square": 2,
            "Sextile": 4,
            "Quintile": 2,
            "Square": 6,
            "Trine": 6,
            "Sesquiquadrate": 2,
            "BiQuintile": 2,
            "Quincunx": 2,
            "Opposite": 6
        }
        config["PLANETS"] = {
            "selected":
                ", ".join(
                    planet
                    for planet in PLANETS
                )
        }
        config["DATABASE"] = {"selected": "None"}
        config["METHOD"] = {"selected": "Subcategory"}
        config.write(f)


def msgbox_info(self, message):
    self.logging_text["state"] = "normal"
    self.logging_text.insert(
        "insert",
        f"- INFO - {dt.now().strftime('%Y.%m.%d %H:%M:%S')} - {message}"
    )
    self.logging_text["state"] = "disabled"


def convert_coordinates(coord):
    if "n" in coord:
        d, _m = coord.split("n")
        if len(_m) == 4:
            m = _m[:2]
            s = _m[2:]
            return dms_to_dd(f"{d}\u00b0{m}'{s}\"")
        return dms_to_dd(coord.replace("n", "\u00b0") + "'0\"")
    elif "s" in coord:
        d, _m = coord.split("s")
        if len(_m) == 4:
            m = _m[:2]
            s = _m[2:]
            return -1 * dms_to_dd(f"{d}\u00b0{m}'{s}\"")
        return -1 * dms_to_dd(coord.replace("s", "\u00b0") + "'0\"")
    elif "e" in coord:
        d, _m = coord.split("e")
        if len(_m) == 4:
            m = _m[:2]
            s = _m[2:]
            return dms_to_dd(f"{d}\u00b0{m}'{s}\"")
        return dms_to_dd(coord.replace("e", "\u00b0") + "'0\"")
    elif "w" in coord:
        d, _m = coord.split("w")
        if len(_m) == 4:
            m = _m[:2]
            s = _m[2:]
            return -1 * dms_to_dd(f"{d}\u00b0{m}'{s}\"")
        return -1 * dms_to_dd(coord.replace("w", "\u00b0") + "'0\"")


def tbutton_command(cvar_list, tlevel, select):
    for item in cvar_list:
        if item[0].get() is True:
            select.append(item[1])
    tlevel.destroy()


def convert_degree(degree):
    for i in range(12):
        if i * 30 <= degree < (i + 1) * 30:
            return degree - (30 * i), [*SIGNS][i]


def reverse_convert_degree(degree, sign):
    return degree + 30 * [*SIGNS].index(sign)


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


def check_all_command(check_all, cvar_list, checkbutton_list):
    if check_all.get() is True:
        for var, c_button in zip(cvar_list, checkbutton_list):
            var[0].set(True)
            c_button.configure(variable=var[0])
    else:
        for var, c_button in zip(cvar_list, checkbutton_list):
            var[0].set(False)
            c_button.configure(variable=var[0])


def progressbar(s, r, n, pframe, pbar, plabel, pstring):
    if r != s:
        pbar["value"] = r
        pbar["maximum"] = s
        pstring.set(
            "{} %, {} minutes remaining.".format(
                int(100 * r / s),
                round(
                    (int(s / (r / (time.time() - n))) -
                     int(time.time() - n)) / 60
                )
            )
        )
    else:
        pframe.destroy()
        pbar.destroy()
        plabel.destroy()