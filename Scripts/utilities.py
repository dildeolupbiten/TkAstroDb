# -*- coding: utf-8 -*-

from .messagebox import MsgBox
from .constants import SIGNS, PLANETS, SHEETS, ASPECTS
from .modules import (
    os, np, ET, json, time, Popen, urlopen,
    Thread, URLError, PhotoImage, ConfigParser
)


def get_xml_file_list(path):
    files = {
        os.path.join(path, i): os.stat(os.path.join(path, i)).st_mtime
        for i in os.listdir(path) if i.endswith("xml")
    }
    return sorted(files, key=files.get, reverse=True)
    
    
def start_merging_databases(files, widget, icons):
    split = [os.path.split(i)[-1][:-4] for i in files][::-1]
    data = []
    names = []
    for file in files:
        for i in from_xml(file)[0]:
            if i[1] not in names:
                data.append(i)
                names.append(i[1])
    with open(
            os.path.join("./Database", "_&_".join(split) + ".json"), 
            "w", 
            encoding="utf-8"
    ) as f:
        json.dump(sorted(data), f, indent=4, ensure_ascii=False)
    if len(files) == 1:
        txt = "Database was"
    else:
        txt = "Databases were \nmerged and"
    if files:
        widget.after(
            0,
            lambda: MsgBox(
                title="Info",
                level="info",
                message=f"{txt} converted!",
                icons=icons
            )
        )
        
    
def merge_databases(multiple_selection, icons, widget):
    files = get_xml_file_list("./Database")
    selection = multiple_selection(
        title="Merge And Convert File(s)",
        catalogue=[os.path.split(i)[-1] for i in files],
        get=True
    )
    files = [os.path.join("./Database", i) for i in selection.result]
    if not files:
        return
    else:
        Thread(
            target=lambda: start_merging_databases(
                files=files,
                widget=widget,
                icons=icons
            ),
            daemon=True
        ).start()


def load_database(filename):
    with open(filename, encoding="utf-8") as f:
        return json.load(f)


def create_image_files(path):
    return {
        i[:-4]: {
            "img": PhotoImage(
                file=os.path.join(os.getcwd(), path, i)
            )
        }
        for i in sorted(os.listdir(os.path.join(os.getcwd(), path)))
    }


def table_selection(multiple_selection):
    config = ConfigParser()
    config.read("defaults.ini")
    multiple_selection(
        title="Table Selection",
        catalogue=config["TABLE SELECTION"].keys()
    )


def load_defaults():
    if os.path.exists("defaults.ini"):
        return
    config = ConfigParser()
    with open("defaults.ini", "w") as f:
        config["ZODIAC"] = {"selected": "Tropical"}
        config["AYANAMSHA"] = {"selected": "Hindu-Lahiri"}
        config["HOUSE SYSTEM"] = {"selected": "Placidus"}
        config["ORB FACTORS"] = {
            "Conjunction": 10,
            "Semi-Sextile": 3,
            "Semi-Square": 3,
            "Sextile": 6,
            "Quintile": 2,
            "Square": 10,
            "Trine": 10,
            "Sesquiquadrate": 3,
            "BiQuintile": 2,
            "Quincunx": 3,
            "Opposite": 10
        }
        config["DATABASE"] = {"selected": "None"}
        config["METHOD"] = {"selected": "Subcategory"}
        config["CATEGORY SELECTION"] = {"selected": "Basic"}
        config["TABLE SELECTION"] = {
            i.replace(" ", "_"): "false"
            for i in SHEETS if i != "Info"
        }
        config["MIDPOINT ORB FACTORS"] = {
            "Conjunction": 2,
            "Semi-Sextile": 1,
            "Semi-Square": 1.5,
            "Sextile": 1,
            "Quintile": 1,
            "Square": 2,
            "Trine": 1,
            "Sesquiquadrate": 1.5,
            "BiQuintile": 1,
            "Quincunx": 1,
            "Opposite": 2
        }
        config.write(f)


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
            return [[*SIGNS][i], degree - (30 * i)]


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


def check_update(icons):
    try:
        new = urlopen(
            "https://raw.githubusercontent.com/dildeolupbiten"
            "/TkAstroDb/master/README.md"
        ).read().decode()
    except URLError:
        MsgBox(
            title="Warning",
            message="Couldn't connect.",
            level="warning",
            icons=icons
        )
        return
    with open("README.md", "r", encoding="utf-8") as f:
        old = f.read()[:-1]
    if new != old:
        with open("README.md", "w", encoding="utf-8") as f:
            f.write(new)
    try:
        scripts = json.load(
            urlopen(
                url=f"https://api.github.com/repos/dildeolupbiten/"
                    f"TkAstroDb/contents/Scripts?ref=master"
            )
        )
    except URLError:
        MsgBox(
            title="Warning",
            message="Couldn't connect.",
            level="warning",
            icons=icons
        )
        return
    update = False
    for i in scripts:
        try:
            file = urlopen(i["download_url"]).read().decode()
        except URLError:
            MsgBox(
                title="Warning",
                message="Couldn't connect.",
                level="warning",
                icons=icons
            )
            return
        if i["name"] not in os.listdir("Scripts"):
            update = True
            with open(f"Scripts/{i['name']}", "w", encoding="utf-8") as f:
                f.write(file)
        else:
            with open(f"Scripts/{i['name']}", "r", encoding="utf-8") as f:
                if file != f.read():
                    update = True
                    with open(
                            f"Scripts/{i['name']}",
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
        if os.path.exists("defaults.ini"):
            os.remove("defaults.ini")
        if os.name == "posix":
            Popen(["python3", "run.py"])
            os.kill(os.getpid(), __import__("signal").SIGKILL)
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


def get_basic_dict(values, indexes, constants, sub_index=(0, 0)):
    if sub_index == (0, 0):
        i1, i2 = 1, 13
    else:
        i1, i2 = sub_index
    return {
        list(constants[0])[index]: {
            list(constants[1])[ind]: value
            for ind, value in enumerate(i[i1:i2])
        } for index, i in enumerate(values[indexes[0]: indexes[1]])
    }


def get_aspect_dict(values, indexes, constants):
    return {
        list(constants[0])[index]: {
            list(constants[0])[ind]: j[index]
            for ind, j in enumerate(
                values[indexes[0]:indexes[1]][index + 1:], index + 1
            )
        }
        for index, i in enumerate(values[indexes[0]: indexes[1]])
    }


def get_3d_pattern_dict(values, c):
    result = {}
    for index, planet in enumerate(PLANETS):
        constants = list(PLANETS)[:index] + list(PLANETS)[index + 1:]
        result[planet] = get_aspect_dict(values, [c, c + 13], [constants])
        c += 15
    return result


def get_4d_pattern_dict(values, c):
    result = {}
    for index, planet in enumerate(PLANETS):
        planets = list(PLANETS)[:index] + list(PLANETS)[index + 1:]
        result[planet] = {}
        for _index, _planet in enumerate(planets):
            constants = planets[:_index] + planets[_index + 1:]
            result[planet][_planet] = get_aspect_dict(
                values,
                [c, c + 12],
                [constants]
            )
            c += 14
    return result


def get_planet_dict(values, planets, c, arrays):
    result = {}
    if planets:
        array = planets
    else:
        array = range(1, 13)
    for i in array:
        result[f"Lord-{i}" if not planets else i] = get_basic_dict(
            values,
            [c, c + 12],
            arrays,
            [2, 14]
        )
        c += 15
    return result


def only_planets(planets):
    return [
        planet for planet in planets
        if planets[planet]["symbol"]
    ]


def from_xml(filename):
    database = []
    category_names = []
    tree = ET.parse(filename)
    root = tree.getroot()
    for i in range(1000000):
        try:
            user_data = []
            for gender, roddenrating, bdata, adb_link, categories in zip(
                root[i + 2][1].findall("gender"),
                root[i + 2][1].findall("roddenrating"),
                root[i + 2][1].findall("bdata"),
                root[i + 2][2].findall("adb_link"),
                root[i + 2][3].findall("categories")
            ):
                name = root[i + 2][1][0].text
                sbdate_dmy = bdata[1].text
                sbtime = bdata[2].text
                jd_ut = bdata[2].get("jd_ut")
                lat = bdata[3].get("slati")
                lon = bdata[3].get("slong")
                place = bdata[3].text
                country = bdata[4].text
                categories = [
                    (
                        categories[j].get("cat_id"),
                        categories[j].text
                    )
                    for j in range(len(categories))
                ]
                for category in categories:
                    if category[1] and category[1] not in category_names:
                        category_names.append(category[1])
                user_data.append(int(root[i + 2].get("adb_id")))
                user_data.append(name)
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
                user_data.append(categories)
                if len(user_data) != 0:
                    database.append(user_data)
        except IndexError:
            break
    return database, sorted(category_names)


def get_element(sign: str):
    if sign in ["Aries", "Leo", "Sagittarius"]:
        return "Fire"
    elif sign in ["Taurus", "Virgo", "Capricorn"]:
        return "Earth"
    elif sign in ["Gemini", "Libra", "Aquarius"]:
        return "Air"
    elif sign in ["Cancer", "Scorpio", "Pisces"]:
        return "Water"


def get_mode(sign: str):
    if sign in ["Aries", "Cancer", "Libra", "Capricorn"]:
        return "Cardinal"
    elif sign in ["Taurus", "Leo", "Scorpio", "Aquarius"]:
        return "Fixed"
    elif sign in ["Gemini", "Virgo", "Sagittarius", "Pisces"]:
        return "Mutable"
        
        
def find_aspect(aspects, temporary, orb, aspect, planet1, planet2):
    if (
        0 < aspect < float(orb["conjunction"])
        or
        360 - float(orb["conjunction"]) < aspect < 360
    ):
        aspects["conjunction"][planet1][planet2] += 1
        if temporary:
            temporary["conjunction"][planet1][planet2] += 1
    elif (
        30 - float(orb["semi-sextile"]) <
        aspect < 30 + float(orb["semi-sextile"])
        or
        330 - float(orb["semi-sextile"]) <
        aspect < 330 + float(orb["semi-sextile"])
    ):
        aspects["semi-sextile"][planet1][planet2] += 1
        if temporary:
            temporary["semi-sextile"][planet1][planet2] += 1
    elif (
        45 - float(orb["semi-square"]) <
        aspect < 45 + float(orb["semi-square"])
        or
        315 - float(orb["semi-square"]) <
        aspect < 315 + float(orb["semi-square"])
    ):
        aspects["semi-square"][planet1][planet2] += 1
        if temporary:
            temporary["semi-square"][planet1][planet2] += 1
    elif (
        60 - float(orb["sextile"]) <
        aspect < 60 + float(orb["sextile"])
        or
        300 - float(orb["sextile"]) <
        aspect < 300 + float(orb["sextile"])
    ):
        aspects["sextile"][planet1][planet2] += 1
        if temporary:
            temporary["sextile"][planet1][planet2] += 1
    elif (
        72 - float(orb["quintile"]) <
        aspect < 72 + float(orb["quintile"])
        or
        288 - float(orb["quintile"]) <
        aspect < 288 + float(orb["quintile"])
    ):
        aspects["quintile"][planet1][planet2] += 1
        if temporary:
            temporary["quintile"][planet1][planet2] += 1
    elif (
        90 - float(orb["square"]) <
        aspect < 90 + float(orb["square"]) or
        270 - float(orb["square"]) <
        aspect < 270 + float(orb["square"])
    ):
        aspects["square"][planet1][planet2] += 1
        if temporary:
            temporary["square"][planet1][planet2] += 1
    elif (
        120 - float(orb["trine"]) <
        aspect < 120 + float(orb["trine"])
        or
        240 - float(orb["trine"]) <
        aspect < 240 + float(orb["trine"])
    ):
        aspects["trine"][planet1][planet2] += 1
        if temporary:
            temporary["trine"][planet1][planet2] += 1
    elif (
        135 - float(orb["sesquiquadrate"]) <
        aspect < 135 + float(orb["sesquiquadrate"])
        or
        225 - float(orb["sesquiquadrate"]) <
        aspect < 225 + float(orb["sesquiquadrate"])
    ):
        aspects["sesquiquadrate"][planet1][planet2] += 1
        if temporary:
            temporary["sesquiquadrate"][planet1][planet2] += 1
    elif (
        144 - float(orb["biquintile"]) <
        aspect < 144 + float(orb["biquintile"])
        or
        216 - float(orb["biquintile"]) <
        aspect < 216 + float(orb["biquintile"])
    ):
        aspects["biquintile"][planet1][planet2] += 1
        if temporary:
            temporary["biquintile"][planet1][planet2] += 1
    elif (
        150 - float(orb["quincunx"]) <
        aspect < 150 + float(orb["quincunx"])
        or
        210 - float(orb["quincunx"]) <
        aspect < 210 + float(orb["quincunx"])
    ):
        aspects["quincunx"][planet1][planet2] += 1
        if temporary:
            temporary["quincunx"][planet1][planet2] += 1
    elif (
        180 - float(orb["opposite"]) <
        aspect < 180 + float(orb["opposite"])
    ):
        aspects["opposite"][planet1][planet2] += 1
        if temporary:
            temporary["opposite"][planet1][planet2] += 1
            
            
def find_midpoints(midpoints, orb, aspect, planet1, planet2, planet3):
    if (
        0 < aspect < orb["Conjunction"] 
        or 
        360 - orb["Conjunction"] < aspect < 360
    ):
        midpoints[planet1][planet2]["Conjunction"][planet3] += 1
    elif (
        30 - orb["Semi-Sextile"] < aspect < 30 + orb["Semi-Sextile"]
        or
        330 - orb["Semi-Sextile"] < aspect < 330 + orb["Semi-Sextile"]
    ):
        midpoints[planet1][planet2]["Semi-Sextile"][planet3] += 1
    elif (
        45 - orb["Semi-Square"] < aspect < 45 + orb["Semi-Square"]
        or
        315 - orb["Semi-Square"] < aspect < 315 + orb["Semi-Square"]
    ):
        midpoints[planet1][planet2]["Semi-Square"][planet3] += 1
    elif (
        60 - orb["Sextile"] < aspect < 60 + orb["Sextile"]
        or
        300 - orb["Sextile"] < aspect < 300 + orb["Sextile"]
    ):
        midpoints[planet1][planet2]["Sextile"][planet3] += 1
    elif (
        72 - orb["Quintile"] < aspect < 72 + orb["Quintile"]
        or
        288 - orb["Quintile"] < aspect < 288 + orb["Quintile"]
    ):
        midpoints[planet1][planet2]["Quintile"][planet3] += 1
    elif (
        90 - orb["Square"] < aspect < 90 + orb["Square"]
        or
        270 - orb["Square"] < aspect < 270 + orb["Square"]
    ):
        midpoints[planet1][planet2]["Square"][planet3] += 1
    elif (
        120 - orb["Trine"] < aspect < 120 + orb["Trine"]
        or
        240 - orb["Trine"] < aspect < 240 + orb["Trine"]
    ):
        midpoints[planet1][planet2]["Trine"][planet3] += 1
    elif (
        135 - orb["Sesquiquadrate"] < aspect < 135 + orb["Sesquiquadrate"]
        or
        225 - orb["Sesquiquadrate"] < aspect < 225 + orb["Sesquiquadrate"]
    ):
        midpoints[planet1][planet2]["Sesquiquadrate"][planet3] += 1
    elif (
        144 - orb["Biquintile"] < aspect < 144 + orb["Biquintile"]
        or
        216 - orb["Biquintile"] < aspect < 216 + orb["Biquintile"]
    ):
        midpoints[planet1][planet2]["Biquintile"][planet3] += 1
    elif (
        150 - orb["Quincunx"] < aspect < 150 + orb["Quincunx"]
        or
        210 - orb["Quincunx"] < aspect < 210 + orb["Quincunx"]
    ):
        midpoints[planet1][planet2]["Quincunx"][planet3] += 1
    elif (
        180 - orb["Opposite"] < aspect < 180 + orb["Opposite"]
    ):
        midpoints[planet1][planet2]["Opposite"][planet3] += 1


def find_midpoint(aspect1, aspect2):
    if aspect1 > aspect2:
        if aspect1 - aspect2 >= 180:
            return ((aspect1 + (360 + aspect2)) / 2) % 360
        else:
            return ((aspect1 + aspect2) / 2) % 360
    elif aspect2 > aspect1:
        if aspect2 - aspect1 >= 180:
            return ((aspect2 + (360 + aspect1)) / 2) % 360
        else:
            return ((aspect2 + aspect1) / 2) % 360


def create_midpoint_dict(planets):
    result = {}
    for i in planets:
        result[i] = {}
        for j in planets:
            if j not in result:
                result[i][j] = {}
                for aspect in ASPECTS:
                    result[i][j][aspect] = {}
                    for k in planets:
                        if k not in [i, j]:
                            result[i][j][aspect][k] = 0
    return result


def get_midpoint_dict(df):
    midpoints = create_midpoint_dict(PLANETS)
    planet3s = [i for i in df.columns[3:]]
    for i in df.values:
        if i[0] is np.nan:
            planet3s = i[3:]
        planet1, planet2, aspect, values = "", "", "", []
        if i[0] is not np.nan and "Total" not in i[0]:
            pair, aspect = i[0].split("(")
            pair = pair.strip()
            aspect = aspect[:-1]
            planet1, planet2 = pair.split(" / ")
            values = [i for i in i[3:]]
        if all([planet1, planet2, aspect, values]):
            for planet3, value in zip(planet3s, values):
                midpoints[planet1][planet2][aspect][planet3] = value
    return midpoints
    
    
def edit_coordinate(coordinate: int, dist: str):
    if dist == "latitude":
        if coordinate > 0:
            return str(abs(coordinate)) + "n"
        else:
            return str(abs(coordinate)) + "s"
    elif dist == "longitude":
        if coordinate > 0:
            return str(abs(coordinate)) + "e"
        else:
            return str(abs(coordinate)) + "w"
            
            
def get_info(df):
    info = {df.columns[0]: df.columns[2]}
    for i in df.values:
        if i[0] in ASPECTS:
            if "Orb Factor" not in info:
                info["Orb Factor"] = {}
            if "Midpoint Orb Factor" not in info:
                info["Midpoint Orb Factor"] = {}
            info["Orb Factor"][i[0]] = i[2]
            info["Midpoint Orb Factor"][i[0]] = i[4]
        else:
            if i[0].startswith("Orb") or i[0].startswith("Aspect"):
                continue
            info[i[0]] = i[2]
    return info
