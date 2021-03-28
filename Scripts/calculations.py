# -*- coding: utf-8 -*-

from .zodiac import Zodiac
from .spreadsheet import Spreadsheet
from .messagebox import MsgBox, ChoiceBox
from .utilities import (
    convert_coordinates, progressbar, get_basic_dict,
    get_planet_dict, get_aspect_dict, only_planets,
    get_3d_pattern_dict, get_4d_pattern_dict,
    create_midpoint_dict, get_midpoint_dict, get_info
)
from .modules import (
    os, dt, pd, tk, ttk, time, binom, shutil,
    Thread, variance, ConfigParser
)
from .constants import (
    HOUSE_SYSTEMS, PLANETS, SIGNS, HOUSES, SHEETS,
    TRADITIONAL_RULERSHIP, MODERN_RULERSHIP, AYANAMSHA
)


def send_warning_message(icons):
    MsgBox(
        title="Warning",
        level="warning",
        message="You can't run many processes \nat the same time.",
        icons=icons
    )


def create_normal_dict(lists, keys):
    return {
        f"{keys[0]}{i}": create_normal_dict(lists=lists[1:], keys=keys[1:])
        if len(lists[1:]) != 0 else 0
        for i in lists[0]
    }


def create_enumerate_dict(lists, keys):
    if len(lists) == 2:
        return {
            f"{keys[0]}{i}": {
                f"{keys[1]}{j}": 0 for j in list(lists[1])[index:]
            } for index, i in enumerate(lists[0], 1)
        }
    elif len(lists) == 3:
        return {
            f"{keys[0]}{i}": {
                f"{keys[1]}{j}": {
                    f"{keys[2]}{k}": 0
                    for k in list(lists[2])[index:]
                } for index, j in enumerate(lists[1], 1)
            } for i in lists[0]
        }


def create_3d_dict(planets):
    result = {}
    for i, p in enumerate(planets):
        result[p] = {}
        _planets = planets[:i] + planets[i + 1:]
        for _i, _p in enumerate(_planets):
            result[p][_p] = {}
            planets_ = _planets[_i + 1:]
            for i_, p_ in enumerate(planets_):
                result[p][_p][p_] = 0
    return result


def create_4d_dict(planets):
    result = {}
    for i, p in enumerate(planets):
        result[p] = {}
        _planets = planets[:i] + planets[i + 1:]
        for _i, _p in enumerate(_planets):
            result[p][_p] = {}
            planets_ = _planets[:_i] + _planets[_i + 1:]
            for i_, p_ in enumerate(planets_):
                result[p][_p][p_] = {}
                __planets = planets_[i_ + 1:]
                for __i, __p in enumerate(__planets):
                    result[p][_p][p_][__p] = 0
    return result


def get_aspects(aspects, aspect_type):
    result = {}
    for keys, values in aspects[aspect_type].items():
        if keys not in result:
            result[keys] = []
        for k, v in values.items():
            if v:
                result[keys].append(k)
                if k in result:
                    result[k] += [keys]
                else:
                    result[k] = [keys]
    return result


def three_point(aspects, aspect_type):
    if aspect_type == "yod":
        first = "quincunx"
        second = "sextile"
    elif aspect_type == "grand trine":
        first = "trine"
        second = "trine"
    elif aspect_type == "t-square":
        first = "square"
        second = "opposite"
    else:
        return
    first = get_aspects(aspects, first)
    second = get_aspects(aspects, second)
    result = []
    for k, v in second.items():
        for i in v:
            for j in first[i]:
                if (
                    j in first[k]
                    and
                    j in first[i]
                    and
                    sorted([k, i, j]) not in
                    [sorted(m) for m in result]
                ):
                    result += [[k, i, j]]
    return result


def four_point(aspects, aspect_type):
    if aspect_type in ["mystic rectangle", "kite"]:
        first = "sextile"
        second = "trine"
        third = "opposite"
    elif aspect_type == "grand cross":
        first = "square"
        second = "square"
        third = "opposite"
    else:
        return
    first = get_aspects(aspects, first)
    second = get_aspects(aspects, second)
    third = get_aspects(aspects, third)
    result = []
    for k, v in third.items():
        for i in v:
            for j in second[i]:
                if j in first[k]:
                    if aspect_type == "kite":
                        for m in second[j]:
                            if (
                                m in first[k]
                                and
                                m in second[i]
                                and
                                sorted([k, i, j, m]) not in
                                [sorted(n) for n in result]
                            ):
                                result += [[k, i, j, m]]
                    else:
                        for m in third[j]:
                            if (
                                m in first[i]
                                and
                                m in second[k]
                                and
                                sorted([k, i, j, m]) not in
                                [sorted(n) for n in result]
                            ):
                                result += [[k, i, j, m]]
    return result


def no_subset(subset, superset):
    result = [i for i in subset]
    for i in subset:
        for j in superset:
            if all(k in j for k in i):
                result.remove(i)
                break
    return result


def get_yod(aspects):
    return three_point(aspects, "yod")


def get_t_square(aspects):
    return no_subset(
        subset=three_point(aspects, "t-square"),
        superset=get_grand_cross(aspects)
    )


def get_grand_trine(aspects):
    return no_subset(
        subset=three_point(aspects, "grand trine"),
        superset=get_kite(aspects)
    )


def get_mystic_rectangle(aspects):
    return four_point(aspects, "mystic rectangle")


def get_grand_cross(aspects):
    return four_point(aspects, "grand cross")


def get_kite(aspects):
    return four_point(aspects, "kite")


def find_observed_values(widget, icons, menu, version):
    config = ConfigParser()
    config.read("defaults.ini")
    if all(v == "false" for k, v in config["TABLE SELECTION"].items()):
        MsgBox(
            title="Warning",
            level="warning",
            message="Select a table from Spreadsheet\nmenu cascade.",
            icons=icons
        )
        return
    displayed_results = []
    selected_categories = []
    ignored_categories = []
    selected_ratings = []
    checkbuttons = {}
    year_from = ""
    year_to = ""
    latitude_from = ""
    latitude_to = ""
    longitude_from = ""
    longitude_to = ""
    for i in widget.winfo_children():
        if hasattr(i, "included"):
            year_from += i.ranges["Year"].widgets["From"].get()
            year_to += i.ranges["Year"].widgets["To"].get()
            latitude_from += i.ranges["Latitude"].widgets["From"].get()
            latitude_to += i.ranges["Latitude"].widgets["To"].get()
            longitude_from += i.ranges["Longitude"].widgets["From"].get()
            longitude_to += i.ranges["Longitude"].widgets["To"].get()
            displayed_results += i.displayed_results
            selected_categories += i.included
            ignored_categories += i.ignored
            selected_ratings += i.selected_ratings
            checkbuttons.update(i.checkbuttons)
            break
    if not displayed_results:
        MsgBox(
            title="Warning",
            level="warning",
            message="No records are selected.",
            icons=icons
        )
        return
    save_categories = [i for i in selected_categories]
    if selected_categories:
        if len(selected_categories) > 1:
            selected_categories = "Control_Group"
        else:
            selected_categories = selected_categories[0]\
                .replace("/ ", "_").replace("/", "_").replace(" : ", "/")
    else:
        if len(displayed_results) == 1:
            selected_categories = displayed_results[0][1]\
                .replace(", ", "_-_")
        else:
            selected_categories = "Gathered Manually"
    if selected_ratings:
        selected_ratings = "+".join(selected_ratings)
    else:
        selected_ratings = "None"
    if year_from and year_to:
        year_range = f"{year_from} - {year_to}"
    else:
        year_range = "None"
    if latitude_from and latitude_to:
        latitude_range = f"{latitude_from} - {latitude_to}"
    else:
        latitude_range = "None"
    if longitude_from and longitude_to:
        longitude_range = f"{longitude_from} - {longitude_to}"
    else:
        longitude_range = "None"
    if config["ZODIAC"]["selected"] == "Sidereal":
        zodiac = f"{config['ZODIAC']['selected']} " \
                 f"(Ayanamsha: {config['AYANAMSHA']['selected']})"
    else:
        zodiac = config["ZODIAC"]["selected"]
    info = {
        "Adb Version": config["DATABASE"]["selected"]
        .replace(".json", "").replace(".xml", ""),
        "TkAstroDb Version": version,
        "Zodiac": zodiac,
        "House System": config["HOUSE SYSTEM"]["selected"],
        "Rodden Rating": selected_ratings,
        "Category": selected_categories,
        "Year Range": year_range,
        "Latitude Range": latitude_range,
        "Longitude Range": longitude_range
    }
    if "event" in checkbuttons:
        info.update(
            {
                key.title(): "True" if value[0].get() == "0" else "False"
                for key, value in checkbuttons.items()
            }
        )
    else:
        info = {"Event": "False", "Human": "True"}
        info.update(
            {
                key.title(): "True" if value[0].get() == "0" else "False"
                for key, value in checkbuttons.items()
            }
        )
    path = os.path.join(
        *selected_categories.split("/"),
        f"RR_{selected_ratings}",
        f"ORB_{'_'.join(config['ORB FACTORS'].values())}",
        zodiac,
        config["HOUSE SYSTEM"]["selected"]
    )
    if info["Event"] == "False" and info["Human"] == "True":
        if info["Male"] == "False" and info["Female"] == "True":
            path = os.path.join(path, "Female")
        elif info["Male"] == "True" and info["Female"] == "False":
            path = os.path.join(path, "Male")
        else:
            path = os.path.join(path, "Human")
    elif info["Event"] == "True" and info["Human"] == "False":
        path = os.path.join(path, "Event")
    else:
        if info["Male"] == "False" and info["Female"] == "True":
            path = os.path.join(path, "Event+Female")
        elif info["Male"] == "True" and info["Female"] == "False":
            path = os.path.join(path, "Event+Male")
        else:
            path = os.path.join(path, "Event+Human")
    if (
        info["South Hemisphere"] == "False"
        and
        info["North Hemisphere"] == "True"
    ):
        path = os.path.join(path, "North")
    elif (
        info["South Hemisphere"] == "True"
        and
        info["North Hemisphere"] == "False"
    ):
        path = os.path.join(path, "South")
    else:
        path = path
    if (
        info["West Hemisphere"] == "False"
        and
        info["East Hemisphere"] == "True"
    ):
        path = os.path.join(path, "East")
    elif (
        info["West Hemisphere"] == "True"
        and
        info["East Hemisphere"] == "False"
    ):
        path = os.path.join(path, "West")
    else:
        path = path
    if os.path.exists(os.path.join(path, "observed_values.xlsx")):
        msg = "The file already exists.\n Do you want to continue?\n" \
              " If you press 'OK', \nthe file will be overwritten. \n" \
              "If you press 'Cancel',\n you will return to the main " \
              "window. \nIt is recommended that you " \
              "\nreconsider your choices."
        choicebox = ChoiceBox(
            title="Warning",
            level="warning",
            message=msg,
            icons=icons,
            width=400,
            height=200,
        )
        if choicebox.choice:
            Thread(
                target=lambda: start_calculation(
                    displayed_results=displayed_results,
                    info=info,
                    config=config,
                    widget=widget,
                    icons=icons,
                    path=path,
                    selected_categories=selected_categories,
                    ignored_categories=ignored_categories,
                    menu=menu,
                    save_categories=save_categories,
                    version=version
                ),
                daemon=True
            ).start()
        else:
            return
    else:
        Thread(
            target=lambda: start_calculation(
                displayed_results=displayed_results,
                info=info,
                config=config,
                widget=widget,
                icons=icons,
                path=path,
                selected_categories=selected_categories,
                ignored_categories=ignored_categories,
                menu=menu,
                save_categories=save_categories,
                version=version
            ),
            daemon=True
        ).start()


def start_calculation(
        displayed_results,
        info,
        config,
        widget,
        icons,
        path,
        selected_categories,
        ignored_categories,
        menu,
        save_categories,
        version,
):
    if config["TABLE SELECTION"]["planets_in_signs"] == "true":
        planets_in_signs = create_normal_dict(
            lists=[only_planets(PLANETS), SIGNS],
            keys=["", ""]
        )
    else:
        planets_in_signs = {}
    if config["TABLE SELECTION"]["planets_in_elements"] == "true":
        planets_in_elements = create_normal_dict(
            lists=[only_planets(PLANETS), ["Fire", "Earth", "Air", "Water"]],
            keys=["", ""]
        )
    else:
        planets_in_elements = {}
    if config["TABLE SELECTION"]["planets_in_modes"] == "true":
        planets_in_modes = create_normal_dict(
            lists=[only_planets(PLANETS), ["Cardinal", "Fixed", "Mutable"]],
            keys=["", ""]
        )
    else:
        planets_in_modes = {}
    if config["TABLE SELECTION"]["houses_in_signs"] == "true":
        houses_in_signs = create_normal_dict(
            lists=[range(1, 13), SIGNS],
            keys=["House-", ""]
        )
    else:
        houses_in_signs = {}
    if config["TABLE SELECTION"]["houses_in_elements"] == "true":
        houses_in_elements = create_normal_dict(
            lists=[range(1, 13), ["Fire", "Earth", "Air", "Water"]],
            keys=["House-", ""]
        )
    else:
        houses_in_elements = {}
    if config["TABLE SELECTION"]["houses_in_modes"] == "true":
        houses_in_modes = create_normal_dict(
            lists=[range(1, 13), ["Cardinal", "Fixed", "Mutable"]],
            keys=["House-", ""]
        )
    else:
        houses_in_modes = {}
    if config["TABLE SELECTION"]["planets_in_houses"] == "true":
        planets_in_houses = create_normal_dict(
            lists=[only_planets(PLANETS), range(1, 13)],
            keys=["", "House-"]
        )
    else:
        planets_in_houses = {}
    if config["TABLE SELECTION"]["planets_in_houses_in_signs"] == "true":
        planets_in_houses_in_signs = create_normal_dict(
            lists=[only_planets(PLANETS), range(1, 13), SIGNS],
            keys=["", "House-", ""]
        )
    else:
        planets_in_houses_in_signs = {}
    if config["TABLE SELECTION"]["detailed_traditional_rulership"] == "true":
        detailed_traditional_rulership = create_normal_dict(
            lists=[
                range(1, 13),
                TRADITIONAL_RULERSHIP.values(),
                range(1, 13)
            ],
            keys=["Lord-", "", "House-"]
        )
    else:
        detailed_traditional_rulership = {}
    if config["TABLE SELECTION"]["detailed_modern_rulership"] == "true":
        detailed_modern_rulership = create_normal_dict(
            lists=[range(1, 13), MODERN_RULERSHIP.values(), range(1, 13)],
            keys=["Lord-", "", "House-"]
        )
    else:
        detailed_modern_rulership = {}
    if config["TABLE SELECTION"]["basic_traditional_rulership"] == "true":
        basic_traditional_rulership = create_normal_dict(
            lists=[range(1, 13), range(1, 13)],
            keys=["Lord-", "House-"]
        )
    else:
        basic_traditional_rulership = {}
    if config["TABLE SELECTION"]["basic_modern_rulership"] == "true":
        basic_modern_rulership = create_normal_dict(
            lists=[range(1, 13), range(1, 13)],
            keys=["Lord-", "House-"]
        )
    else:
        basic_modern_rulership = {}
    if config["TABLE SELECTION"]["aspects"] == "true":
        aspects = create_enumerate_dict(
            lists=[list(config["ORB FACTORS"]), PLANETS, PLANETS],
            keys=["", "", ""],
        )
    else:
        aspects = {}
    if config["TABLE SELECTION"]["sum_of_aspects"] == "true":
        sum_of_aspects = create_enumerate_dict(
            lists=[PLANETS, PLANETS],
            keys=["", ""]
        )
    else:
        sum_of_aspects = {}
    if config["TABLE SELECTION"]["yod"] == "true":
        yod = create_3d_dict(list(PLANETS))
    else:
        yod = {}
    if config["TABLE SELECTION"]["t-square"] == "true":
        t_square = create_3d_dict(list(PLANETS))
    else:
        t_square = {}
    if config["TABLE SELECTION"]["grand_trine"] == "true":
        grand_trine = create_3d_dict(list(PLANETS))
    else:
        grand_trine = {}
    if config["TABLE SELECTION"]["mystic_rectangle"] == "true":
        mystic_rectangle = create_4d_dict(list(PLANETS))
    else:
        mystic_rectangle = {}
    if config["TABLE SELECTION"]["grand_cross"] == "true":
        grand_cross = create_4d_dict(list(PLANETS))
    else:
        grand_cross = {}
    if config["TABLE SELECTION"]["kite"] == "true":
        kite = create_4d_dict(list(PLANETS))
    else:
        kite = {}
    if config["TABLE SELECTION"]["midpoints"] == "true":
        midpoints = create_midpoint_dict(PLANETS)
        midpoint_orb_factors = {
            i.title(): float(config["MIDPOINT ORB FACTORS"][i])
            for i in config["MIDPOINT ORB FACTORS"]
        }
    else:
        midpoints = {}
        midpoint_orb_factors = {}
    orb_factors = config["ORB FACTORS"]
    size = len(displayed_results) + 1
    received = 0
    now = time.time()
    pframe = tk.Frame(master=widget)
    pbar = ttk.Progressbar(
        master=pframe,
        orient="horizontal",
        length=200,
        mode="determinate"
    )
    pstring = tk.StringVar()
    plabel = tk.Label(master=pframe, textvariable=pstring)
    pframe.pack()
    pbar.pack(side="left")
    plabel.pack(side="left")
    log = open("output.log", "w", encoding="utf-8")
    if selected_categories != "Control_Group":
        log.write(
            f"Adb Version: {info['Adb Version']}\n"
            f"TkAstroDb Version: {version}\n"
            f"Zodiac: {info['Zodiac']}\n"
            f"House System: {info['House System']}\n"
            f"Rodden Rating: {info['Rodden Rating']}\n"
            f"Orb Factor: {'_'.join(config['ORB FACTORS'].values())}\n"
            f"Category: {info['Category']}\n"
            f"Year Range: {info['Year Range']}\n"
            f"Latitude Range: {info['Latitude Range']}\n"
            f"Longitude Range: {info['Longitude Range']}\n"
        )
    else:
        log.write(
            f"Adb Version: {info['Adb Version']}\n"
            f"TkAstroDb Version: {version}\n"
            f"Zodiac: {info['Zodiac']}\n"
            f"House System: {info['House System']}\n"
            f"Rodden Rating: {info['Rodden Rating']}\n"
            f"Orb Factor: {'_'.join(config['ORB FACTORS'].values())}\n"
            f"Category: {info['Category']}\n"
            f"Year Range: {info['Year Range']}\n"
            f"Latitude Range: {info['Latitude Range']}\n"
            f"Longitude Range: {info['Longitude Range']}\n"
            f"Selected Categories:\n"
        )        
        for index, i in enumerate(save_categories, 1):
            log.write(f"\t{index}. {i}\n".expandtabs(20))
    log.write("Ignored:\n")
    for index, i in enumerate(ignored_categories, 1):
        log.write(f"\t{index}. {i}\n".expandtabs(20))
    log.write("\n")
    log.write(
        f"|{dt.now().strftime('%Y-%m-%d %H:%M:%S')}| Process started.\n\n"
    )
    log.flush()
    menu.entryconfigure(0, command=lambda: send_warning_message(icons=icons))
    total = 0
    for i in displayed_results:
        jd = float(i[6])
        lat = convert_coordinates(i[7])
        lon = convert_coordinates(i[8])
        if any(
            [
                yod,
                t_square,
                grand_trine,
                mystic_rectangle,
                grand_cross,
                kite
            ]
        ):
            temporary = create_enumerate_dict(
                lists=[list(config["ORB FACTORS"]), PLANETS, PLANETS],
                keys=["", "", ""],
            )
        else:
            temporary = {}
        try:
            Zodiac(
                jd=jd,
                lat=lat,
                lon=lon,
                hsys=HOUSE_SYSTEMS[config["HOUSE SYSTEM"]["selected"]],
                zodiac=config["ZODIAC"]["selected"],
                ayanamsha=AYANAMSHA[config["AYANAMSHA"]["selected"]]
            ).patterns(
                planets_in_signs,
                planets_in_elements,
                planets_in_modes,
                houses_in_signs,
                houses_in_elements,
                houses_in_modes,
                planets_in_houses,
                planets_in_houses_in_signs,
                basic_traditional_rulership,
                basic_modern_rulership,
                detailed_traditional_rulership,
                detailed_modern_rulership,
                aspects,
                temporary,
                midpoints,
                orb_factors,
                midpoint_orb_factors,
            )
        except BaseException as err:
            log.write(
                f"|{dt.now().strftime('%Y-%m-%d %H:%M:%S')}| "
                f"Error Type: '{err}'\n"
            )
            log.write(
                f"\tRecord: {i}\n\n".expandtabs(22)
            )
            log.flush()
            received += 1
            progressbar(
                s=size,
                r=received,
                n=now,
                pframe=pframe,
                pbar=pbar,
                plabel=plabel,
                pstring=pstring
            )
            continue
        if yod:
            special_aspect_3d_pattern(temporary, yod, get_yod)
        if t_square:
            special_aspect_3d_pattern(temporary, t_square, get_t_square)
        if grand_trine:
            special_aspect_3d_pattern(
                temporary,
                grand_trine,
                get_grand_trine,
                apex=False
            )
        if mystic_rectangle:
            special_aspect_4d_pattern(
                temporary,
                mystic_rectangle,
                get_mystic_rectangle,
                apex=False
            )
        if grand_cross:
            special_aspect_4d_pattern(
                temporary,
                grand_cross,
                get_grand_cross,
                apex=False
            )
        if kite:
            special_aspect_4d_pattern(
                temporary,
                kite,
                get_kite,
            )
        received += 1
        progressbar(
            s=size,
            r=received,
            n=now,
            pframe=pframe,
            pbar=pbar,
            plabel=plabel,
            pstring=pstring
        )
        total += 1
    if sum_of_aspects:
        for aspect in aspects:
            for index, planet in enumerate(PLANETS, 1):
                for _planet in list(PLANETS)[index:]:
                    sum_of_aspects[planet][_planet] += \
                        aspects[aspect][planet][_planet]
    log.write(
        f"|{dt.now().strftime('%Y-%m-%d %H:%M:%S')}| Process finished."
    )
    log.close()
    if not os.path.exists(os.path.join(".", path)):
        os.makedirs(path)
    filename = os.path.join(path, "observed_values.xlsx")
    info["Number Of Records"] = total
    info["Orb Factor"] = {
        k.title(): v for k, v in config["ORB FACTORS"].items()
    }
    info["Midpoint Orb Factor"] = {
        k.title(): v for k, v in config["MIDPOINT ORB FACTORS"].items()
    }
    Spreadsheet(
        filename=filename,
        info=info,
        planets_in_signs=planets_in_signs,
        planets_in_elements=planets_in_elements,
        planets_in_modes=planets_in_modes,
        houses_in_signs=houses_in_signs,
        houses_in_elements=houses_in_elements,
        houses_in_modes=houses_in_modes,
        planets_in_houses=planets_in_houses,
        aspects=aspects,
        sum_of_aspects=sum_of_aspects,
        planets_in_houses_in_signs=planets_in_houses_in_signs,
        basic_traditional_rulership=basic_traditional_rulership,
        basic_modern_rulership=basic_modern_rulership,
        detailed_traditional_rulership=detailed_traditional_rulership,
        detailed_modern_rulership=detailed_modern_rulership,
        yod=yod,
        t_square=t_square,
        grand_trine=grand_trine,
        mystic_rectangle=mystic_rectangle,
        grand_cross=grand_cross,
        kite=kite,
        midpoints=midpoints,
    )
    received += 1
    progressbar(
        s=size,
        r=received,
        n=now,
        pframe=pframe,
        pbar=pbar,
        plabel=plabel,
        pstring=pstring
    )
    shutil.move(
        src=os.path.join(os.getcwd(), "output.log"),
        dst=os.path.join(path, "output.log")
    )
    widget.after(
        0, 
        lambda: MsgBox(
            icons=icons,
            title="Info",
            level="info",
            message="Calculation process is completed!"
        )
    )
    menu.entryconfigure(
        0,
        command=lambda: Thread(
            target=lambda: find_observed_values(
                widget=widget,
                icons=icons,
                menu=menu,
                version=version
            ),
            daemon=True
        ).start()
    )


def special_aspect_3d_pattern(temporary, pattern, func, apex=True):
    for i in func(temporary):
        if not apex:
            ordered = sorted(i, key=list(PLANETS).index)
            pattern[ordered[0]][ordered[1]][ordered[2]] += 1
            pattern[ordered[1]][ordered[0]][ordered[2]] += 1
            pattern[ordered[2]][ordered[0]][ordered[1]] += 1
        else:
            ordered = sorted(i[:-1], key=list(PLANETS).index)
            pattern[i[-1]][ordered[0]][ordered[1]] += 1


def special_aspect_4d_pattern(temporary, pattern, func, apex=True):
    for i in func(temporary):
        if not apex:
            ordered1 = sorted(i[2:], key=list(PLANETS).index)
            pattern[i[0]][i[1]][ordered1[0]][ordered1[1]] += 1
            pattern[i[1]][i[0]][ordered1[0]][ordered1[1]] += 1
            ordered2 = sorted(i[:2], key=list(PLANETS).index)
            pattern[i[2]][i[3]][ordered2[0]][ordered2[1]] += 1
            pattern[i[3]][i[2]][ordered2[0]][ordered2[1]] += 1
        else:
            ordered = sorted(i[2:], key=list(PLANETS).index)
            pattern[i[0]][i[1]][ordered[0]][ordered[1]] += 1


def get_values(filename):
    dfs = [
        pd.read_excel(filename, sheet_name=name)
        for name in SHEETS
    ]
    info = get_info(dfs[0])
    total = info["Number Of Records"]
    config = ConfigParser()
    config.read("defaults.ini")
    elements = ["Fire", "Earth", "Air", "Water"]
    modes = ["Cardinal", "Fixed", "Mutable"]
    planets = only_planets(PLANETS)
    pattern_2d = []
    for i, j, k, m in zip(
        range(1, 8),
        [planets] * 3 + [HOUSES] * 3 + [planets],
        [SIGNS, elements, modes] * 2 + [HOUSES],
        [(1, 13), (1, 5), (1, 4), (1, 13), (1, 5), (1, 4), (1, 13)]
    ):
        values = dfs[i].values
        if len(values) != 0:
            pattern_2d.append(
                get_basic_dict(
                    values=values,
                    indexes=[0, 12],
                    constants=[j, k],
                    sub_index=m
                )
            )
        else:
            pattern_2d.append({})
    values = dfs[8].values
    if len(values) != 0:
        planets_in_houses_in_signs = get_planet_dict(
            values=values,
            planets=planets,
            c=0,
            arrays=[HOUSES, SIGNS]
        )
    else:
        planets_in_houses_in_signs = {}
    values = dfs[9].values
    if len(values) != 0:
        basic_traditional_rulership = get_basic_dict(
            values=values,
            indexes=[0, 12],
            constants=[[f"Lord-{i}" for i in range(1, 13)], HOUSES]
        )
    else:
        basic_traditional_rulership = {}
    values = dfs[10].values
    if len(values) != 0:
        basic_modern_rulership = get_basic_dict(
            values=values,
            indexes=[0, 12],
            constants=[[f"Lord-{i}" for i in range(1, 13)], HOUSES]
        )
    else:
        basic_modern_rulership = {}
    values = dfs[11].values
    if len(values) != 0:
        detailed_traditional_rulership = get_planet_dict(
            values=values,
            planets=None,
            c=0,
            arrays=[list(TRADITIONAL_RULERSHIP.values()), HOUSES]
        )
    else:
        detailed_traditional_rulership = {}
    values = dfs[12].values
    if len(values) != 0:
        detailed_modern_rulership = get_planet_dict(
            values=values,
            planets=None,
            c=0,
            arrays=[list(MODERN_RULERSHIP.values()), HOUSES]
        )
    else:
        detailed_modern_rulership = {}
    values = dfs[13].values
    if len(values) != 0:
        c = 0
        aspects = {}
        for key in config["ORB FACTORS"]:
            aspects[key] = get_aspect_dict(
                values=values,
                indexes=[c, c + 14],
                constants=[PLANETS]
            )
            c += 16
    else:
        aspects = {}
    values = dfs[14].values
    if len(values) != 0:
        sum_of_aspects = get_aspect_dict(
            values=values,
            indexes=[0, 14],
            constants=[PLANETS]
        )
    else:
        sum_of_aspects = {}
    pattern_3d = []
    for i in range(15, 18):
        values = dfs[i].values
        if len(values) != 0:
            pattern_3d.append(
                get_3d_pattern_dict(values=values, c=0)
            )
        else:
            pattern_3d.append({})
    pattern_4d = []
    for i in range(18, 21):
        values = dfs[i].values
        if len(values) != 0:
            pattern_4d.append(
                get_4d_pattern_dict(values=values, c=0)
            )
        else:
            pattern_4d.append({})
    if len(dfs[21].values) != 0:
        midpoints = get_midpoint_dict(dfs[21])
    else:
        midpoints = {}
    return (
        total,
        *pattern_2d,
        aspects,
        sum_of_aspects,
        *pattern_3d,
        *pattern_4d,
        planets_in_houses_in_signs,
        basic_traditional_rulership,
        basic_modern_rulership,
        detailed_traditional_rulership,
        detailed_modern_rulership,
        info,
        midpoints
    )


def probability_mass_function(n, k, p):
    result = 0
    for i in range(k + 1):
        result += binom.pmf(n=n, k=i, p=p) * 100
    return round(result, 6)


def select_dict(
        d1,
        d2,
        method,
        calculation_type,
        cancel,
        x_total,
        y_total
):
    for (key, value), (_key, _value) in zip(d1.items(), d2.items()):
        save1 = {m: n for m, n in value.items()}
        save2 = {m: n for m, n in _value.items()}
        for (k, v), (_k, _v) in zip(value.items(), _value.items()):
            try:
                if calculation_type == "expected":
                    if method == "Subcategory":
                        d1[key][k] = _v * x_total / y_total
                    elif method == "Independent":
                        d1[key][k] = \
                            x_total * (v + _v) / (x_total + y_total)
                if calculation_type == "effect-size":
                    d1[key][k] = v / _v
                elif calculation_type == "chi-square":
                    d1[key][k] = (v - _v) ** 2 / _v
                elif calculation_type == "cohen's d":
                    if cancel:
                        d1[key][k] = ""
                    else:
                        d1[key][k] = (v - _v) / \
                            (
                                (
                                    variance(save1.values()) +
                                    variance(save2.values())
                                ) / 2
                            ) ** 0.5
                elif calculation_type == "binomial limit":
                    p1 = probability_mass_function(
                        n=x_total, k=v, p=_v / y_total
                    )
                    p2 = 100 - probability_mass_function(
                        n=x_total, k=v - 1, p=_v / y_total
                    )
                    if p1 < p2:
                        d1[key][k] = -p1
                    elif p1 > p2:
                        d1[key][k] = p2
            except ZeroDivisionError:
                d1[key][k] = 0
    
    
def select_calculation(
        icons,
        calculation_type,
        input1, 
        input2, 
        output, 
        widget
):
    if not os.path.exists(input1):
        MsgBox(
            icons=icons,
            title="Warning",
            level="warning",
            message=f"{input1} is not found."
        )
        return
    if not os.path.exists(input2):
        MsgBox(
            icons=icons,
            title="Warning",
            level="warning",
            message=f"{input2} is not found."
        )
        return
    size = 26
    received = 0
    now = time.time()
    pframe = tk.Frame(master=widget)
    pbar = ttk.Progressbar(
        master=pframe,
        orient="horizontal",
        length=200,
        mode="determinate"
    )
    pstring = tk.StringVar()
    plabel = tk.Label(master=pframe, textvariable=pstring)
    pframe.pack()
    pbar.pack(side="left")
    plabel.pack(side="left")
    x = get_values(filename=input1)
    received += 1
    progressbar(
        s=size,
        r=received,
        n=now,
        pframe=pframe,
        pbar=pbar,
        plabel=plabel,
        pstring=pstring
    )
    y = get_values(filename=input2)
    received += 1 
    progressbar(
        s=size,
        r=received,
        n=now,
        pframe=pframe,
        pbar=pbar,
        plabel=plabel,
        pstring=pstring
    )
    
    x_info = x[-2]
    y_info = y[-2]
    if calculation_type in ["expected", "binomial limit"]:
        for k in x_info:
            if k not in ["Orb Factor", "Midpoint Orb Factor"]:
                x_info[k] = f"{x_info[k]} & {y_info[k]}"
            else:
                for key, value in x_info[k].items():
                    x_info[k][key] = f"{x_info[k][key]} & {y_info[k][key]}"
    else:
        x_info = y_info
    config = ConfigParser()
    config.read("defaults.ini")
    method = config["METHOD"]["selected"]
    for i in range(len(y)):
        if i in [0, 21]:
            received += 1
            progressbar(
                s=size,
                r=received,
                n=now,
                pframe=pframe,
                pbar=pbar,
                plabel=plabel,
                pstring=pstring
            )
            continue
        if i in [1, 2, 3, 4, 5, 6, 7, 9, 17, 18]:
            if i == 9 and calculation_type == "cohen's d":
                cancel = True
            else:
                cancel = False
            select_dict(
                d1=x[i],
                d2=y[i],
                method=method,
                calculation_type=calculation_type,
                cancel=cancel,
                x_total=x[0],
                y_total=y[0]
            )
            received += 1
            progressbar(
                s=size,
                r=received,
                n=now,
                pframe=pframe,
                pbar=pbar,
                plabel=plabel,
                pstring=pstring
            )
        else:
            if (
                i in [8, 10, 11, 12, 13, 14, 15, 22]
                and
                calculation_type == "cohen's d"
            ):
                cancel = True
            else:
                cancel = False
            if i in [13, 14, 15, 22]:
                for key in x[i]:
                    for k in x[i][key]:
                        select_dict(
                            d1=x[i][key][k],
                            d2=y[i][key][k],
                            method=method,
                            calculation_type=calculation_type,
                            cancel=cancel,
                            x_total=x[0],
                            y_total=y[0]
                        )
            else:
                for key in x[i]:
                    select_dict(
                        d1=x[i][key],
                        d2=y[i][key],
                        method=method,
                        calculation_type=calculation_type,
                        cancel=cancel,
                        x_total=x[0],
                        y_total=y[0]
                    )
            received += 1
            progressbar(
                s=size,
                r=received,
                n=now,
                pframe=pframe,
                pbar=pbar,
                plabel=plabel,
                pstring=pstring
            )
    results = [
        x[i] if len(x[i]) == len(y[i]) else {}
        for i in range(1, 23)
    ]
    Spreadsheet(
        filename=output,
        info=x_info,
        planets_in_signs=results[0],
        planets_in_elements=results[1],
        planets_in_modes=results[2],
        houses_in_signs=results[3],
        houses_in_elements=results[4],
        houses_in_modes=results[5],
        planets_in_houses=results[6],
        aspects=results[7],
        sum_of_aspects=results[8],
        yod=results[9],
        t_square=results[10],
        grand_trine=results[11],
        mystic_rectangle=results[12],
        grand_cross=results[13],
        kite=results[14],
        planets_in_houses_in_signs=results[15],
        basic_traditional_rulership=results[16],
        basic_modern_rulership=results[17],
        detailed_traditional_rulership=results[18],
        detailed_modern_rulership=results[19],
        midpoints=results[21]
    )
    received += 1
    progressbar(
        s=size,
        r=received,
        n=now,
        pframe=pframe,
        pbar=pbar,
        plabel=plabel,
        pstring=pstring
    )
    widget.after(
        0, 
        lambda: MsgBox(
            icons=icons,
            title="Info",
            level="info",
            message="Calculation process is completed!"
        )
    )
