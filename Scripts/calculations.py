# -*- coding: utf-8 -*-

from .zodiac import Zodiac
from .spreadsheet import Spreadsheet
from .messagebox import MsgBox, ChoiceBox
from .utilities import convert_coordinates, progressbar
from .modules import (
    os, dt, pd, tk, ttk, time, binom, shutil, Thread, variance, ConfigParser
)
from .constants import (
    HOUSE_SYSTEMS, PLANETS, SIGNS, TRADITIONAL_RULERSHIP, MODERN_RULERSHIP
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


def find_observed_values(widget, icons, menu):
    displayed_results = []
    selected_categories = []
    selected_ratings = []
    checkbuttons = {}
    mode = ""
    for i in widget.winfo_children():
        if hasattr(i, "displayed_results"):
            mode += i.mode
            displayed_results += [
                i.treeview.item(j)["values"]
                for j in i.treeview.get_children()
            ]
            selected_categories += i.selected_categories
            selected_ratings += i.selected_ratings
            checkbuttons.update(i.checkbuttons)
            break
    if not displayed_results:
        return
    save_categories = [i for i in selected_categories]
    if selected_categories:
        if len(selected_categories) > 1:
            selected_categories = "Control_Group"
        else:
            selected_categories = selected_categories[0]\
                .replace(" : ", "/").replace(" ", "_")
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
    config = ConfigParser()
    config.read("defaults.ini")
    if "event" in checkbuttons:
        info = {
            key.title(): "True" if value[0].get() == "0" else "False"
            for key, value in checkbuttons.items()
        }
    else:
        info = {"Event": "False", "Human": "True"}
        info.update(
            {
                key.title(): "True" if value[0].get() == "0" else "False"
                for key, value in checkbuttons.items()
            }
        )
    info.update(
        {
            "Database": config["DATABASE"]["selected"]
            .replace(".json", "").replace(".xml", ""),
            "House System": config["HOUSE SYSTEM"]["selected"],
            "Rodden Rating": selected_ratings,
            "Category": selected_categories
        }
    )
    path = os.path.join(
        *selected_categories.split("/"),
        f"RR_{selected_ratings}",
        f"ORB_{'_'.join(config['ORB FACTORS'].values())}",
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
    if os.path.exists(os.path.join(path, "observed_values.xlsx")):
        msg = "The file already exists.\n Do you want to continue?\n" \
              " If you press 'Yes', \nthe file will be overwritten. \n" \
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
                    mode=mode,
                    menu=menu,
                    save_categories=save_categories
                )
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
                mode=mode,
                menu=menu,
                save_categories=save_categories
            )
        ).start()


def start_calculation(
        displayed_results,
        info,
        config,
        widget,
        icons,
        path,
        selected_categories,
        mode,
        menu,
        save_categories
):
    planets_in_signs = create_normal_dict(
        lists=[PLANETS, SIGNS],
        keys=["", ""]
    )
    houses_in_signs = create_normal_dict(
        lists=[range(1, 13), SIGNS],
        keys=["House-", ""]
    )
    planets_in_houses = create_normal_dict(
        lists=[PLANETS, range(1, 13)],
        keys=["", "House-"]
    )
    planets_in_houses_in_signs = create_normal_dict(
        lists=[PLANETS, range(1, 13), SIGNS],
        keys=["", "House-", ""]
    )
    aspects = create_enumerate_dict(
        lists=[list(config["ORB FACTORS"]), PLANETS, PLANETS],
        keys=["", "", ""],
    )
    total_aspects = create_enumerate_dict(
        lists=[PLANETS, PLANETS],
        keys=["", ""]
    )
    traditional_rulership = create_normal_dict(
        lists=[range(1, 13), TRADITIONAL_RULERSHIP.values(), range(1, 13)],
        keys=["Lord-", "", "House-"]
    )
    modern_rulership = create_normal_dict(
        lists=[range(1, 13), MODERN_RULERSHIP.values(), range(1, 13)],
        keys=["Lord-", "", "House-"]
    )
    total_traditional_rulership = create_normal_dict(
        lists=[range(1, 13), range(1, 13)],
        keys=["Lord-", "House-"]
    )
    total_modern_rulership = create_normal_dict(
        lists=[range(1, 13), range(1, 13)],
        keys=["Lord-", "House-"]
    )
    size = len(displayed_results)
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
            f"Database: {info['Database']}\n"
            f"House System: {info['House System']}\n"
            f"Rodden Rating: {info['Rodden Rating']}\n"
            f"Orb Factor: {'_'.join(config['ORB FACTORS'].values())}\n"
            f"Category: {info['Category']}\n\n"
        )
    else:
        log.write(
            f"Database: {info['Database']}\n"
            f"House System: {info['House System']}\n"
            f"Rodden Rating: {info['Rodden Rating']}\n"
            f"Orb Factor: {'_'.join(config['ORB FACTORS'].values())}\n"
            f"Category: {info['Category']}\n"
            f"Selected Categories:\n"
        )        
        for index, i in enumerate(save_categories, 1):
            log.write(f"\t{index}. {i}\n".expandtabs(20))
        log.write("\n")
    log.write(
        f"|{dt.now().strftime('%Y-%m-%d %H:%M:%S')}| Process started.\n\n"
    )
    log.flush()
    menu.entryconfigure(0, command=lambda: send_warning_message(icons=icons))
    for i in displayed_results:
        if mode == "adb":
            jd = float(i[6])
            lat = convert_coordinates(i[7])
            lon = convert_coordinates(i[8])
        else:
            jd = i[6]
            lat = i[7]
            lon = i[8]
        try:
            patterns = Zodiac(
                jd=jd,
                lat=lat,
                lon=lon,
                hsys=HOUSE_SYSTEMS[config["HOUSE SYSTEM"]["selected"]]
            ).patterns()
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
        for index, p in enumerate(patterns[0], 1):
            planets_in_signs[p[0]][p[1]] += 1
            planets_in_houses[p[0]][p[3]] += 1
            planets_in_houses_in_signs[p[0]][p[3]][p[1]] += 1
            for _p in patterns[0][index:]:
                find_aspect(
                    aspects=aspects,
                    orb=config["ORB FACTORS"],
                    planet1=p[0],
                    planet2=_p[0],
                    aspect=abs(p[2] - _p[2])
                )
        for index, h in enumerate(patterns[1], 1):
            houses_in_signs[f"House-{h[0]}"][h[1]] += 1
            lord_traditional = TRADITIONAL_RULERSHIP[h[1]]
            lord_modern = MODERN_RULERSHIP[h[1]]
            for p in patterns[0]:
                for lord, rulership in [
                        [lord_traditional, traditional_rulership],
                        [lord_modern, modern_rulership]
                ]:
                    select_rulership(
                        lord=lord,
                        rulership=rulership,
                        p=p,
                        index=index
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
    log.write(
        f"|{dt.now().strftime('%Y-%m-%d %H:%M:%S')}| Process finished."
    )
    log.close()
    for aspect in aspects:
        for index, planet in enumerate(PLANETS, 1):
            for _planet in list(PLANETS)[index:]:
                total_aspects[planet][_planet] += \
                    aspects[aspect][planet][_planet]
    get_total_of_rulership(
        constant=TRADITIONAL_RULERSHIP,
        rulership=traditional_rulership,
        total=total_traditional_rulership
    )
    get_total_of_rulership(
        constant=MODERN_RULERSHIP,
        rulership=modern_rulership,
        total=total_modern_rulership
    )
    if not os.path.exists(os.path.join(".", path)):
        os.makedirs(path)
    filename = os.path.join(path, "observed_values.xlsx")
    Spreadsheet(
        filename=filename,
        info=info,
        planets_in_signs=planets_in_signs,
        houses_in_signs=houses_in_signs,
        planets_in_houses=planets_in_houses,
        aspects=aspects,
        total_aspects=total_aspects,
        planets_in_houses_in_signs=planets_in_houses_in_signs,
        total_traditional_rulership=total_traditional_rulership,
        total_modern_rulership=total_modern_rulership,
        traditional_rulership=traditional_rulership,
        modern_rulership=modern_rulership
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
                menu=menu
            )
        ).start()
    )


def select_rulership(lord, rulership, index, p):
    if lord[:-4] == p[0]:
        rulership[
            f"Lord-{index}"
        ][lord][p[3]] += 1


def get_total_of_rulership(constant, rulership, total):
    for i in range(1, 13):
        for key, value in constant.items():
            for j in range(1, 13):
                total[
                    f"Lord-{i}"
                ][f"House-{j}"] += \
                    rulership[f"Lord-{i}"][value][f"House-{j}"]


def find_aspect(aspects, orb, aspect, planet1, planet2):
    if (
            0 < aspect < float(orb["conjunction"])
            or
            360 - float(orb["conjunction"]) < aspect < 360
    ):
        aspects["conjunction"][planet1][planet2] += 1
    elif (
            30 - float(orb["semi-sextile"]) <
            aspect < 30 + float(orb["semi-sextile"])
            or
            330 - float(orb["semi-sextile"]) <
            aspect < 330 + float(orb["semi-sextile"])
    ):
        aspects["semi-sextile"][planet1][planet2] += 1
    elif (
            45 - float(orb["semi-square"]) <
            aspect < 45 + float(orb["semi-square"])
            or
            315 - float(orb["semi-square"]) <
            aspect < 315 + float(orb["semi-square"])
    ):
        aspects["semi-square"][planet1][planet2] += 1
    elif (
            60 - float(orb["sextile"]) <
            aspect < 60 + float(orb["sextile"])
            or
            300 - float(orb["sextile"]) <
            aspect < 300 + float(orb["sextile"])
    ):
        aspects["sextile"][planet1][planet2] += 1
    elif (
            72 - float(orb["sextile"]) <
            aspect < 72 + float(orb["sextile"])
            or
            288 - float(orb["sextile"]) <
            aspect < 288 + float(orb["sextile"])
    ):
        aspects["quintile"][planet1][planet2] += 1
    elif (
            90 - float(orb["square"]) <
            aspect < 90 + float(orb["square"]) or
            270 - float(orb["square"]) <
            aspect < 270 + float(orb["square"])
    ):
        aspects["square"][planet1][planet2] += 1
    elif (
            120 - float(orb["trine"]) <
            aspect < 120 + float(orb["trine"])
            or
            240 - float(orb["trine"]) <
            aspect < 240 + float(orb["trine"])
    ):
        aspects["trine"][planet1][planet2] += 1
    elif (
            135 - float(orb["sesquiquadrate"]) <
            aspect < 135 + float(orb["sesquiquadrate"])
            or
            225 - float(orb["sesquiquadrate"]) <
            aspect < 225 + float(orb["sesquiquadrate"])
    ):
        aspects["sesquiquadrate"][planet1][planet2] += 1
    elif (
            144 - float(orb["biquintile"]) <
            aspect < 144 + float(orb["biquintile"])
            or
            216 - float(orb["biquintile"]) <
            aspect < 216 + float(orb["biquintile"])
    ):
        aspects["biquintile"][planet1][planet2] += 1
    elif (
            150 - float(orb["quincunx"]) <
            aspect < 150 + float(orb["quincunx"])
            or
            210 - float(orb["quincunx"]) <
            aspect < 210 + float(orb["quincunx"])
    ):
        aspects["quincunx"][planet1][planet2] += 1
    elif (
            180 - float(orb["opposite"]) <
            aspect < 180 + float(orb["opposite"])
    ):
        aspects["opposite"][planet1][planet2] += 1


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


def get_values(filename):
    df = pd.read_excel(filename)
    values = df.values
    info = {df.columns[0].replace(":", ""): df.columns[2]}
    info.update({i[0].replace(":", ""): i[2] for i in values[:5]})
    info.update({"Database": df.columns[5]})
    info.update(
        {
            i[3].replace(":", ""): i[5]
            for i in values[:5] if not isinstance(i[3], float)
        }
    )
    config = ConfigParser()
    config.read("defaults.ini")
    HOUSES = [f"House-{i}" for i in range(1, 13)]
    total = sum(values[7][1: 13])
    planets_in_signs = get_basic_dict(values, [7, 19], [PLANETS, SIGNS])
    houses_in_signs = get_basic_dict(values, [21, 33], [HOUSES, SIGNS])
    planets_in_houses = get_basic_dict(values, [35, 47], [PLANETS, HOUSES])
    aspects = {}
    c = 49
    for key in config["ORB FACTORS"]:
        aspects[key] = get_aspect_dict(values, [c, c + 12], [PLANETS])
        c += 14
    total_aspects = get_aspect_dict(values, [203, 215], [PLANETS])
    planets_in_houses_in_signs = {}
    c = 217
    for planet in PLANETS:
        planets_in_houses_in_signs[planet] = get_basic_dict(
            values, [c, c + 12], [HOUSES, SIGNS], [2, 14]
        )
        c += 15
    total_traditional_rulership = get_basic_dict(
        values, [398, 410], [[f"Lord-{i}" for i in range(1, 13)], HOUSES]
    )
    total_modern_rulership = get_basic_dict(
        values, [414, 426], [[f"Lord-{i}" for i in range(1, 13)], HOUSES]
    )
    traditional_rulership = {}
    c = 430
    for i in range(1, 13):
        traditional_rulership[f"Lord-{i}"] = get_basic_dict(
            values,
            [c, c + 12],
            [list(TRADITIONAL_RULERSHIP.values()), HOUSES],
            [2, 14]
        )
        c += 15
    modern_rulership = {}
    c = 611
    for i in range(1, 13):
        modern_rulership[f"Lord-{i}"] = get_basic_dict(
            values,
            [c, c + 12],
            [list(MODERN_RULERSHIP.values()), HOUSES],
            [2, 14]
        )
        c += 15
    return total, planets_in_signs, houses_in_signs, planets_in_houses, \
        aspects, total_aspects, planets_in_houses_in_signs, \
        total_traditional_rulership, total_modern_rulership, \
        traditional_rulership, modern_rulership, info


def probability_mass_function(n, k, p):
    result = 0
    for i in range(k + 1):
        result += binom.pmf(n=n, k=i, p=p) * 100
    return round(result, 6)
                
                
def select_basic(
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
            

def select_detailed(
        d1, 
        d2, 
        method, 
        calculation_type, 
        cancel, 
        x_total, 
        y_total
):
    for (key, value), (_key, _value) in zip(d1.items(), d2.items()):
        for (k, v), (k_, v_) in zip(value.items(), _value.items()):
            save1 = {m: n for m, n in v.items()}
            save2 = {m: n for m, n in v_.items()}
            for (_k, _v), (__k, __v) in zip(v.items(), v_.items()):
                try:
                    if calculation_type == "expected":
                        if method == "Subcategory":
                            d1[key][k][_k] = __v * x_total / y_total
                        elif method == "Independent":
                            d1[key][k][_k] = \
                                x_total * (_v + __v) / (x_total + y_total)
                    elif calculation_type == "effect-size":
                        d1[key][k][_k] = _v / __v
                    elif calculation_type == "chi-square":
                        d1[key][k][_k] = (_v - __v) ** 2 / __v
                    elif calculation_type == "cohen's d":
                        if cancel:
                            d1[key][k][_k] = ""
                        else:
                            d1[key][k][_k] = (_v - __v) / \
                                (
                                    (
                                        variance(save1.values()) +
                                        variance(save2.values())
                                    ) / 2
                                ) ** 0.5
                    elif calculation_type == "binomial limit":
                        p1 = probability_mass_function(
                            n=x_total, k=_v, p=__v / y_total
                        )
                        p2 = 100 - probability_mass_function(
                            n=x_total, k=_v - 1, p=__v / y_total
                        )
                        if p1 < p2:
                            d1[key][k][_k] = -p1
                        elif p1 > p2:
                            d1[key][k][_k] = p2
                except ZeroDivisionError:
                    d1[key][k][_k] = 0
    
    
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
    x = get_values(filename=input1)
    y = get_values(filename=input2)
    x_info = x[-1]
    y_info = y[-1]
    if calculation_type in ["expected", "binomial limit"]:
        for k in x_info:
            x_info[k] = x_info[k] + " & " + y_info[k]
    else:
        x_info = y_info
    config = ConfigParser()
    config.read("defaults.ini")
    method = config["METHOD"]["selected"]
    for i in range(len(y)):
        if i in [0, 11]:
            continue
        if i in [1, 2, 3, 5, 7, 8]:
            if i == 5 and calculation_type == "cohen's d":
                cancel = True
            else:
                cancel = False
            select_basic(
                d1=x[i], 
                d2=y[i],
                method=method,
                calculation_type=calculation_type,
                cancel=cancel,
                x_total=x[0],
                y_total=y[0]
            )
        else:
            if i == 4 and calculation_type == "cohen's d":
                cancel = True
            else:
                cancel = False
            select_detailed(
                d1=x[i], 
                d2=y[i],
                method=method,
                calculation_type=calculation_type,
                cancel=cancel,
                x_total=x[0],
                y_total=y[0]
            )
    Spreadsheet(
        filename=output,
        info=x_info,
        planets_in_signs=x[1],
        houses_in_signs=x[2],
        planets_in_houses=x[3],
        aspects=x[4],
        total_aspects=x[5],
        planets_in_houses_in_signs=x[6],
        total_traditional_rulership=x[7],
        total_modern_rulership=x[8],
        traditional_rulership=x[9],
        modern_rulership=x[10]
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
