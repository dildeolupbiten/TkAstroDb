# -*- coding: utf-8 -*-

from .consts import *
from .zodiac import Zodiac
from .spreadsheet import Spreadsheet, XLFile
from .libs import (
    e, os, dt, pi, path, getcwd, binom, quad, ElementTree, perf_counter, 
    time as now, QFrame, QVBoxLayout, QPushButton, QMessageBox
)


def get_defaults():
    return {
        "ayanamsha": list(AYANAMSHA),
        "house_systems": list(HOUSE_SYSTEMS),
        "planets": list(PLANETS)[:-2],
        "aspects": list(ORB_FACTORS),
        "sheets": SHEETS[1:],
        "categories": None,
        "rodden_ratings": list(RODDEN_RATINGS),
        "ignored_categories": None,
        "ignored_records": list(IGNORED_RECORDS),
        "defaults": {
            "zodiac": "Tropical",
            "house_systems": "Placidus",
            "ayanamsha": "Hindu-Lahiri",
            "orb_factors": ORB_FACTORS,
            "midpoint_orb_factors": MIDPONT_ORB_FACTORS,
            "rodden_ratings": RODDEN_RATINGS,
            "sheets": {k: True for k in SHEETS[1:]},
            "planets": {k: True for k in list(PLANETS)[:-2]},
            "ignored_records": IGNORED_RECORDS,
            "ignored_categories": []
        }
    }


def get_info(info: dict, total: int, version: str, adb_version: str):
    defaults = get_defaults()
    return {
        "Adb Version": adb_version,
        "TkAstroDb Version": version,
        "Zodiac": info["zodiac"] if info["zodiac"] == "Tropical" else info["zodiac"] + " (" + info["ayanamsha"] + ")",
        "House System": info["house_systems"],
        "Rodden Rating": "+".join(info["rodden_ratings"]),
        "Category": info["selected_categories"][0].replace("/ ", "_").replace("/", "_").replace(" : ", "/")
        if len(info["selected_categories"]) == 1 else "Control_Group",
        "Year Range": " - ".join(info["year_range"]) if all(info["year_range"].values()) else "None",
        "Latitude Range": " - ".join(info["latitude_range"]) if all(info["latitude_range"].values()) else "None",
        "Longitude Range": " - ".join(info["longitude_range"]) if all(info["longitude_range"].values()) else "None",
        **{
            i: str(i not in info["ignored_records"])
            for i in [
                "Event",
                "Human",
                "Male",
                "Female",
                "South Hemisphere",
                "North Hemisphere",
                "West Hemisphere",
                "East Hemisphere"
            ]
        },
        "Number Of Records": total,
        "Orb Factor": {j: info["orb_factors"][i] for i, j in enumerate(list(defaults["defaults"]["orb_factors"]))},
        "Midpoint Orb Factor": {
            j: info["midpoint_orb_factors"][i]
            for i, j in enumerate(list(defaults["defaults"]["midpoint_orb_factors"]))
        },
    }


def recursive_numeric_dict(keys: list, repeat: bool = True, index_of_non_repeat: int = 0, ignore: bool = False):
    if not keys:
        return 0
    d = {}
    for index, key in enumerate(keys[0]):
        d[key] = recursive_numeric_dict(
            list(
                map(
                    lambda i: (i if not ignore else tuple(filter(lambda j: j != key, i)))
                    [(0 if (repeat is True or repeat == index_of_non_repeat) else index + (0 if ignore else 1)):],
                    map(tuple, keys[1:])
                )
            ),
            repeat + 1 if repeat is False else True,
            index_of_non_repeat,
            ignore
        )
    return d


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


def find_aspect(orb_factors, aspect):
    if 0 < aspect < orb_factors[0] or 360 - orb_factors[0] < aspect < 360:
        return "Conjunction"
    elif (
        30 - orb_factors[1] < aspect < 30 + orb_factors[1]
        or
        330 - orb_factors[1] < aspect < 330 + orb_factors[1]
    ):
        return "Semi-Sextile"
    elif (
        45 - orb_factors[2] < aspect < 45 + orb_factors[2]
        or
        315 - orb_factors[2] < aspect < 315 + orb_factors[2]
    ):
        return "Semi-Square"
    elif (
        60 - orb_factors[3] < aspect < 60 + orb_factors[3]
        or
        300 - orb_factors[3] < aspect < 300 + orb_factors[3]
    ):
        return "Sextile"
    elif (
        72 - orb_factors[4] < aspect < 72 + orb_factors[4]
        or
        288 - orb_factors[4] < aspect < 288 + orb_factors[4]
    ):
        return "Quintile"
    elif (
        90 - orb_factors[5] < aspect < 90 + orb_factors[5]
        or
        270 - orb_factors[5] < aspect < 270 + orb_factors[5]
    ):
        return "Square"
    elif (
        120 - orb_factors[6] < aspect < 120 + orb_factors[6]
        or
        240 - orb_factors[6] < aspect < 240 + orb_factors[6]
    ):
        return "Trine"
    elif (
        135 - orb_factors[7] < aspect < 135 + orb_factors[7]
        or
        225 - orb_factors[7] < aspect < 225 + orb_factors[7]
    ):
        return "Sesquiquadrate"
    elif (
        144 - orb_factors[8] < aspect < 144 + orb_factors[8]
        or
        216 - orb_factors[8] < aspect < 216 + orb_factors[8]
    ):
        return "BiQuintile"
    elif (
        150 - orb_factors[9] < aspect < 150 + orb_factors[9]
        or
        210 - orb_factors[9] < aspect < 210 + orb_factors[9]
    ):
        return "Quincunx"
    elif 180 - orb_factors[10] < aspect < 180 + orb_factors[10]:
        return "Opposite"


def has_aspect_pattern(planets, patterns):
    if not patterns:
        return True
    diff = abs(planets[0] - planets[1])
    diff = 360 - diff if diff > 180 else diff
    orb = ORB_FACTORS[patterns[0][0]]
    return (
        patterns[0][1] - orb <= diff <= patterns[0][1] + orb
        and
        has_aspect_pattern(
            planets[1:] + ([planets[0]] if len(patterns) == len(planets) else []),
            patterns[1:]
        )
    )


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


def filter_database(database, filters):
    filtered = filters["selected_records"]
    for record in database:
        year = get_year(record[4])
        jd = float(record[6])
        lat = list_to_float(get_coordinates(record[7]))
        lon = list_to_float(get_coordinates(record[8]))
        categories = get_categories(record[-1])
        if record[3] not in filters["rodden_ratings"]:
            continue
        if record[2] == "M" and ("Male" in filters["ignored_records"] or "Human" in filters["ignored_records"]):
            continue
        if record[2] == "F" and ("Female" in filters["ignored_records"] or "Human" in filters["ignored_records"]):
            continue
        if record[2] not in ["M", "F", "MF", "FM", "M/F", "F/M"] and "Event" in filters["ignored_records"]:
            continue
        if lat > 0 and "North Hemisphere" in filters["ignored_records"]:
            continue
        if lat < 0 and "South Hemisphere" in filters["ignored_records"]:
            continue
        if lon > 0 and "East Hemisphere" in filters["ignored_records"]:
            continue
        if lon < 0 and "West Hemisphere" in filters["ignored_records"]:
            continue
        if all(filters["year_range"].values()) and year not in range(*filters["year_range"].values()):
            continue
        if (
                all(filters["latitude_range"].values())
                and
                not (filters["latitude_range"]["From"] <= lat <= filters["latitude_range"]["To"])
        ):
            continue
        if (
                all(filters["longitude_range"].values())
                and
                not (filters["longitude_range"]["From"] <= lat <= filters["longitude_range"]["To"])
        ):
            continue
        if not any(i and i in filters["selected_categories"] for i in categories):
            continue
        if any(i and i in filters["ignored_categories"] for i in categories):
            continue
        data = {"jd": jd, "lat": lat, "lon": lon}
        if data not in filtered:
            filtered.append({"jd": jd, "lat": lat, "lon": lon})
    return filtered


def variance(n: int, p: float) -> float:
    return n * p * (1 - p)


def std_dev(n: int, p: float) -> float:
    return variance(n, p) ** 0.5


def z_score(n: int, k: int, p: float) -> float:
    stdev = std_dev(n, k / n)
    if stdev:
        return (k - (n * p)) / stdev
    else:
        return 0


def erf(x: float):
    return (2 / (pi ** 0.5)) * quad(lambda t: e ** (-t ** 2), a=0, b=x)[0]


def has_significance(n: int, k: int, p: float, alpha: float) -> bool:
    return abs(erf(z_score(n, k, p) / (2 ** 0.5))) > (1 - alpha)


def pmf(n, k, p):
    result = 0
    for i in range(k + 1):
        result += binom.pmf(n=n, k=i, p=p) * 100
    return round(result, 6)


def safe_n(n: int | str):
    if isinstance(n, int):
        return n
    else:
        return int(n.split(" - ")[0])


def safe_cohen(x1, x2, n1, n2):
    divisor = ((variance(x1, x1/safe_n(n1)) + variance(x2, x2/safe_n(n2))) / 2) ** .5
    if divisor:
        return (x1 - x2) / divisor
    else:
        return 0


def select_calculations(calc_type: str, method: str, alpha: float):
    if calc_type == "Calculate Expected Values":
        if method == "Subcategory":
            return lambda x1, x2, n1, n2: x2 * safe_n(n1) / safe_n(n2)
        else:
            return lambda x1, x2, n1, n2: safe_n(n1) * (x1 + x2) / (safe_n(n1) + safe_n(n2))
    elif calc_type == "Calculate Effect-Size Values":
        return lambda x1, x2, n1, n2: (x1 / x2) if x2 else 0
    elif calc_type == "Calculate Chi-Square Values":
        return lambda x1, x2, n1, n2: ((x1 - x2) ** 2 / x2) if x2 else 0
    elif calc_type == "Calculate Cohen's D Values":
        return lambda x1, x2, n1, n2: safe_cohen(x1=x1, x2=x2, n1=n1, n2=n2)
    elif calc_type == "Calculate Binomial Probability Values":
        return lambda x1, x2, n1, n2: (
            -p1 if (
                (p1 := pmf(n=safe_n(n1), k=x1, p=x2/safe_n(n2)))
                <
                (p2 := 100 - pmf(n=safe_n(n1), k=x1 - 1, p=x2 / safe_n(n2))))
            else p2
        )
    elif calc_type == "Calculate Z-Score Values":
        return lambda x1, x2, n1, n2: z_score(n=safe_n(n1), k=x1, p=x2/safe_n(n2))
    elif calc_type == "Calculate Significance Values":
        return lambda x1, x2, n1, n2: has_significance(n=safe_n(n1), k=x1, p=x2/safe_n(n2), alpha=alpha)


def from_xml(self):
    tree = ElementTree.parse(self.filename)
    root = tree.getroot()
    database = []
    category_names = []
    for i in range(2, len(root) - 1):
        user_data = []
        for gender, roddenrating, bdata, adb_link, categories in zip(
            root[i][1].findall("gender"),
            root[i][1].findall("roddenrating"),
            root[i][1].findall("bdata"),
            root[i][2].findall("adb_link"),
            root[i][3].findall("categories")
        ):
            name = root[i][1][0].text
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
            user_data.append(int(root[i].get("adb_id")))
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
    self.database = database
    self.categories = sorted(category_names)
    self.finished.emit()


def create_containers(filters, sheets, planets):
    return {
        sheets[0]: recursive_numeric_dict([planets, SIGNS]) if sheets[0] in filters["sheets"] else None,
        sheets[1]: recursive_numeric_dict([planets, ELEMENTS]) if sheets[1] in filters["sheets"] else None,
        sheets[2]: recursive_numeric_dict([planets, MODES]) if sheets[2] in filters["sheets"] else None,
        sheets[3]: recursive_numeric_dict([HOUSES, SIGNS]) if sheets[3] in filters["sheets"] else None,
        sheets[4]: recursive_numeric_dict([HOUSES, ELEMENTS]) if sheets[4] in filters["sheets"] else None,
        sheets[5]: recursive_numeric_dict([HOUSES, MODES]) if sheets[5] in filters["sheets"] else None,
        sheets[6]: recursive_numeric_dict([planets, HOUSES]) if sheets[6] in filters["sheets"] else None,
        sheets[7]: recursive_numeric_dict([planets, HOUSES, SIGNS]) if sheets[7] in filters["sheets"] else None,
        sheets[8]: recursive_numeric_dict([LORDS, HOUSES]) if sheets[8] in filters["sheets"] else None,
        sheets[9]: recursive_numeric_dict([LORDS, HOUSES]) if sheets[9] in filters["sheets"] else None,
        sheets[10]: recursive_numeric_dict(
            [LORDS, TRADITIONAL_RULERSHIP.values(), HOUSES]) if sheets[10] in filters["sheets"] else None,
        sheets[11]: recursive_numeric_dict(
            [LORDS, MODERN_RULERSHIP.values(), HOUSES]) if sheets[11] in filters["sheets"] else None,
        sheets[12]: recursive_numeric_dict(
            [ORB_FACTORS, PLANETS, PLANETS], repeat=False) if sheets[12] in filters["sheets"] else None,
        sheets[13]: recursive_numeric_dict(
            [PLANETS, PLANETS], repeat=False, index_of_non_repeat=1) if sheets[13] in filters["sheets"] else None,
        sheets[14]: recursive_numeric_dict(
            [PLANETS, PLANETS, PLANETS],
            repeat=False,
            index_of_non_repeat=0, ignore=True) if sheets[14] in filters["sheets"] else None,
        sheets[15]: recursive_numeric_dict(
            [PLANETS, PLANETS, PLANETS],
            repeat=False,
            index_of_non_repeat=0, ignore=True) if sheets[15] in filters["sheets"] else None,
        sheets[16]: recursive_numeric_dict(
            [PLANETS, PLANETS, PLANETS],
            repeat=False,
            index_of_non_repeat=0, ignore=True) if sheets[16] in filters["sheets"] else None,
        sheets[17]: recursive_numeric_dict(
            [PLANETS, PLANETS, ORB_FACTORS, PLANETS],
            repeat=True,
            ignore=True) if sheets[17] in filters["sheets"] else None
    }


def log(msg: str):
    return f"- INFO - {dt.now().strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n"


def get_chart_patterns(filters, records, version, adb_version, self):
    swe.set_ephe_path(path.join(getcwd(), "eph"))
    total = len(records)
    sheets = SHEETS[1:]
    planets = list(PLANETS)[:-2]
    output = log(msg="Analyzing started!")
    filters["orb_factors"] = list(filters["orb_factors"].values())
    filters["midpoint_orb_factors"] = list(filters["midpoint_orb_factors"].values())
    sheet_indexes = [14, 15, 16]
    aspect_patterns = [
        (["Quincunx", 150], ["Sextile", 60], ["Quincunx", 150]),
        (["Square", 90], ["Opposite", 180], ["Square", 90]),
        (["Trine", 120], ["Trine", 120], ["Trine", 120]),
    ]
    containers = {
        "Info": get_info(filters, total, version, adb_version),
        **create_containers(filters, sheets, planets)
    }
    for i, record in enumerate(records):
        try:
            zodiac = Zodiac(
                jd=record["jd"],
                lat=record["lat"],
                lon=record["lon"],
                hsys=HOUSE_SYSTEMS[filters["house_systems"]],
                zodiac=filters["zodiac"],
                ayanamsha=AYANAMSHA[filters["ayanamsha"]] if filters["ayanamsha"] else "Hindu-Lahiri",
                swe=swe
            )
            patterns = zodiac.patterns()
        except BaseException as Error:
            output += log(msg=f"Error: {Error}")
            continue
        pps, hps = patterns[:]
        dpps = {pp[0]: pp for pp in pps}
        traditional, modern = {}, {}
        for index, hp in enumerate(hps):
            traditional[f"Lord-{index + 1}"] = TRADITIONAL_RULERSHIP[hp[1]]
            modern[f"Lord-{index + 1}"] = MODERN_RULERSHIP[hp[1]]
            if containers[sheets[3]]:
                containers[sheets[3]][hp[0]][hp[1]] += 1
            if containers[sheets[4]]:
                containers[sheets[4]][hp[0]][get_element(hp[1])] += 1
            if containers[sheets[5]]:
                containers[sheets[5]][hp[0]][get_mode(hp[1])] += 1
        for index, pp in enumerate(pps[:-2]):
            if containers[sheets[0]]:
                containers[sheets[0]][pp[0]][pp[1]] += 1
            if containers[sheets[1]]:
                containers[sheets[1]][pp[0]][get_element(pp[1])] += 1
            if containers[sheets[2]]:
                containers[sheets[2]][pp[0]][get_mode(pp[1])] += 1
            if containers[sheets[6]]:
                containers[sheets[6]][pp[0]][pp[-1]] += 1
            if containers[sheets[7]]:
                containers[sheets[7]][pp[0]][pp[-1]][pp[1]] += 1
            if containers[sheets[8]]:
                for k, v in traditional.items():
                    if v.split(" ")[0] == pp[0]:
                        containers[sheets[8]][k][pp[-1]] += 1
            if containers[sheets[9]]:
                for k, v in modern.items():
                    if v.split(" ")[0] == pp[0]:
                        containers[sheets[9]][k][pp[-1]] += 1
            if containers[sheets[8]] and containers[sheets[10]]:
                for k, v in traditional.items():
                    if v.split(" ")[0] == pp[0]:
                        containers[sheets[10]][k][v][pp[-1]] += 1
            if containers[sheets[9]] and containers[sheets[11]]:
                for k, v in modern.items():
                    if v.split(" ")[0] == pp[0]:
                        containers[sheets[11]][k][v][pp[-1]] += 1
            if containers[sheets[12]]:
                for other_pp in pps[index + 1:]:
                    aspect = find_aspect(
                        filters["orb_factors"],
                        abs(zodiac.reverse_convert_degree(pp) - zodiac.reverse_convert_degree(other_pp))
                    )
                    if aspect:
                        containers[sheets[12]][aspect][pp[0]][other_pp[0]] += 1
                        if containers[sheets[13]]:
                            containers[sheets[13]][pp[0]][other_pp[0]] += 1
        if containers[sheets[12]]:
            for pp in pps[-2:]:
                for other_pp in pps[-1:]:
                    aspect = find_aspect(
                        filters["orb_factors"],
                        abs(zodiac.reverse_convert_degree(pp) - zodiac.reverse_convert_degree(other_pp))
                    )
                    if aspect:
                        containers[sheets[12]][aspect][pp[0]][other_pp[0]] += 1
                        if containers[sheets[13]]:
                            containers[sheets[13]][pp[0]][other_pp[0]] += 1
        if containers[sheets[14]] or containers[sheets[15]] or containers[sheets[16]]:
            for sheet_index, aspect_pattern in zip(sheet_indexes, aspect_patterns):
                if containers[sheets[sheet_index]]:
                    for planet1, v1 in containers[sheets[sheet_index]].items():
                        for planet2, v2 in v1.items():
                            for planet3, v3 in v2.items():
                                p1, p2, p3 = map(dpps.get, [planet1, planet2, planet3])
                                pp1, pp2, pp3 = map(zodiac.reverse_convert_degree, (p1, p2, p3))
                                if has_aspect_pattern([pp1, pp2, pp3], aspect_pattern):
                                    containers[sheets[sheet_index]][planet1][planet2][planet3] += 1
        if containers[sheets[17]]:
            for p1 in pps:
                for p2 in list(filter(lambda p: p not in [p1], pps)):
                    for p3 in list(filter(lambda p: p not in [p1, p2], pps)):
                        pp1, pp2, pp3 = map(zodiac.reverse_convert_degree, (p1, p2, p3))
                        aspect = find_aspect(filters["midpoint_orb_factors"], abs(find_midpoint(pp1, pp2) - pp3))
                        if aspect:
                            containers[sheets[17]][p1[0]][p2[0]][aspect][p3[0]] += 1
        self.progress.emit(i)
    output += log(msg="Analyzing completed!")
    create_spreadsheet(containers, output, "Observed Values")
    self.finished.emit(total)


def float_to_list(arg: float, units: list[int]):
    return ([int(arg * units[0]) if len(units) > 1 else round(arg * units[0], 4)] +
            float_to_list((arg - int(arg * units[0]) / units[0]), units[1:])) if units else []


def list_to_float(arg: list, units: tuple[int] = (1, 60, 3600)):
    return arg[0] / units[0] + list_to_float(arg[1:], units[1:]) if arg else 0


def get_coordinates(s: str):
    char = "".join(filter(lambda i: i in "nswe", s))
    degree, minute = s.split(char)
    degree = int(degree)
    degree *= 1 if ("n" in s or "e" in s) else -1
    minute, second = (minute, 0) if len(minute) <= 2 else (minute[:2], minute[2:])
    return degree, int(minute), int(second)


def get_year(s: str):
    split = s.split(" ")
    if "Jul" in s and "greg" in s:
        return int(split[-2])
    elif "Jul" not in s and "greg" in s:
        return int(split[2])
    elif "Jul" not in s and "greg" not in s:
        return int(split[-1])
    elif "Jul" in s and "greg" not in s:
        return int(split[2])


def get_categories(record: list):
    return list(map(lambda j: j[1], record))


def create_spreadsheet(containers, output, filename):
    filename = f"{select_filename(filename)}-{int(now())}.xlsx"
    output += log(msg=f"Creating {filename}")
    Spreadsheet(
        filename=filename,
        info=containers["Info"],
        planets_in_signs=containers["Planets In Signs"],
        planets_in_elements=containers["Planets In Elements"],
        planets_in_modes=containers["Planets In Modes"],
        houses_in_signs=containers["Houses In Signs"],
        houses_in_elements=containers["Houses In Elements"],
        houses_in_modes=containers["Houses In Modes"],
        planets_in_houses=containers["Planets In Houses"],
        aspects=containers["Aspects"],
        sum_of_aspects=containers["Sum Of Aspects"],
        planets_in_houses_in_signs=containers["Planets In Houses In Signs"],
        basic_traditional_rulership=containers["Basic Traditional Rulership"],
        basic_modern_rulership=containers["Basic Modern Rulership"],
        detailed_traditional_rulership=containers["Detailed Traditional Rulership"],
        detailed_modern_rulership=containers["Detailed Modern Rulership"],
        yod=containers["Yod"],
        t_square=containers["T-Square"],
        grand_trine=containers["Grand Trine"],
        midpoints=containers["Midpoints"],
    )
    output += log(msg="Completed successfully!")
    with open(filename.replace("xlsx", "log"), "w") as f:
        f.write(output)


def select_filename(key):
    return {
        "Observed Values": "observed_values",
        "Calculate Expected Values": "expected_values",
        "Calculate Chi-Square Values": "chi_square_values",
        "Calculate Effect-Size Values": "effect_size_values",
        "Calculate Cohen's D Values": "cohens_d_values",
        "Calculate Binomial Probability Values": "binomial_probability_values",
        "Calculate Z-Score Values": "z_score_values",
        "Calculate Significance Values": "significance_values"
    }[key]


def calculate(file1, file2, calc_type, method, alpha, self):
    total = len(SHEETS) + 2
    output = log(msg="Reading files!")
    try:
        xl_file = XLFile(
            file1,
            other=XLFile(file2),
            func=select_calculations(calc_type, method, alpha),
        )
    except BaseException as err:
        output += log(msg=f"Format Error: {err}")
        self.progress.emit(total)
        return
    output += log(msg="Completed reading files!")
    data = {}
    for index, sheet in enumerate(SHEETS):
        try:
            data[sheet] = getattr(xl_file, sheet.lower().replace(" ", "_").replace("-", "_"))
        except IndexError:
            data[sheet] = None
            pass
        self.progress.emit(index + 3)
    create_spreadsheet(data, output, calc_type)
    self.finished.emit(total)


def check_update(msgbox):
    try:
        new = urlopen(
            "https://raw.githubusercontent.com/dildeolupbiten"
            "/TkAstroDb/master/README.md"
        ).read().decode()
    except URLError:
        msgbox.info(
            "Warning",
            "Couldn't connect."
        )
        return
    if path.exists("README.md"):
        with open("README.md", "r", encoding="utf-8") as f:
            old = f.read()[:-1]
        if new != old:
            with open("README.md", "w", encoding="utf-8") as f:
                f.write(new)
    try:
        scripts = load(
            urlopen(
                url=f"https://api.github.com/repos/dildeolupbiten/"
                    f"TkAstroDb/contents/src?ref=master"
            )
        )
    except URLError:
        msgbox.info(
            "Warning",
            "Couldn't connect."
        )
        return
    update = False
    for i in scripts:
        try:
            file = urlopen(i["download_url"]).read().decode()
        except URLError:
            msgbox.info(
                "Warning",
                "Couldn't connect."
            )
            return
        if i["name"] not in os.listdir("src"):
            update = True
            with open(f"src/{i['name']}", "w", encoding="utf-8") as f:
                f.write(file)
        else:
            with open(f"src/{i['name']}", "r", encoding="utf-8") as f:
                if file != f.read():
                    update = True
                    with open(
                            f"src/{i['name']}",
                            "w",
                            encoding="utf-8"
                    ) as g:
                        g.write(file)
    if update:
        msgbox.info(
            "Info",
            "Program is updated."
        )
        if os.name == "posix":
            Popen(["python3", "app.py"])
            os.kill(os.getpid(), __import__("signal").SIGKILL)
        elif os.name == "nt":
            Popen(["python", "app.py"])
            os.system(f"TASKKILL /F /PID {os.getpid()}")
    else:
        msgbox.info(
            "Info",
            "Program is up-to-date."
        )


def create_frames(self, frames: dict):
    self.box.frame = {}
    for k, v in frames.items():
        self.box.frame[k] = QFrame(self)
        self.box.frame[k].setObjectName(k)
        if "height" in v:
            self.box.frame[k].setFixedHeight(v["height"])
        if "width" in v:
            self.box.frame[k].setFixedWidth(v["width"])
        self.box.frame[k].layout = v["layout"](self.box.frame[k])
        self.box.addWidget(self.box.frame[k])


def create_widgets(self, labels: list[str]):
    for i, label in enumerate(labels):
        self.widgets += [{
            "button": QPushButton(label, self.box.frame["button"]),
            "frame": QFrame(self.box.frame["frame"]),
        }]
        self.widgets[-1]["layout"] = QVBoxLayout(self.widgets[-1]["frame"])
        self.widgets[-1]["frame"].setVisible(False)
        self.widgets[-1]["frame"].setObjectName(label)
        self.widgets[-1]["button"].clicked.connect(lambda ev, index=i: toggle(self, index))
        self.box.frame["button"].layout.addWidget(self.widgets[-1]["button"])
        self.box.frame["frame"].layout.addWidget(self.widgets[-1]["frame"])


def toggle(self, index):
    style = """
    QPushButton {
        background-color: bg!;
        border: bd!;
        color: white;
        padding: 10px;
        border-radius: 5px;
    }

    QPushButton:hover {
        background-color: #212529;
        border: 1px solid #343a40;
    }
    """
    old_style = style.replace("bg!", "#343a40").replace("bd!", "none")
    new_style = style.replace("bg!", "#212529").replace("bd!", "1px solid #343a40")
    for i in filter(lambda ind: ind != index, range(len(self.widgets))):
        self.widgets[i]["button"].setStyleSheet(old_style)
        self.widgets[i]["frame"].setVisible(False)
    self.widgets[index]["frame"].setVisible(not self.widgets[index]["frame"].isVisible())
    if self.widgets[index]["frame"].isVisible():
        self.widgets[index]["button"].setStyleSheet(new_style)
    else:
        self.widgets[index]["button"].setStyleSheet(old_style)


def get_required_data_parts(records):
    return [
        {
            "jd": float(record[5]),
            "lat": list_to_float(get_coordinates(record[6])),
            "lon": list_to_float(get_coordinates(record[7]))
        } for record in records
    ]


def get_selections(self):
    return {
        "selected_categories": self.select_categories.get(),
        "selected_records": get_required_data_parts(self.select_categories.table.selected),
        "rodden_ratings": self.rodden_ratings.get(),
        "sheets": self.sheets.get(),
        "zodiac": self.zodiac.get(),
        "ayanamsha": self.zodiac.get_other(),
        "house_systems": self.house_systems.get(),
        "orb_factors": self.orb_factors.get(),
        "midpoint_orb_factors": self.midpoint_orb_factors.get(),
        "ignored_categories": self.ignored_categories.get(),
        "ignored_records": self.ignored_records.get(),
        "year_range": self.year_range.get(),
        "latitude_range": self.latitude_range.get(),
        "longitude_range": self.longitude_range.get(),
    }
