# -*- coding: utf-8 -*-

from .libs import np, pd, BytesIO, Workbook
from .consts import SHEETS, PLANETS, ORB_FACTORS


class XLFile:
    def __init__(self, output: BytesIO, other: 'XLFile' = None, func=None):
        super().__init__()
        self.__xl = pd.ExcelFile(output, engine="openpyxl")
        self.other = other
        self.__info = self.__xl.parse("Info")
        self.n1 = self.info["Number Of Records"]
        self.func = (lambda x1, x2: func(x1=x1, x2=x2, n1=self.n1, n2=self.other.n1)) if func else None
        self.__planets_in_signs = self.__xl.parse("Planets In Signs")
        self.__planets_in_elements = self.__xl.parse("Planets In Elements")
        self.__planets_in_modes = self.__xl.parse("Planets In Modes")
        self.__houses_in_signs = self.__xl.parse("Houses In Signs")
        self.__houses_in_elements = self.__xl.parse("Houses In Elements")
        self.__houses_in_modes = self.__xl.parse("Houses In Modes")
        self.__planets_in_houses = self.__xl.parse("Planets In Houses")
        self.__basic_traditional_rulership = self.__xl.parse("Basic Traditional Rulership")
        self.__basic_modern_rulership = self.__xl.parse("Basic Modern Rulership")
        self.__planets_in_houses_in_signs = self.__xl.parse("Planets In Houses In Signs")
        self.__detailed_traditional_rulership = self.__xl.parse("Detailed Traditional Rulership")
        self.__detailed_modern_rulership = self.__xl.parse("Detailed Modern Rulership")
        self.__aspects = self.__xl.parse("Aspects")
        self.__sum_of_aspects = self.__xl.parse("Sum Of Aspects")
        self.__yod = self.__xl.parse("Yod")
        self.__t_square = self.__xl.parse("T-Square")
        self.__grand_trine = self.__xl.parse("Grand Trine")
        self.__midpoints = self.__xl.parse("Midpoints")

    @property
    def info(self):
        return self.__info

    @property
    def planets_in_signs(self):
        return self.__planets_in_signs

    @property
    def planets_in_elements(self):
        return self.__planets_in_elements

    @property
    def planets_in_modes(self):
        return self.__planets_in_modes

    @property
    def houses_in_signs(self):
        return self.__houses_in_signs

    @property
    def houses_in_elements(self):
        return self.__houses_in_elements

    @property
    def houses_in_modes(self):
        return self.__houses_in_modes

    @property
    def planets_in_houses(self):
        return self.__planets_in_houses

    @property
    def basic_traditional_rulership(self):
        return self.__basic_traditional_rulership

    @property
    def basic_modern_rulership(self):
        return self.__basic_modern_rulership

    @property
    def planets_in_houses_in_signs(self):
        return self.__planets_in_houses_in_signs

    @property
    def detailed_traditional_rulership(self):
        return self.__detailed_traditional_rulership

    @property
    def detailed_modern_rulership(self):
        return self.__detailed_modern_rulership

    @property
    def aspects(self):
        return self.__aspects

    @property
    def sum_of_aspects(self):
        return self.__aspects

    @property
    def yod(self):
        return self.__yod

    @property
    def t_square(self):
        return self.__t_square

    @property
    def grand_trine(self):
        return self.__grand_trine

    @property
    def midpoints(self):
        return self.__midpoints

    @staticmethod
    def clean(x):
        return (x.replace("Cusp ", "House-")
                .replace("Lord ", "Lord-")
                .replace("House ", "House-")
                .replace("Biquintile", "BiQuintile"))

    @staticmethod
    def safe_v(v):
        if isinstance(v, float) or isinstance(v, int):
            return True
        elif isinstance(v, str):
            if " - " in v:
                return False
        return True

    def __two_dimensional_1(self, attr, other_attr=None):
        if other_attr is None:
            return {
                cleaned: {self.clean(y): attr.values[r, c] for c, y in enumerate(attr.columns.values[1:-1], 1)}
                for r, x in enumerate(attr.values[0:, 0])
                if (cleaned := self.clean(x)) != "Total"
            }
        else:
            return {
                cleaned: {
                    self.clean(y): self.func(x1=attr.values[r, c], x2=other_attr.values[r, c])
                    for c, y in enumerate(attr.columns.values[1:-1], 1)}
                for r, x in enumerate(attr.values[0:, 0])
                if (cleaned := self.clean(x)) != "Total"
            }

    def __two_dimensional_2(self, planets, array, d, other_array=None):
        planet = None
        index = 0
        count = 0
        if other_array is None:
            for j in array:
                if isinstance(j, str):
                    planet = j
                    count = 0
                    index += 1
                    d[planet] = {}
                else:
                    d[planet][planets[index + count]] = j
                    count += 1
        else:
            for j, k in zip(array, other_array):
                if isinstance(j, str):
                    planet = j
                    count = 0
                    index += 1
                    d[planet] = {}
                else:
                    d[planet][planets[index + count]] = self.func(x1=j, x2=k)
                    count += 1

    def __two_dimensional_3(self, attr, other_attr=None):
        values = attr.values.transpose()
        values = values[~pd.isna(values)]
        data = {}
        planets = list(PLANETS)
        if other_attr is None:
            self.__two_dimensional_2(planets, values, data)
        else:
            other_values = other_attr.values.transpose()
            other_values = other_values[~pd.isna(other_values)]
            self.__two_dimensional_2(planets, values, data, other_values)

        return data

    def __three_dimensional_1(self, attr, other_attr=None):
        data = {}
        planet = None
        if not self.other and not self.func:
            for row in attr.values:
                if row[0] == "Total":
                    continue
                if pd.isna(row[0]):
                    if pd.isna(row[1]):
                        continue
                    data[self.clean(planet)].update({
                        row[1]: {self.clean(s): v for s, v in zip(attr.columns[2:-1], row[2:-1])}
                    })
                elif "\n" in row[0]:
                    planet = row[0].split("\n")[0]
                    data[self.clean(planet)] = {
                        row[1]: {self.clean(s): v for s, v in zip(attr.columns[2:-1], row[2:-1])}
                    }
        else:
            for row1, row2 in zip(attr.values, other_attr.values):
                if row1[0] == "Total":
                    continue
                if pd.isna(row1[0]):
                    if pd.isna(row1[1]):
                        continue
                    data[self.clean(planet)].update({row1[1]: {
                        self.clean(s): self.func(x1=v1, x2=v2)
                        for s, v1, v2 in zip(attr.columns[2:-1], row1[2:-1], row2[2:-1])}
                    })
                elif "\n" in row1[0]:
                    planet = row1[0].split("\n")[0]
                    data[self.clean(planet)] = {row1[1]: {
                        self.clean(s): self.func(x1=v1, x2=v2)
                        for s, v1, v2 in zip(attr.columns[2:-1], row1[2:-1], row2[2:-1])}
                    }
        return data

    def __three_dimensional_2(self, attr, other_attr=None):
        values = np.vstack([attr.columns, attr.values])
        data = {}
        planets = list(PLANETS)
        if other_attr is None:
            for i in range(0, len(values), 16):
                aspect = self.clean(values[i][0])
                data[aspect] = {}
                sub = values[i + 1: i + 16].transpose()
                sub = sub[~pd.isna(sub)]
                self.__two_dimensional_2(planets, sub, data[aspect])
        else:
            other_values = np.vstack([other_attr.columns, other_attr.values])
            for i in range(0, len(values), 16):
                aspect = self.clean(values[i][0])
                data[aspect] = {}
                sub1 = values[i + 1: i + 16].transpose()
                sub1 = sub1[~pd.isna(sub1)]
                sub2 = other_values[i + 1: i + 16].transpose()
                sub2 = sub2[~pd.isna(sub2)]
                self.__two_dimensional_2(planets, sub1, data[aspect], sub2)
        return data

    def __three_dimensional_3(self, attr, other_attr=None):
        values = np.vstack([attr.columns, attr.values])
        planets = list(PLANETS)
        data = {}
        if other_attr is None:
            for i in range(0, len(values), 15):
                sub = values[i + 1: i + 15].transpose()
                sub = sub[~pd.isna(sub)]
                apex = str(values[i][0]).split(": ")[1][:-1]
                data[apex] = {}
                self.__two_dimensional_2(planets, sub, data[apex])
        else:
            other_values = np.vstack([other_attr.columns, other_attr.values])
            for i in range(0, len(values), 15):
                sub1 = values[i + 1: i + 15].transpose()
                sub1 = sub1[~pd.isna(sub1)]
                sub2 = other_values[i + 1: i + 15].transpose()
                sub2 = sub2[~pd.isna(sub2)]
                apex = str(values[i][0]).split(": ")[1][:-1]
                data[apex] = {}
                self.__two_dimensional_2(planets, sub1, data[apex], sub2)
        return data

    def __four_dimensional(self, attr, other_attr=None):
        values = np.vstack([attr.columns, attr.values])
        data = {}
        if other_attr is None:
            for i in range(0, len(values), 13):
                sub = values[i + 1: i + 13]
                planets = values[i][3:]
                for j in np.append(sub[:-1, :1], sub[:-1, 3:], axis=1):
                    pp, aspect = map(lambda m: m[:-1], j[0].split("("))
                    p1, p2 = pp.split(" / ")
                    if not data.get(p1):
                        data[p1] = {p2: {aspect: {}}}
                    else:
                        if p2 not in data[p1]:
                            data[p1][p2] = {aspect: {}}
                        else:
                            if aspect not in data[p1][p2]:
                                data[p1][p2][aspect] = {}
                    for col, k in enumerate(j[1:]):
                        data[p1][p2][aspect][planets[col]] = k
        else:
            other_values = np.vstack([other_attr.columns, other_attr.values])
            for i in range(0, len(values), 13):
                sub1 = values[i + 1: i + 13]
                sub2 = other_values[i + 1: i + 13]
                planets = values[i][3:]
                for j, n in zip(
                        np.append(sub1[:-1, :1], sub1[:-1, 3:], axis=1),
                        np.append(sub2[:-1, :1], sub2[:-1, 3:], axis=1)
                ):
                    pp, aspect = map(lambda m: m[:-1], j[0].split("("))
                    p1, p2 = pp.split(" / ")
                    if not data.get(p1):
                        data[p1] = {p2: {aspect: {}}}
                    else:
                        if p2 not in data[p1]:
                            data[p1][p2] = {aspect: {}}
                        else:
                            if aspect not in data[p1][p2]:
                                data[p1][p2][aspect] = {}
                    for col, (k, o) in enumerate(zip(j[1:], n[1:])):
                        data[p1][p2][aspect][planets[col]] = self.func(x1=k, x2=o)
        return data

    @info.getter
    def info(self):
        if not self.other:
            raw = np.vstack([self.__info.columns.values, self.__info.values])
            cleaned = {k: v for k, v in raw[:18, [0, 2]]}
            cleaned["Orb Factor"] = {k: v for k, v in raw[20:, [0, 2]]}
            cleaned["Midpoint Orb Factor"] = {k: v for k, v in raw[20:, [0, 4]]}
        else:
            raw1 = np.vstack([self.__info.columns.values, self.__info.values])
            raw2 = np.vstack([self.other.__info.columns.values, self.other.__info.values])
            cleaned = {
                k1: (f"{v1} - {v2}" if self.safe_v(v2) else v2)
                for (k1, v1), (k2, v2) in zip(raw1[:18, [0, 2]], raw2[:18, [0, 2]])
            }
            cleaned["Orb Factor"] = {
                k1: (f"{v1} - {v2}" if self.safe_v(v2) else v2)
                for (k1, v1), (k2, v2) in zip(raw1[20:, [0, 2]], raw2[20:, [0, 2]])
            }
            cleaned["Midpoint Orb Factor"] = {
                k1: (f"{v1} - {v2}" if self.safe_v(v2) else v2)
                for (k1, v1), (k2, v2) in zip(raw1[20:, [0, 4]], raw2[20:, [0, 4]])
            }
        if pd.isna(cleaned["Year Range"]):
            cleaned["Year Range"] = "None"
        if pd.isna(cleaned["Latitude Range"]):
            cleaned["Latitude Range"] = "None"
        if pd.isna(cleaned["Longitude Range"]):
            cleaned["Longitude Range"] = "None"
        return cleaned

    @planets_in_signs.getter
    def planets_in_signs(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__planets_in_signs)
        else:
            return self.__two_dimensional_1(self.__planets_in_signs, self.other.__planets_in_signs)

    @planets_in_elements.getter
    def planets_in_elements(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__planets_in_elements)
        else:
            return self.__two_dimensional_1(self.__planets_in_elements, self.other.__planets_in_elements)

    @planets_in_modes.getter
    def planets_in_modes(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__planets_in_modes)
        else:
            return self.__two_dimensional_1(self.__planets_in_modes, self.other.__planets_in_modes)

    @houses_in_signs.getter
    def houses_in_signs(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__houses_in_signs)
        else:
            return self.__two_dimensional_1(self.__houses_in_signs, self.other.__houses_in_signs)

    @houses_in_elements.getter
    def houses_in_elements(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__houses_in_elements)
        else:
            return self.__two_dimensional_1(self.__houses_in_elements, self.other.__houses_in_elements)

    @houses_in_modes.getter
    def houses_in_modes(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__houses_in_modes)
        else:
            return self.__two_dimensional_1(self.__houses_in_modes, self.other.__houses_in_modes)

    @planets_in_houses.getter
    def planets_in_houses(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__planets_in_houses)
        else:
            return self.__two_dimensional_1(self.__planets_in_houses, self.other.__planets_in_houses)

    @basic_traditional_rulership.getter
    def basic_traditional_rulership(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__basic_traditional_rulership)
        else:
            return self.__two_dimensional_1(
                self.__basic_traditional_rulership,
                self.other.__basic_traditional_rulership
            )

    @basic_modern_rulership.getter
    def basic_modern_rulership(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_1(self.__basic_modern_rulership)
        else:
            return self.__two_dimensional_1(self.__basic_modern_rulership, self.other.__basic_modern_rulership)

    @planets_in_houses_in_signs.getter
    def planets_in_houses_in_signs(self):
        if not all([self.other, self.func]):
            return self.__three_dimensional_1(self.__planets_in_houses_in_signs)
        else:
            return self.__three_dimensional_1(
                self.__planets_in_houses_in_signs,
                self.other.__planets_in_houses_in_signs
            )

    @detailed_traditional_rulership.getter
    def detailed_traditional_rulership(self):
        if not all([self.other, self.func]):
            return self.__three_dimensional_1(self.__detailed_traditional_rulership)
        else:
            return self.__three_dimensional_1(
                self.__detailed_traditional_rulership,
                self.other.__detailed_traditional_rulership
            )

    @detailed_modern_rulership.getter
    def detailed_modern_rulership(self):
        if not all([self.other, self.func]):
            return self.__three_dimensional_1(self.__detailed_modern_rulership)
        else:
            return self.__three_dimensional_1(
                self.__detailed_modern_rulership,
                self.other.__detailed_modern_rulership
            )

    @aspects.getter
    def aspects(self):
        if not all([self.other, self.func]):
            return self.__three_dimensional_2(self.__aspects)
        else:
            return self.__three_dimensional_2(self.__aspects, self.other.__aspects)

    @sum_of_aspects.getter
    def sum_of_aspects(self):
        if not all([self.other, self.func]):
            return self.__two_dimensional_3(self.__sum_of_aspects)
        else:
            return self.__two_dimensional_3(self.__sum_of_aspects, self.other.__sum_of_aspects)

    @yod.getter
    def yod(self):
        if not all([self.other, self.func]):
            return self.__three_dimensional_3(self.__yod)
        else:
            return self.__three_dimensional_3(self.__yod, self.other.__yod)

    @t_square.getter
    def t_square(self):
        if not all([self.other, self.func]):
            return self.__three_dimensional_3(self.__t_square)
        else:
            return self.__three_dimensional_3(self.__t_square, self.other.__t_square)

    @grand_trine.getter
    def grand_trine(self):
        if not all([self.other, self.func]):
            return self.__three_dimensional_3(self.__grand_trine)
        else:
            return self.__three_dimensional_3(self.__grand_trine, self.other.__grand_trine)

    @midpoints.getter
    def midpoints(self):
        if not all([self.other, self.func]):
            return self.__four_dimensional(self.__midpoints)
        else:
            return self.__four_dimensional(self.__midpoints, self.other.__midpoints)


class Spreadsheet(Workbook):
    def __init__(
            self,
            info,
            planets_in_signs,
            planets_in_elements,
            planets_in_modes,
            houses_in_signs,
            houses_in_elements,
            houses_in_modes,
            planets_in_houses,
            aspects,
            sum_of_aspects,
            planets_in_houses_in_signs,
            basic_traditional_rulership,
            basic_modern_rulership,
            detailed_traditional_rulership,
            detailed_modern_rulership,
            yod,
            t_square,
            grand_trine,
            midpoints,
            **kwargs
    ):
        super().__init__(filename=kwargs["filename"])
        self.sheets = {
            name: self.add_worksheet(name=name)
            for name in SHEETS
        }
        self.cols = [
            "A", "B", "C",
            "D", "E", "F",
            "G", "H", "I",
            "J", "K", "L",
            "M", "N", "O",
            "P", "Q", "R"
        ]
        if info:
            self.write_info(
                sheet=self.sheets["Info"],
                info=info
            )
        for index, name in enumerate(SHEETS):
            if name == "Planets In Signs":
                if planets_in_signs:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=planets_in_signs,
                        row=1
                    )
                    planets_in_signs.clear()
            elif name == "Planets In Elements":
                if planets_in_elements:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=planets_in_elements,
                        row=1
                    )
                    planets_in_elements.clear()
            elif name == "Planets In Modes":
                if planets_in_modes:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=planets_in_modes,
                        row=1
                    )
                    planets_in_modes.clear()
            elif name == "Houses In Signs":
                if houses_in_signs:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=houses_in_signs,
                        row=1
                    )
                    houses_in_signs.clear()
            elif name == "Houses In Elements":
                if houses_in_elements:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=houses_in_elements,
                        row=1
                    )
                    houses_in_elements.clear()
            elif name == "Houses In Modes":
                if houses_in_modes:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=houses_in_modes,
                        row=1
                    )
                    houses_in_modes.clear()
            elif name == "Planets In Houses":
                if planets_in_houses:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=planets_in_houses,
                        row=1
                    )
                    planets_in_houses.clear()
            elif name == "Planets In Houses In Signs":
                if planets_in_houses_in_signs:
                    self.write_advanced(
                        sheet=self.sheets[name],
                        data=planets_in_houses_in_signs,
                        row=1
                    )
                    planets_in_houses_in_signs.clear()
            elif name in [
                "Basic Traditional Rulership",
                "Basic Modern Rulership"
            ]:
                if name == "Basic Traditional Rulership":
                    d = basic_traditional_rulership
                else:
                    d = basic_modern_rulership
                if d:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=d,
                        row=1,
                        total=True
                    )
                    d.clear()
            elif name in [
                "Detailed Traditional Rulership",
                "Detailed Modern Rulership"
            ]:
                if name == "Detailed Traditional Rulership":
                    d = detailed_traditional_rulership
                else:
                    d = detailed_modern_rulership
                if d:
                    self.write_advanced(
                        sheet=self.sheets[name],
                        data=d,
                        row=1
                    )
                    d.clear()
            elif name == "Aspects":
                if aspects:
                    row = 1
                    for aspect in aspects:
                        self.write_aspects(
                            sheet=self.sheets[name],
                            data=aspects[aspect],
                            row=row,
                            aspect=aspect.title(),
                            orb_factor=""
                        )
                        row += 16
                    aspects.clear()
            elif name == "Sum Of Aspects":
                if sum_of_aspects:
                    self.write_aspects(
                        sheet=self.sheets[name],
                        data=sum_of_aspects,
                        row=1,
                        aspect=name,
                        orb_factor=""
                    )
                    sum_of_aspects.clear()
            elif name in ["Yod", "T-Square", "Grand Trine"]:
                if name == "Yod":
                    d = yod
                    title = "Apex"
                elif name == "T-Square":
                    d = t_square
                    title = "Apex"
                else:
                    d = grand_trine
                    title = "Planet"
                if d:
                    row = 1
                    for k, v in d.items():
                        self.write_aspects(
                            sheet=self.sheets[name],
                            data=v,
                            row=row,
                            aspect=name + " ",
                            orb_factor=f"({title}: {k})"
                        )
                        row += 15
                    d.clear()
            elif name == "Midpoints":
                if midpoints:
                    self.write_midpoints(
                        sheet=self.sheets[name],
                        midpoints=midpoints
                    )
                    midpoints.clear()
        self.close()

    def format(
            self,
            bold=False,
            align="center",
            font_name="Arial",
            font_size=10,
            font_color="black"
    ):
        return self.add_format(
            {
                "bold": bold,
                "align": align,
                "valign": "vcenter",
                "font_name": font_name,
                "font_size": font_size,
                "font_color": font_color
            }
        )

    def write_basic(self, sheet, data, row, total=False):
        totals = [[] for _ in range(13)]
        pos = self.cols[len(data[list(data)[0]]) + 1]
        sheet.write(
            f"{pos}{row}", "Total", self.format(align="center", bold=True)
        )
        for keys, values in data.items():
            if keys.startswith("House"):
                keys = keys.replace("House-", "Cusp ")
            sheet.write(
                f"A{row + 1}",
                keys.replace("-", " "),
                self.format(align="center", bold=True)
            )
            t = 0
            for column, (key, value) in enumerate(values.items()):
                if row == 1:
                    sheet.write(
                        f"{self.cols[1 + column]}{row}",
                        key,
                        self.format(align="center", bold=True)
                    )
                t += value
                sheet.write(
                    f"{self.cols[1 + column]}{row + 1}",
                    value,
                    self.format(align="center", bold=False)
                )
                if total:
                    totals[column].append(value)
            sheet.write(
                f"{pos}{row + 1}",
                t,
                self.format(align="center", bold=False)
            )
            totals[12].append(t)
            row += 1
        if total:
            row += 1
            sheet.write(
                f"A{row}", "Total", self.format(align="center", bold=True)
            )
            for index, i in enumerate(totals):
                sheet.write(
                    f"{self.cols[1 + index]}{row}",
                    sum(i),
                    self.format(align="center", bold=False)
                )

    def write_aspects(self, sheet, data, row, aspect, orb_factor):
        sheet.merge_range(
            f"A{row}:C{row}",
            aspect.title() + orb_factor,
            self.format(align="left", bold=True)
        )
        for index, (key, value) in enumerate(data.items()):
            sheet.write(
                f"{self.cols[index]}{row + 1}",
                key,
                self.format(align="center", bold=True)
            )
            r = 0
            for k, v in value.items():
                sheet.write(
                    f"{self.cols[index]}{row + 2 + r}",
                    v,
                    self.format(align="center", bold=False)
                )
                r += 1
            row += 1

    def write_advanced(self, sheet, data, row):
        for key, value in data.items():
            if "Lord" in key:
                word = "is"
            else:
                word = "in"
            sheet.merge_range(
                f"A{row + 1}:A{row + 12}",
                f"{key.replace('-', ' ')}\n{word}",
                self.format(bold=True, align="center")
            )
            sheet.merge_range(
                f"A{row + 13}:B{row + 13}",
                "Total",
                self.format(bold=True, align="center")
            )
            r = 0
            totals = [[] for _ in range(13)]
            for k, v in value.items():
                if r == 0:
                    col = 0
                    for _k, _v in v.items():
                        sheet.write(
                            f"{self.cols[2 + col]}{row}",
                            _k.replace("-", " "),
                            self.format(bold=True, align="center")
                        )
                        col += 1
                    sheet.write(
                        f"{self.cols[2 + col]}{row}",
                        "Total",
                        self.format(bold=True, align="center")
                    )
                col = 0
                total = 0
                for _key, _value in v.items():
                    sheet.write(
                        f"{self.cols[2 + col]}{row + 1 + r}",
                        _value,
                        self.format(bold=False, align="center")
                    )
                    total += _value
                    totals[col].append(_value)
                    col += 1
                sheet.write(
                    f"{self.cols[2 + col]}{row + 1 + r}",
                    total,
                    self.format(bold=False, align="center")
                )
                totals[col].append(total)
                sheet.write(
                    f"B{row + 1 + r}",
                    k,
                    self.format(bold=True, align="center")
                )
                r += 1
            for index, i in enumerate(totals):
                sheet.write(
                    f"{self.cols[2 + index]}{row + 13}",
                    sum(i),
                    self.format(bold=False, align="center")
                )
            row += 15

    def write_info(self, sheet, info):
        info_order = [
            "Adb Version",
            "TkAstroDb Version",
            "Zodiac",
            "House System",
            "Rodden Rating",
            "Category",
            "Year Range",
            "Latitude Range",
            "Longitude Range",
            "Event",
            "Human",
            "Male",
            "Female",
            "North Hemisphere",
            "South Hemisphere",
            "West Hemisphere",
            "East Hemisphere",
            "Number Of Records",
            "Orb Factor",
            "Midpoint Orb Factor"
        ]
        info = {k: info[k] for k in info_order}
        info["Orb Factor"] = {k: info["Orb Factor"][k] for k in ORB_FACTORS}
        info["Midpoint Orb Factor"] = {k: info["Midpoint Orb Factor"][k] for k in ORB_FACTORS}
        row = 1
        factors = {
            "Orb Factor": "Planet To Planet",
            "Midpoint Orb Factor": "Midpoint To Planet"
        }
        key1_row, key2_row = 0, 0
        for key, value in info.items():
            if key in ["Orb Factor", "Midpoint Orb Factor"]:
                new = {"Aspect": factors[key]}
                new.update(info[key])
                if key1_row == 0:
                    sheet.merge_range(
                        f"A{row}:F{row}",
                        "Orb Factors",
                        self.format(bold=True, align="center")
                    )
                    row += 1
                    key1_row, key2_row = row, row
                for index, (k, v) in enumerate(new.items()):
                    if key == "Orb Factor":
                        bold = True if index == 0 else False
                        align = "center" if index == 0 else "left"
                        sheet.merge_range(
                            f"A{key1_row}:B{key1_row}",
                            k,
                            self.format(bold=bold, align=align)
                        )
                        sheet.merge_range(
                            f"C{key1_row}:D{key1_row}",
                            v,
                            self.format(bold=bold, align=align)
                        )
                        key1_row += 1
                    elif key == "Midpoint Orb Factor":
                        bold = True if index == 0 else False
                        align = "center" if index == 0 else "left"
                        sheet.merge_range(
                            f"E{key2_row}:F{key2_row}",
                            v,
                            self.format(bold=bold, align=align)
                        )
                        key2_row += 1
            else:
                sheet.merge_range(
                    f"A{row}:B{row}",
                    key,
                    self.format(bold=True, align="left")
                )
                sheet.merge_range(
                    f"C{row}:N{row}",
                    value,
                    self.format(bold=False, align="left")
                )
                row += 1

    def write_midpoints(self, sheet, midpoints):
        p1_row = 1
        row1 = 0
        for planet1 in midpoints:
            p2_row = 0
            for planet2 in midpoints[planet1]:
                aspect_row = 0
                totals = [0 for _ in range(len(midpoints) - 2)]
                for aspect in midpoints[planet1][planet2]:
                    row1 = p1_row + p2_row + aspect_row
                    sheet.merge_range(
                        f"A{row1 + 1}:C{row1 + 1}",
                        f"{planet1} / {planet2} ({aspect})",
                        self.format(bold=True, align="left")
                    )
                    c = 3
                    for index, (planet3, value) in enumerate(
                        midpoints[planet1][planet2][aspect].items()
                    ):
                        if aspect_row == 0:
                            sheet.write(
                                f"{self.cols[c]}{p1_row + p2_row}",
                                planet3,
                                self.format(bold=True, align="center")
                            )
                        sheet.write(
                            f"{self.cols[c]}{row1 + 1}",
                            value,
                            self.format(bold=False, align="center")
                        )
                        if value:
                            totals[index] += value
                        c += 1
                    aspect_row += 1
                sheet.merge_range(
                    f"A{row1 + 2}:C{row1 + 2} (Total)",
                    f"{planet1} / {planet2} (Total)",
                    self.format(bold=True, align="left")
                )
                if totals[0]:
                    for index, total in enumerate(totals, 3):
                        sheet.write(
                            f"{self.cols[index]}{row1 + 2}",
                            total,
                            self.format(bold=False, align="center")
                        )
                p2_row += aspect_row + 2
            p1_row += p2_row
