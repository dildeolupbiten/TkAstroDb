# -*- coding: utf-8 -*-

from .modules import Workbook, pd
from .constants import (
    SIGNS, HOUSES, SHEETS, PLANETS,
    MODERN_RULERSHIP, TRADITIONAL_RULERSHIP
)
from .utilities import (
    only_planets, get_basic_dict, get_planet_dict,
    get_aspect_dict, get_3d_pattern_dict, get_4d_pattern_dict,
    get_midpoint_dict
)


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
            mystic_rectangle,
            grand_cross,
            kite,
            midpoints,
            *args,
            **kwargs
    ):
        super().__init__(*args, **kwargs)
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
        planets = only_planets(PLANETS)
        for index, name in enumerate(SHEETS):
            try:
                df = pd.read_excel(kwargs["filename"], sheet_name=name)
            except FileNotFoundError:
                df = None
            if name == "Planets In Signs":
                if not planets_in_signs:
                    if df is not None and len(df.values) != 0:
                        planets_in_signs = get_basic_dict(
                            values=df.values,
                            indexes=[0, 13],
                            constants=[planets, SIGNS]
                        )
                if planets_in_signs:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=planets_in_signs,
                        row=1
                    )
                    planets_in_signs.clear()
            elif name == "Planets In Elements":
                if not planets_in_elements:
                    if df is not None and len(df.values) != 0:
                        planets_in_elements = get_basic_dict(
                            values=df.values,
                            indexes=[0, 13],
                            constants=[
                                planets,
                                ["Fire", "Earth", "Air", "Water"]
                            ],
                            sub_index=(1, 5)
                        )
                if planets_in_elements:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=planets_in_elements,
                        row=1
                    )
                    planets_in_elements.clear()
            elif name == "Planets In Modes":
                if not planets_in_modes:
                    if df is not None and len(df.values) != 0:
                        planets_in_modes = get_basic_dict(
                            values=df.values,
                            indexes=[0, 13],
                            constants=[
                                planets,
                                ["Cardinal", "Fixed", "Mutable"]
                            ],
                            sub_index=(1, 4)
                        )
                if planets_in_modes:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=planets_in_modes,
                        row=1
                    )
                    planets_in_modes.clear()
            elif name == "Houses In Signs":
                if not houses_in_signs:
                    if df is not None and len(df.values) != 0:
                        houses_in_signs = get_basic_dict(
                            values=df.values,
                            indexes=[0, 12],
                            constants=[HOUSES, SIGNS]
                        )
                if houses_in_signs:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=houses_in_signs,
                        row=1
                    )
                    houses_in_signs.clear()
            elif name == "Houses In Elements":
                if not houses_in_elements:
                    if df is not None and len(df.values) != 0:
                        houses_in_elements = get_basic_dict(
                            values=df.values,
                            indexes=[0, 12],
                            constants=[
                                HOUSES,
                                ["Fire", "Earth", "Air", "Water"]
                            ],
                            sub_index=(1, 5)
                        )
                if houses_in_elements:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=houses_in_elements,
                        row=1
                    )
                    houses_in_elements.clear()
            elif name == "Houses In Modes":
                if not houses_in_modes:
                    if df is not None and len(df.values) != 0:
                        houses_in_modes = get_basic_dict(
                            values=df.values,
                            indexes=[0, 12],
                            constants=[
                                HOUSES,
                                ["Cardinal", "Fixed", "Mutable"]
                            ],
                            sub_index=(1, 4)
                        )
                if houses_in_modes:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=houses_in_modes,
                        row=1
                    )
                    houses_in_modes.clear()
            elif name == "Planets In Houses":
                if not planets_in_houses:
                    if df is not None and len(df.values) != 0:
                        planets_in_houses = get_basic_dict(
                            values=df.values,
                            indexes=[0, 13],
                            constants=[planets, HOUSES]
                        )
                if planets_in_houses:
                    self.write_basic(
                        sheet=self.sheets[name],
                        data=planets_in_houses,
                        row=1
                    )
                    planets_in_houses.clear()
            elif name == "Planets In Houses In Signs":
                if not planets_in_houses_in_signs:
                    if df is not None and len(df.values) != 0:
                        planets_in_houses_in_signs = get_planet_dict(
                            values=df.values,
                            planets=planets,
                            c=0,
                            arrays=[HOUSES, SIGNS]
                        )
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
                if not d:
                    if df is not None and len(df.values) != 0:
                        d = get_basic_dict(
                            values=df.values,
                            indexes=[0, 12],
                            constants=[
                                [f"Lord-{i}" for i in range(1, 13)],
                                HOUSES
                            ]
                        )
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
                    rulership = list(TRADITIONAL_RULERSHIP.values())
                else:
                    d = detailed_modern_rulership
                    rulership = list(MODERN_RULERSHIP.values())
                if not d:
                    if df is not None and len(df.values) != 0:
                        d = get_planet_dict(
                            df.values,
                            None,
                            0,
                            [rulership, HOUSES]
                        )
                if d:
                    self.write_advanced(
                        sheet=self.sheets[name],
                        data=d,
                        row=1
                    )
                    d.clear()
            elif name == "Aspects":
                if not aspects:
                    if df is not None and len(df.values) != 0:
                        c = 0
                        for key in info["Orb Factor"]:
                            aspects[key] = get_aspect_dict(
                                values=df.values,
                                indexes=[c, c + 15],
                                constants=[PLANETS]
                            )
                            c += 16
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
                if not sum_of_aspects:
                    if df is not None and len(df.values) != 0:
                        sum_of_aspects = get_aspect_dict(
                            values=df.values,
                            indexes=[0, 15],
                            constants=[PLANETS]
                        )
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
                if not d:
                    if df is not None and len(df.values) != 0:
                        d = get_3d_pattern_dict(
                            values=df.values,
                            c=0
                        )
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
            elif name in ["Mystic Rectangle", "Grand Cross", "Kite"]:
                if name == "Mystic Rectangle":
                    d = mystic_rectangle
                    title = "Opposite"
                    apex = ""
                elif name == "Grand Cross":
                    d = grand_cross
                    title = "Opposite"
                    apex = ""
                else:
                    d = kite
                    title = "Opposite"
                    apex = "Apex "
                if not d:
                    if df is not None and len(df.values) != 0:
                        d = get_4d_pattern_dict(
                            values=df.values,
                            c=0
                        )
                if d:
                    row = 1
                    for key, value in d.items():
                        for k, v in value.items():
                            self.write_aspects(
                                sheet=self.sheets[name],
                                data=v,
                                row=row,
                                aspect=name + " ",
                                orb_factor=f"({apex}{key} {title} {k})"
                            )
                            row += 14
                    d.clear()
            elif name == "Midpoints":
                if not midpoints:
                    if df is not None and len(df.values) != 0:
                        midpoints = get_midpoint_dict(df)
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
            for i in range(12):
                sheet.merge_range(
                    f"A{row + 13}:B{row + 13}",
                    "Total",
                    self.format(bold=True, align="center")
                )
            for index, i in enumerate(totals):
                sheet.write(
                    f"{self.cols[2 + index]}{row + 13}",
                    sum(i),
                    self.format(bold=False, align="center")
                )
            row += 15

    def write_info(self, sheet, info):
        row = 1
        factors = {
            "Orb Factor": "Planet To Planet",
            "Midpoint Orb Factor": "Midpoint To Planet"
        }
        key1_row, key2_row = 0, 0
        for key, value in info.items():
            if key in ["Orb Factor", "Midpoint Orb Factor"]:
                new = {"Aspect": factors[key]}
                new.update(value)
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
                    if key == "Midpoint Orb Factor":
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
