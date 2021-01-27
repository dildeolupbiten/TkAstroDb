# -*- coding: utf-8 -*-

from .modules import Workbook, ConfigParser


class Spreadsheet(Workbook):
    def __init__(
            self,
            info,
            planets_in_signs,
            houses_in_signs,
            planets_in_houses,
            aspects,
            total_aspects,
            planets_in_houses_in_signs,
            total_traditional_rulership,
            total_modern_rulership,
            traditional_rulership,
            modern_rulership,
            *args,
            **kwargs
    ):
        super().__init__(*args, **kwargs)
        self.sheet = self.add_worksheet()
        self.cols = [
            "A", "B", "C",
            "D", "E", "F",
            "G", "H", "I",
            "J", "K", "L",
            "M", "N", "O",
            "P", "Q", "R"
        ]
        self.config = ConfigParser()
        self.config.read("defaults.ini")
        self.orb_factor = self.config["ORB FACTORS"]
        if info:
            self.write_info(info=info)
        row = 8
        for i in [planets_in_signs, houses_in_signs, planets_in_houses]:
            self.write_basic(data=i, row=row)
            row += 14
        row = 50
        for aspect in aspects:
            self.write_aspects(
                data=aspects[aspect],
                row=row,
                aspect=aspect.title(),
                orb_factor=self.config["ORB FACTORS"][aspect]
            )
            row += 14
        self.write_aspects(
            data=total_aspects,
            row=row,
            aspect="All Aspects",
            orb_factor=""
        )
        self.write_advanced(data=planets_in_houses_in_signs, row=218)
        row = 398
        for i, j in zip(
                ["Traditional", "Modern"],
                [total_traditional_rulership, total_modern_rulership]
        ):
            self.write_basic(
                data=j,
                row=row,
                title=f"{i} House Rulership"
            )
            row += 16
        row = 430
        for i, j in zip(
                ["Traditional", "Modern"],
                [traditional_rulership, modern_rulership]
        ):
            self.write_advanced(
                data=j,
                row=row,
                title=f"Detailed {i} House Rulership"
            )
            row += 181
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

    def write_basic(self, data, row, title=""):
        title_row = row
        totals = [[] for _ in range(13)]
        if title:
            title_row = row + 1
            self.sheet.merge_range(
                f"A{row}:N{row}",
                title,
                self.format(align="center", bold=True)
            )
            self.sheet.write(
                f"A{row + 14}",
                "Total",
                self.format(align="center", bold=True)
            )
            row += 1
        self.sheet.write(
            f"N{row}", "Total", self.format(align="center", bold=True)
        )
        for keys, values in data.items():
            if keys.startswith("House"):
                keys = keys.replace("House-", "Cusp ")
            self.sheet.write(
                f"A{row + 1}",
                keys.replace("-", " "),
                self.format(align="center", bold=True)
            )
            total = 0
            for column, (key, value) in enumerate(values.items()):
                total += value
                self.sheet.write(
                    f"{self.cols[1 + column]}{row + 1}",
                    value,
                    self.format(align="center", bold=False)
                )
                if title:
                    totals[column].append(value)
                if row == title_row:
                    if "-" in key:
                        key = key.replace("-", " ")
                    self.sheet.write(
                        f"{self.cols[1 + column]}{row}",
                        key,
                        self.format(align="center", bold=True)
                    )
            self.sheet.write(
                f"N{row + 1}",
                total,
                self.format(align="center", bold=False)
            )
            totals[12].append(total)
            row += 1
        if title:
            for index, i in enumerate(totals):
                self.sheet.write(
                    f"{self.cols[1 + index]}{row + 1}",
                    sum(i),
                    self.format(align="center", bold=False)
                )

    def write_aspects(self, data, row, aspect, orb_factor):
        if orb_factor:
            orb = ": Orb Factor: +- " + orb_factor
        else:
            orb = ""
        self.sheet.merge_range(
            f"A{row}:C{row}",
            aspect.title() + orb,
            self.format(align="left", bold=True)
        )
        for index, (key, value) in enumerate(data.items()):
            self.sheet.write(
                f"{self.cols[index]}{row + 1}",
                key,
                self.format(align="center", bold=True)
            )
            r = 0
            for k, v in value.items():
                self.sheet.write(
                    f"{self.cols[index]}{row + 2 + r}",
                    v,
                    self.format(align="center", bold=False)
                )
                r += 1
            row += 1

    def write_advanced(self, data, row, title=""):
        if title:
            self.sheet.merge_range(
                f"A{row}:O{row}",
                title,
                self.format(align="center", bold=True)
            )
            row += 1
        for key, value in data.items():
            if "Lord" in key:
                word = "is"
            else:
                word = "in"
            self.sheet.merge_range(
                f"A{row + 1}:A{row + 12}",
                f"{key.replace('-', ' ')}\n{word}",
                self.format(bold=True, align="center")
            )
            self.sheet.merge_range(
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
                        self.sheet.write(
                            f"{self.cols[2 + col]}{row}",
                            _k.replace("-", " "),
                            self.format(bold=True, align="center")
                        )
                        col += 1
                    self.sheet.write(
                        f"{self.cols[2 + col]}{row}",
                        "Total",
                        self.format(bold=True, align="center")
                    )
                col = 0
                total = 0
                for _key, _value in v.items():
                    self.sheet.write(
                        f"{self.cols[2 + col]}{row + 1 + r}",
                        _value,
                        self.format(bold=False, align="center")
                    )
                    total += _value
                    totals[col].append(_value)
                    col += 1
                self.sheet.write(
                    f"{self.cols[2 + col]}{row + 1 + r}",
                    total,
                    self.format(bold=False, align="center")
                )
                totals[col].append(total)
                self.sheet.write(
                    f"B{row + 1 + r}",
                    k,
                    self.format(bold=True, align="center")
                )
                r += 1
            for i in range(12):
                self.sheet.merge_range(
                    f"A{row + 13}:B{row + 13}",
                    "Total",
                    self.format(bold=True, align="center")
                )
            for index, i in enumerate(totals):
                self.sheet.write(
                    f"{self.cols[2 + index]}{row + 13}",
                    sum(i),
                    self.format(bold=False, align="center")
                )
            row += 15

    def write_info(self, info):
        for index, (key, value) in enumerate(info.items()):
            if index < 6:
                self.sheet.merge_range(
                    f"A{index + 1}:B{index + 1}",
                    key + ":",
                    self.format(bold=True, align="left")
                )
                self.sheet.write(
                    f"C{index + 1}",
                    value,
                    self.format(bold=False, align="left")
                )
            else:
                self.sheet.merge_range(
                    f"D{index - 5}:E{index - 5}",
                    key + ":",
                    self.format(bold=True, align="left")
                )
                self.sheet.merge_range(
                    f"F{index - 5}:N{index - 5}",
                    value,
                    self.format(bold=False, align="left")
                )
