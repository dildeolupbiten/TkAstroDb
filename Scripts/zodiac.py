# -*- coding: utf-8 -*-

from .modules import os, swe, ConfigParser
from .constants import PLANETS, TRADITIONAL_RULERSHIP, MODERN_RULERSHIP
from .utilities import convert_degree, get_element, get_mode, find_aspect

swe.set_ephe_path(os.path.join(os.getcwd(), "Eph"))


class Zodiac:
    def __init__(self, jd, lat, lon, hsys):
        self.jd = jd
        self.lat = lat
        self.lon = lon
        self.hsys = hsys

    def planet_pos(self, planet):
        degree = swe.calc(self.jd, planet)[0]
        return [convert_degree(degree)[0], degree]

    def house_pos(self):
        return [
            [i + 1, convert_degree(j)[0], j]
            for i, j in enumerate(
                swe.houses(
                    self.jd,
                    self.lat,
                    self.lon,
                    bytes(self.hsys.encode("utf-8"))
                )[0]
            )
        ]        

    def patterns(
        self,
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
        temporary
    ):
        pp = []
        hp = self.house_pos()
        traditional = {}
        modern = {}
        config = ConfigParser()
        config.read("defaults.ini")
        if any(
            [
                houses_in_signs,
                houses_in_elements,
                houses_in_modes,
                basic_traditional_rulership,
                basic_modern_rulership,
            ]
        ):
            for i in hp:
                if houses_in_signs:
                    houses_in_signs[f"House-{i[0]}"][i[1]] += 1
                if houses_in_elements:
                    houses_in_elements[f"House-{i[0]}"][
                        get_element(i[1])
                    ] += 1
                if houses_in_modes:
                    houses_in_modes[f"House-{i[0]}"][
                        get_mode(i[1])
                    ] += 1
                if (
                    basic_traditional_rulership 
                    or 
                    detailed_traditional_rulership
                ):
                    traditional[
                        f"Lord-{i[0]}"
                    ] = TRADITIONAL_RULERSHIP[i[1]]
                if (
                    basic_modern_rulership
                    or
                    detailed_modern_rulership
                ):
                    modern[
                        f"Lord-{i[0]}"
                    ] = MODERN_RULERSHIP[i[1]]             
        for key, value in PLANETS.items():
            if value["number"] is None:
                continue
            pos = self.planet_pos(planet=value["number"])
            if aspects:
                for i in pp:
                    find_aspect(
                        aspects=aspects,
                        temporary=temporary,
                        orb=config["ORB FACTORS"],
                        planet1=i[0],
                        planet2=key,
                        aspect=abs(pos[1] - i[2])
                    )
                pp.append([key, *pos])
            if planets_in_signs:
                planets_in_signs[key][pos[0]] += 1
            if planets_in_elements:
                planets_in_elements[key][get_element(pos[0])] += 1
            if planets_in_modes:
                planets_in_modes[key][get_mode(pos[0])] += 1
            house = 0
            for i in range(12):
                if i != 11:
                    if (
                        hp[i][-1] < hp[i + 1][-1]
                        and
                        hp[i][-1] < pos[1] < hp[i + 1][-1]
                    ):
                        house = i + 1
                        break
                    elif (
                        hp[i][-1] > hp[i + 1][-1]
                        and
                        hp[i][-1] < pos[1] > hp[i + 1][-1]
                    ):
                        house = i + 1
                        break
                    elif (
                        hp[i][-1] > hp[i + 1][-1] > pos[1] < hp[i][-1]
                    ):
                        house = i + 1
                        break
                else:
                    if (
                        hp[11][-1] < hp[0][-1]
                        and
                        hp[11][-1] < pos[1] < hp[0][-1]
                    ):
                        house = i + 1
                        break
                    elif (
                        hp[11][-1] > hp[0][-1]
                        and
                        hp[11][-1] < pos[1] > hp[0][-1]
                    ):
                        house = i + 1
                        break
                    elif (
                        hp[11][-1] > hp[0][-1] > pos[1] < hp[11][-1]
                    ):
                        house = i + 1
                        break
            if planets_in_houses:
                planets_in_houses[key][f"House-{house}"] += 1
            if planets_in_houses_in_signs:
                planets_in_houses_in_signs[key][f"House-{house}"][
                    pos[0]
                ] += 1
            if (
                basic_traditional_rulership
                or
                detailed_traditional_rulership
            ):
                for k, v in traditional.items():
                    if v.split(" ")[0] == key:
                        if basic_traditional_rulership:
                            basic_traditional_rulership[k][
                                f"House-{house}"
                            ] += 1
                        if detailed_traditional_rulership:
                            detailed_traditional_rulership[k][v][
                                f"House-{house}"
                            ] += 1
            if (
                basic_modern_rulership
                or
                detailed_modern_rulership
            ):
                for k, v in modern.items():
                    if v.split(" ")[0] == key:
                        if basic_modern_rulership:
                            basic_modern_rulership[k][
                                f"House-{house}"
                            ] += 1
                        if detailed_modern_rulership:
                            detailed_modern_rulership[k][v][
                                f"House-{house}"
                            ] += 1
        if aspects:
            asc = ["Ascendant"] + hp[0][1:]
            mc = ["Midheaven"] + hp[9][1:]
            for key in [asc, mc]:
                for i in pp:
                    find_aspect(
                        aspects=aspects,
                        temporary=temporary,
                        orb=orb_factors,
                        planet1=i[0],
                        planet2=key[0],
                        aspect=abs(key[2] - i[2])
                    )
                pp.append(key)
