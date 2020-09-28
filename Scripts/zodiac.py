# -*- coding: utf-8 -*-

from .modules import os, swe
from .constants import PLANETS
from .utilities import convert_degree, reverse_convert_degree

swe.set_ephe_path(os.path.join(os.getcwd(), "Eph"))


class Zodiac:
    def __init__(self, jd, lat, lon, hsys):
        self.jd = jd
        self.lat = lat
        self.lon = lon
        self.hsys = hsys

    def planet_pos(self, planet):
        degree = swe.calc(self.jd, planet)[0]
        if isinstance(degree, tuple):
            degree = degree[0]
        calc = convert_degree(degree=degree)
        return calc[1], reverse_convert_degree(calc[0], calc[1])

    def house_pos(self):
        house = []
        asc = 0
        degree = []
        for i, j in enumerate(swe.houses(
                self.jd, self.lat, self.lon,
                bytes(self.hsys.encode("utf-8")))[0]):
            if i == 0:
                asc += j
            degree.append(j)
            house.append((
                f"{i + 1}",
                j,
                f"{convert_degree(j)[1]}"))
        return house, asc, degree

    def patterns(self):
        planet_positions = []
        house_positions = []
        for i in range(12):
            house = [
                int(self.house_pos()[0][i][0]),
                self.house_pos()[0][i][-1],
                float(self.house_pos()[0][i][1]),
            ]
            house_positions.append(house)
        hp = [j[-1] for j in house_positions]
        for key, value in PLANETS.items():
            if value["number"] is None:
                continue
            planet = self.planet_pos(planet=value["number"])
            house = 0
            for i in range(12):
                if i != 11:
                    if hp[i] < planet[1] < hp[i + 1]:
                        house = i + 1
                        break
                    elif hp[i] < planet[1] > hp[i + 1] \
                            and hp[i] - hp[i + 1] > 240:
                        house = i + 1
                        break
                    elif hp[i] > planet[1] < hp[i + 1] \
                            and hp[i] - hp[i + 1] > 240:
                        house = i + 1
                        break
                else:
                    if hp[i] < planet[1] < hp[0]:
                        house = i + 1
                        break
                    elif hp[i] < planet[1] > hp[0] \
                            and hp[i] - hp[0] > 240:
                        house = i + 1
                        break
                    elif hp[i] > planet[1] < hp[0] \
                            and hp[i] - hp[0] > 240:
                        house = i + 1
                        break
            planet_info = [
                key,
                planet[0],
                planet[1],
                f"House-{house}"
            ]
            planet_positions.append(planet_info)
        asc = house_positions[0] + ["House-1"]
        asc[0] = "Ascendant"
        mc = house_positions[9] + ["House-10"]
        mc[0] = "Midheaven"
        return planet_positions, house_positions
