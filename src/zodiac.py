# -*- coding: utf-8 -*-

from .consts import PLANETS, SIGNS, HOUSE_SYSTEMS


class Zodiac:
    def __init__(
        self,
        jd,
        lat,
        lon,
        hsys,
        zodiac,
        ayanamsha,
        swe
    ):
        self.jd = jd
        self.lat = lat
        self.lon = lon
        self.hsys = hsys.encode("utf-8")
        self.zodiac = zodiac
        self.swe = swe
        if zodiac == "Sidereal":
            self.swe.set_sid_mode(ayanamsha)
        self.diff = self.swe.get_ayanamsa(jd) if zodiac == "Sidereal" else 0

    @staticmethod
    def get_hsys():
        return {i[0]: i[i] for i in HOUSE_SYSTEMS}

    @staticmethod
    def convert_degree(degree):
        q, r = divmod(degree, 30)
        return tuple(SIGNS)[int(q)], r

    @staticmethod
    def reverse_convert_degree(degree):
        return tuple(SIGNS).index(degree[1]) * 30 + degree[2]

    def planet_pos(self, hpos):
        for key in tuple(PLANETS)[:-2]:
            degree = (self.swe.calc(self.jd, PLANETS[key])[0][0] - self.diff) % 360
            yield key, *self.convert_degree(degree), self.calc_planet_house(degree, hpos)

    def house_pos(self):

        for i, j in enumerate(self.swe.houses(self.jd, self.lat, self.lon, self.hsys)[0]):
            yield f"House-{i + 1}", *self.convert_degree((j - self.diff) % 360)

    def calc_planet_house(self, ppos, hpos):
        for i, j in enumerate(hpos):
            previous, current = tuple(map(self.reverse_convert_degree, [hpos[i - 1], hpos[i]]))
            if previous < ppos < current:
                return f"House-{12 if not i else i}"
            if previous > current:
                if previous > current > ppos or ppos > previous > current:
                    return f"House-{12 if not i else i}"

    def patterns(self):
        hp = tuple(self.house_pos())
        pp = tuple(self.planet_pos(hp))
        pp += (("Ascendant", *hp[0][1:], "House-1"), ("Midheaven", *hp[9][1:], "House-10"))
        return pp, hp
