# -*- coding: utf-8 -*-

from .modules import swe

SIGNS = {
    "Aries": {
        "symbol": "\u2648",
        "color": "#FF0000"
    },
    "Taurus": {
        "symbol": "\u2649",
        "color": "#00FF00"
    },
    "Gemini": {
        "symbol": "\u264A",
        "color": "#FFFF00"
    },
    "Cancer": {
        "symbol": "\u264B",
        "color": "#0000FF"
    },
    "Leo": {
        "symbol": "\u264C",
        "color": "#FF0000"
    },
    "Virgo": {
        "symbol": "\u264D",
        "color": "#00FF00"
    },
    "Libra": {
        "symbol": "\u264E",
        "color": "#FFFF00"
    },
    "Scorpio": {
        "symbol": "\u264F",
        "color": "#0000FF"
    },
    "Sagittarius": {
        "symbol": "\u2650",
        "color": "#FF0000"
    },
    "Capricorn": {
        "symbol": "\u2651",
        "color": "#00FF00"
    },
    "Aquarius": {
        "symbol": "\u2652",
        "color": "#FFFF00"
    },
    "Pisces": {
        "symbol": "\u2653",
        "color": "#0000FF"
    }
}

PLANETS = {
    "Sun": {
        "number": swe.SUN,
        "symbol": "\u2299"
    },
    "Moon": {
        "number": swe.MOON,
        "symbol": "\u263E"
    },
    "Mercury": {
        "number": swe.MERCURY,
        "symbol": "\u263F"
    },
    "Venus": {
        "number": swe.VENUS,
        "symbol": "\u2640"
    },
    "Mars": {
        "number": swe.MARS,
        "symbol": "\u2642"
    },
    "Jupiter": {
        "number": swe.JUPITER,
        "symbol": "\u2643"
    },
    "Saturn": {
        "number": swe.SATURN,
        "symbol": "\u2644"
    },
    "Uranus": {
        "number": swe.URANUS,
        "symbol": "\u2645"
    },
    "Neptune": {
        "number": swe.NEPTUNE,
        "symbol": "\u2646"
    },
    "Pluto": {
        "number": swe.PLUTO,
        "symbol": "\u2647"
    },
    "True Node": {
        "number": swe.TRUE_NODE,
        "symbol": "\u260A"
    },
    "Mean Node": {
        "number": swe.MEAN_NODE,
        "symbol": "\u260A"
    },
    "Chiron": {
        "number": swe.CHIRON,
        "symbol": "\u26B7"
    },
    "Ascendant": {
        "number": None,
        "symbol": None
    },
    "Midheaven": {
        "number": None,
        "symbol": None
    }
}

AYANAMSHA = {
    "Aldebaran 15 Taurus": swe.SIDM_ALDEBARAN_15TAU,
    "Aryabhata": swe.SIDM_ARYABHATA,
    "Aryabhata (Mean Sun)": swe.SIDM_ARYABHATA_MSUN,
    "B1950": swe.SIDM_B1950,
    "Babylonian (Huber)": swe.SIDM_BABYL_HUBER,
    "Babylonian (Kugler 1)": swe.SIDM_BABYL_KUGLER1,
    "Babylonian (Kugler 2)": swe.SIDM_BABYL_KUGLER2,
    "Babylonian (Kugler 3)": swe.SIDM_BABYL_KUGLER3,
    "De Luce": swe.SIDM_DELUCE,
    "Fagan-Bradley": swe.SIDM_FAGAN_BRADLEY,
    "Galactic Center 0 Sagittarius": swe.SIDM_GALCENT_0SAG,
    "Hindu-Lahiri": swe.SIDM_LAHIRI,
    "Hipparchus": swe.SIDM_HIPPARCHOS,  
    "J1900": swe.SIDM_J1900,
    "J2000": swe.SIDM_J2000,
    "JN Bhasin": swe.SIDM_JN_BHASIN,
    "Krishnamurti": swe.SIDM_KRISHNAMURTI,
    "Raman": swe.SIDM_RAMAN,
    "Sassanian": swe.SIDM_SASSANIAN,
    "Suryasiddhanta": swe.SIDM_SURYASIDDHANTA,
    "Suryasiddhanta Citra": swe.SIDM_SS_CITRA,
    "Suryasiddhanta Revati": swe.SIDM_SS_REVATI,
    "Suryasiddhanta (Mean Sun)": swe.SIDM_SURYASIDDHANTA_MSUN,
    "True Citra": swe.SIDM_TRUE_CITRA,
    "True Revati": swe.SIDM_TRUE_REVATI,   
    "Usha & Sashi": swe.SIDM_USHASHASHI,
    "Yukteshwar": swe.SIDM_YUKTESHWAR   
}

HOUSES = [f"House-{i}" for i in range(1, 13)]

HOUSE_SYSTEMS = {
    "Placidus": "P",
    "Koch": "K",
    "Porphyrius": "O",
    "Regiomontanus": "R",
    "Campanus": "C",
    "Equal": "E",
    "Whole Signs": "W"
}

MODERN_RULERSHIP = {
    "Aries": "Pluto (+)",
    "Taurus": "Venus (-)",
    "Gemini": "Mercury (+)",
    "Cancer": "Moon (-)",
    "Leo": "Sun (+)",
    "Virgo": "Mercury (-)",
    "Libra": "Venus (+)",
    "Scorpio": "Pluto (-)",
    "Sagittarius": "Neptune (+)",
    "Capricorn": "Uranus (-)",
    "Aquarius": "Uranus (+)",
    "Pisces": "Neptune (-)"
}

TRADITIONAL_RULERSHIP = {
    "Aries": "Mars (+)",
    "Taurus": "Venus (-)",
    "Gemini": "Mercury (+)",
    "Cancer": "Moon (-)",
    "Leo": "Sun (+)",
    "Virgo": "Mercury (-)",
    "Libra": "Venus (+)",
    "Scorpio": "Mars (-)",
    "Sagittarius": "Jupiter (+)",
    "Capricorn": "Saturn (-)",
    "Aquarius": "Saturn (+)",
    "Pisces": "Jupiter (-)"
}

SHEETS = [
    "Info",
    "Planets In Signs",
    "Planets In Elements",
    "Planets In Modes",
    "Houses In Signs",
    "Houses In Elements",
    "Houses In Modes",
    "Planets In Houses",
    "Planets In Houses In Signs",
    "Basic Traditional Rulership",
    "Basic Modern Rulership",
    "Detailed Traditional Rulership",
    "Detailed Modern Rulership",
    "Aspects",
    "Sum Of Aspects",
    "Yod",
    "T-Square",
    "Grand Trine",
    "Mystic Rectangle",
    "Grand Cross",
    "Kite",
    "Midpoints"
]

ASPECTS = [
    "Conjunction",
    "Semi-Sextile",
    "Semi-Square",
    "Sextile",
    "Quintile",
    "Square",
    "Trine",
    "Sesquiquadrate",
    "Biquintile",
    "Quincunx",
    "Opposite"
]
