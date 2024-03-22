# -*- coding: utf-8 -*-

from .libs import swe

ZODIAC = ["Tropical", "Sidereal"]

HOUSE_SYSTEMS = {
    "Placidus": "P",
    "Koch": "K",
    "Porphyrius": "O",
    "Regiomontanus": "R",
    "Campanus": "C",
    "Equal": "E",
    "Whole Signs": "W"
}

ORB_FACTORS = {
    "Conjunction": 10,
    "Semi-Sextile": 3,
    "Semi-Square": 3,
    "Sextile": 6,
    "Quintile": 2,
    "Square": 10,
    "Trine": 10,
    "Sesquiquadrate": 3,
    "BiQuintile": 2,
    "Quincunx": 3,
    "Opposite": 10
}

MIDPONT_ORB_FACTORS = {
    "Conjunction": 2,
    "Semi-Sextile": 1,
    "Semi-Square": 1.5,
    "Sextile": 1,
    "Quintile": 1,
    "Square": 2,
    "Trine": 1,
    "Sesquiquadrate": 1.5,
    "BiQuintile": 1,
    "Quincunx": 1,
    "Opposite": 2
}

OPTIONAL_SELECTIONS = [
    "Zodiac",
    "House Systems",
    "Orb Factors",
    "Midpoint Orb Factors",
    "Ignored Categories",
    "Ignored Records",
    "Year Range",
    "Latitude Range",
    "Longitude Range"
]

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
    "Midpoints"
]

IGNORED_RECORDS = [
    "Event",
    "Human",
    "Male",
    "Female",
    "North Hemisphere",
    "South Hemisphere",
    "West Hemisphere",
    "East Hemisphere"
]

RODDEN_RATINGS = [
    "AA",
    "A",
    "B",
    "C",
    "DD",
    "X",
    "XX"
]

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

SIGNS = (
    "Aries",
    "Taurus",
    "Gemini",
    "Cancer",
    "Leo",
    "Virgo",
    "Libra",
    "Scorpio",
    "Sagittarius",
    "Capricorn",
    "Aquarius",
    "Pisces"
)
PLANETS = {
    'Sun': swe.SUN,
    'Moon': swe.MOON,
    'Mercury': swe.MERCURY,
    'Venus': swe.VENUS,
    'Mars': swe.MARS,
    'Jupiter': swe.JUPITER,
    'Saturn': swe.SATURN,
    'Uranus': swe.URANUS,
    'Neptune': swe.NEPTUNE,
    'Pluto': swe.PLUTO,
    'Mean Node': swe.MEAN_NODE,
    'True Node': swe.TRUE_NODE,
    'Chiron': swe.CHIRON,
    'Ascendant': None,
    'Midheaven': None
}

HOUSES = tuple(f"House-{i}" for i in range(1, 13))
ELEMENTS = ("Fire", "Earth", "Air", "Water")
MODES = ("Cardinal", "Fixed", "Mutable")
LORDS = tuple(f"Lord-{i}" for i in range(1, 13))

CALC_TYPES = [
    "Calculate Expected Values",
    "Calculate Chi-Square Values",
    "Calculate Effect-Size Values",
    "Calculate Cohen's D Values",
    "Calculate Binomial Probability Values",
    "Calculate Z-Score Values",
    "Calculate Significance Values"
]
