# materials.py
# Alle materialdatabaser, konstanter og profiler for byggTotal.
# Priser hentet fra Norsk Prisbok 2024 og interne prosjektfaktorer.
# Oppdater PRICE_VERSION ved ny Prisbok-utgave.

PRICE_VERSION = "Norsk Prisbok 2024"

# ---------------------------------------------------------------------------
# Tettheter (kg/m³)
# ---------------------------------------------------------------------------
STEEL_DENSITY = 7850.0
GLULAM_DENSITY = 460.0
CLT_DENSITY = 500.0
TIMBER_DENSITY = 450.0
CONCRETE_DENSITY = 2400.0

# ---------------------------------------------------------------------------
# IFC-konstanter
# ---------------------------------------------------------------------------
BYGGTOTAL_PSET_NAME = "Pset_ByggTotal"
BYGGTOTAL_CHANGED_PROP = "ByggTotal_Changed"
BYGGTOTAL_OLD_MATERIAL_PROP = "ByggTotal_OldMaterial"
BYGGTOTAL_NEW_MATERIAL_PROP = "ByggTotal_NewMaterial"
BYGGTOTAL_PROFILE_PROP = "ByggTotal_NewProfile"

SUPPORTED_IFC_TYPES = [
    "IfcBeam", "IfcColumn", "IfcSlab", "IfcWall", "IfcWallStandardCase", "IfcRoof", "IfcMember", "IfcFooting"
]

IFC_GEOMETRY_FALLBACK_TYPES = {"IfcSlab", "IfcBeam", "IfcColumn"}
MAX_PROFILE_OPTIONS_DEFAULT = 100

# ---------------------------------------------------------------------------
# Materialdatabase – primærdatabase for enkle materialvalg
# Priser i NOK (Norsk Prisbok 2024)
# ---------------------------------------------------------------------------
MATERIAL_DATABASE = {
    "Stål": {"unit": "kg", "price": 47.0, "co2": 0.73, "density": 7850.0, "label": "Stål"},
    "Limtre": {"unit": "m3", "price": 28000.0, "co2": 100.0, "density": 460.0, "label": "Limtre"},
    "Massivtre": {"unit": "m3", "price": 30000.0, "co2": 110.0, "density": 500.0, "label": "Massivtre / CLT"},
    "Tre": {"unit": "m3", "price": 5000.0, "co2": 120.0, "density": 450.0, "label": "Tre"},
    "Betong_volum": {"unit": "m3", "price": 1800.0, "co2": 350.0, "density": 2400.0, "label": "Betong volum"},
    "Hulldekke": {"unit": "m2", "price": 1635.0, "co2": 84.56, "density": 2400.0, "label": "Hulldekke"},
    "Hulldekke_lavCO2": {"unit": "m2", "price": 1821.0, "co2": 64.86, "density": 2400.0, "label": "Hulldekke lavCO₂"},
    "Plasstøpt_betong": {"unit": "m2", "price": 2422.0, "co2": 69.59, "density": 2400.0, "label": "Plasstøpt betong"},
    "Plasstøpt_betong_lavCO2": {"unit": "m2", "price": 3015.0, "co2": 54.64, "density": 2400.0, "label": "Plasstøpt betong lavCO₂"},
    "Massivtre_vegg": {"unit": "m2", "price": 1337.0, "co2": 8.93, "density": 500.0, "label": "Massivtre vegg"},
    "Betong_vegg": {"unit": "m2", "price": 2910.0, "co2": 52.84, "density": 2400.0, "label": "Betong vegg"},
    "Betong_vegg_lavCO2": {"unit": "m2", "price": 3370.0, "co2": 43.54, "density": 2400.0, "label": "Betong vegg lavCO₂"},
    "Ukjent": {"unit": "m3", "price": 1000.0, "co2": 200.0, "density": 1000.0, "label": "Ukjent"},
}

# ---------------------------------------------------------------------------
# Norsk Prisbok-database – detaljerte produktvalg med koder
# Kilde: Norsk Prisbok 2024
# ---------------------------------------------------------------------------
NORSK_PRISBOK_DATABASE = {
    "Betong_vegg_150": {"category": "Vegg", "unit": "m2", "price": 2566.0, "co2": 56.10, "ak": 141.78, "label": "Prefab betongyttervegg over mark, t = 150 mm", "npb_code": "02.3.B.001", "source": "Norsk Prisbok", "thickness_mm": 150},
    "Betong_vegg_180": {"category": "Vegg", "unit": "m2", "price": 2885.0, "co2": 67.32, "ak": 159.39, "label": "Prefab betongyttervegg over mark, t = 180 mm", "npb_code": "02.3.B.002", "source": "Norsk Prisbok", "thickness_mm": 180},
    "Betong_vegg_200": {"category": "Vegg", "unit": "m2", "price": 3100.0, "co2": 74.80, "ak": 171.28, "label": "Prefab betongyttervegg over mark, t = 200 mm", "npb_code": "02.3.B.003", "source": "Norsk Prisbok", "thickness_mm": 200},
    "Massivtre_vegg_100": {"category": "Vegg", "unit": "m2", "price": 1575.0, "co2": 11.16, "ak": 96.06, "label": "Massive treelementer, yttervegg, t = 100 mm", "npb_code": "02.3.1.5.0110", "source": "Norsk Prisbok", "thickness_mm": 100},
    "Massivtre_vegg_120": {"category": "Vegg", "unit": "m2", "price": 1879.0, "co2": 13.40, "ak": 114.58, "label": "Massive treelementer, yttervegg, t = 120 mm", "npb_code": "02.3.1.5.0120", "source": "Norsk Prisbok", "thickness_mm": 120},
    "Massivtre_vegg_140": {"category": "Vegg", "unit": "m2", "price": 2177.0, "co2": 15.63, "ak": 132.70, "label": "Massive treelementer, yttervegg, t = 140 mm", "npb_code": "02.3.1.5.0130", "source": "Norsk Prisbok", "thickness_mm": 140},
    "Massivtre_vegg_160": {"category": "Vegg", "unit": "m2", "price": 2466.0, "co2": 17.86, "ak": 150.32, "label": "Massive treelementer, yttervegg, t = 160 mm", "npb_code": "02.3.1.5.0140", "source": "Norsk Prisbok", "thickness_mm": 160},
    "Massivtre_vegg_200": {"category": "Vegg", "unit": "m2", "price": 2897.0, "co2": 22.33, "ak": 160.08, "label": "Massive treelementer, yttervegg, t = 200 mm", "npb_code": "02.3.1.5.0160", "source": "Norsk Prisbok", "thickness_mm": 200},
    "Massivtre_vegg_240": {"category": "Vegg", "unit": "m2", "price": 3225.0, "co2": 26.80, "ak": 178.20, "label": "Massive treelementer, yttervegg, t = 240 mm", "npb_code": "02.3.1.5.0170", "source": "Norsk Prisbok", "thickness_mm": 240},
    "Plasstopt_dekke_180": {"category": "Dekke", "unit": "m2", "price": 2285.0, "co2": 62.87, "ak": 126.26, "label": "Betongdekke, t = 180 mm", "npb_code": "02.5.B.001", "source": "Norsk Prisbok", "thickness_mm": 180},
    "Plasstopt_dekke_200": {"category": "Dekke", "unit": "m2", "price": 2422.0, "co2": 69.59, "ak": 133.81, "label": "Betongdekke, t = 200 mm", "npb_code": "02.5.B.002", "source": "Norsk Prisbok", "thickness_mm": 200},
    "Plasstopt_dekke_220": {"category": "Dekke", "unit": "m2", "price": 2559.0, "co2": 76.31, "ak": 141.37, "label": "Betongdekke, t = 220 mm", "npb_code": "02.5.B.003", "source": "Norsk Prisbok", "thickness_mm": 220},
    "Plasstopt_dekke_250": {"category": "Dekke", "unit": "m2", "price": 2764.0, "co2": 86.40, "ak": 152.71, "label": "Betongdekke, t = 250 mm", "npb_code": "02.5.B.004", "source": "Norsk Prisbok", "thickness_mm": 250},
    "Plasstopt_dekke_300": {"category": "Dekke", "unit": "m2", "price": 3157.0, "co2": 103.25, "ak": 174.41, "label": "Betongdekke, t = 300 mm", "npb_code": "02.5.B.005", "source": "Norsk Prisbok", "thickness_mm": 300},
    "Plasstopt_dekke_350": {"category": "Dekke", "unit": "m2", "price": 3499.0, "co2": 120.06, "ak": 193.31, "label": "Betongdekke, t = 350 mm", "npb_code": "02.5.B.006", "source": "Norsk Prisbok", "thickness_mm": 350},
    "Plasstopt_dekke_lavCO2": {"category": "Dekke", "unit": "m2", "price": 3015.0, "co2": 54.64, "ak": 166.58, "label": "Betongdekke med redusert klimagassutslipp", "npb_code": "02.5.B.007", "source": "Norsk Prisbok", "thickness_mm": None},
    "Hulldekke_200": {"category": "Dekke", "unit": "m2", "price": 1490.0, "co2": 65.06, "ak": 82.34, "label": "HD-element, t = 200 mm", "npb_code": "02.5.C.001", "source": "Norsk Prisbok", "thickness_mm": 200},
    "Hulldekke_220": {"category": "Dekke", "unit": "m2", "price": 1577.0, "co2": 72.14, "ak": 87.13, "label": "HD-element, t = 220 mm", "npb_code": "02.5.C.002", "source": "Norsk Prisbok", "thickness_mm": 220},
    "Hulldekke_265": {"category": "Dekke", "unit": "m2", "price": 1635.0, "co2": 84.56, "ak": 90.32, "label": "HD-element, t = 265 mm", "npb_code": "02.5.C.003", "source": "Norsk Prisbok", "thickness_mm": 265},
    "Hulldekke_265_lavCO2": {"category": "Dekke", "unit": "m2", "price": 1821.0, "co2": 64.86, "ak": 100.64, "label": "HD-element, t = 265 mm, lavCO₂", "npb_code": "02.5.C.004", "source": "Norsk Prisbok", "thickness_mm": 265},
    "Hulldekke_290": {"category": "Dekke", "unit": "m2", "price": 1721.0, "co2": 86.94, "ak": 95.11, "label": "HD-element, t = 290 mm", "npb_code": "02.5.C.005", "source": "Norsk Prisbok", "thickness_mm": 290},
    "Hulldekke_320": {"category": "Dekke", "unit": "m2", "price": 1779.0, "co2": 93.80, "ak": 98.31, "label": "HD-element, t = 320 mm", "npb_code": "02.5.C.006", "source": "Norsk Prisbok", "thickness_mm": 320},
    "Hulldekke_320_lavCO2": {"category": "Dekke", "unit": "m2", "price": 1970.0, "co2": 71.62, "ak": 108.82, "label": "HD-element, t = 320 mm, lavCO₂", "npb_code": "02.5.C.007", "source": "Norsk Prisbok", "thickness_mm": 320},
    "Hulldekke_340": {"category": "Dekke", "unit": "m2", "price": 1808.0, "co2": 95.40, "ak": 99.90, "label": "HD-element, t = 340 mm", "npb_code": "02.5.C.008", "source": "Norsk Prisbok", "thickness_mm": 340},
    "Hulldekke_400": {"category": "Dekke", "unit": "m2", "price": 1837.0, "co2": 100.48, "ak": 101.50, "label": "HD-element, t = 400 mm", "npb_code": "02.5.C.009", "source": "Norsk Prisbok", "thickness_mm": 400},
    "Hulldekke_420": {"category": "Dekke", "unit": "m2", "price": 1924.0, "co2": 107.76, "ak": 106.29, "label": "HD-element, t = 420 mm", "npb_code": "02.5.C.010", "source": "Norsk Prisbok", "thickness_mm": 420},
    "Hulldekke_500": {"category": "Dekke", "unit": "m2", "price": 2110.0, "co2": 127.88, "ak": 116.60, "label": "HD-element, t = 500 mm", "npb_code": "02.5.C.011", "source": "Norsk Prisbok", "thickness_mm": 500},
    "Massivtre_dekke_160": {"category": "Dekke", "unit": "m2", "price": 2570.0, "co2": 17.86, "ak": 142.01, "label": "Massivtre dekke, t = 160 mm", "npb_code": "02.5.C.031", "source": "Norsk Prisbok", "thickness_mm": 160},
    "Massivtre_dekke_180": {"category": "Dekke", "unit": "m2", "price": 2798.0, "co2": 20.10, "ak": 154.61, "label": "Massivtre dekke, t = 180 mm", "npb_code": "02.5.C.032", "source": "Norsk Prisbok", "thickness_mm": 180},
    "Massivtre_dekke_200": {"category": "Dekke", "unit": "m2", "price": 3018.0, "co2": 22.33, "ak": 166.77, "label": "Massivtre dekke, t = 200 mm", "npb_code": "02.5.C.033", "source": "Norsk Prisbok", "thickness_mm": 200},
    "Massivtre_dekke_220": {"category": "Dekke", "unit": "m2", "price": 3161.0, "co2": 24.56, "ak": 174.64, "label": "Massivtre dekke, t = 220 mm", "npb_code": "02.5.C.034", "source": "Norsk Prisbok", "thickness_mm": 220},
    "Massivtre_dekke_240": {"category": "Dekke", "unit": "m2", "price": 3419.0, "co2": 26.80, "ak": 188.88, "label": "Massivtre dekke, t = 240 mm", "npb_code": "02.5.C.035", "source": "Norsk Prisbok", "thickness_mm": 240},
    "Massivtre_dekke_260": {"category": "Dekke", "unit": "m2", "price": 3700.0, "co2": 29.03, "ak": 204.42, "label": "Massivtre dekke, t = 260 mm", "npb_code": "02.5.C.036", "source": "Norsk Prisbok", "thickness_mm": 260},
    "Massivtre_dekke_280": {"category": "Dekke", "unit": "m2", "price": 3972.0, "co2": 31.26, "ak": 219.46, "label": "Massivtre dekke, t = 280 mm", "npb_code": "02.5.C.037", "source": "Norsk Prisbok", "thickness_mm": 280},
}

# ---------------------------------------------------------------------------
# EPD-database – klimagasskoeffisienter (kgCO₂e per enhet)
# Kilde: EPD-Norge / prosjektfaktorer
# ---------------------------------------------------------------------------
EPD_DATABASE = {
    "Stål": {"unit": "kg", "co2": 0.73, "source": "EPD / prosjektfaktor"},
    "Limtre": {"unit": "m3", "co2": 100.0, "source": "EPD / prosjektfaktor"},
    "Massivtre": {"unit": "m3", "co2": 110.0, "source": "EPD / prosjektfaktor"},
    "Tre": {"unit": "m3", "co2": 120.0, "source": "EPD / prosjektfaktor"},
    "Betong_volum": {"unit": "m3", "co2": 350.0, "source": "EPD / prosjektfaktor"},
    "Hulldekke": {"unit": "m2", "co2": 84.56, "source": "EPD / prosjektfaktor"},
    "Hulldekke_lavCO2": {"unit": "m2", "co2": 64.86, "source": "EPD / prosjektfaktor"},
    "Plasstøpt_betong": {"unit": "m2", "co2": 69.59, "source": "EPD / prosjektfaktor"},
    "Plasstøpt_betong_lavCO2": {"unit": "m2", "co2": 54.64, "source": "EPD / prosjektfaktor"},
    "Massivtre_vegg": {"unit": "m2", "co2": 8.93, "source": "EPD / prosjektfaktor"},
    "Betong_vegg": {"unit": "m2", "co2": 52.84, "source": "EPD / prosjektfaktor"},
    "Betong_vegg_lavCO2": {"unit": "m2", "co2": 43.54, "source": "EPD / prosjektfaktor"},
}

# ---------------------------------------------------------------------------
# Profilbibliotek
# ---------------------------------------------------------------------------
PROFILE_LIBRARY = {
    "Limtre": ["90x315", "90x405", "115x315", "115x360", "115x405", "140x315", "140x360", "140x405", "140x450", "165x315", "165x360", "165x405", "190x405", "190x450", "215x405", "215x450"],
    "Massivtre": ["100x300", "120x300", "120x400", "140x400", "160x400", "200x400"],
    "Stål": ["KFHUP 120x120x8", "KFHUP 140x140x10", "KFHUP 160x160x10", "KFHUP 180x180x12.5", "KFHUP 200x200x12.5", "KFHUP 220x220x12.5"],
    "Betong": ["200x200", "250x250", "300x300", "350x350", "400x400"],
}

# ---------------------------------------------------------------------------
# Kvalitetsbibliotek – stål/limtre/betong/massivtre med tetthet, CO₂ og pris
# ---------------------------------------------------------------------------
QUALITY_LIBRARY = {
    "Stål": {
        "S235": {"density": 7850.0, "co2": 0.73, "price": 43.0},
        "S355": {"density": 7850.0, "co2": 0.73, "price": 47.0},
        "S460": {"density": 7850.0, "co2": 0.78, "price": 55.0},
    },
    "Limtre": {
        "GL24c": {"density": 430.0, "co2": 95.0, "price": 26000.0},
        "GL30c": {"density": 460.0, "co2": 100.0, "price": 28000.0},
        "GL32h": {"density": 480.0, "co2": 105.0, "price": 30500.0},
    },
    "Betong": {
        "B25": {"density": 2350.0, "co2": 310.0, "price": 1650.0},
        "B35": {"density": 2400.0, "co2": 350.0, "price": 1800.0},
        "B45": {"density": 2450.0, "co2": 390.0, "price": 2050.0},
        "Lavkarbon B35": {"density": 2400.0, "co2": 260.0, "price": 2150.0},
    },
    "Massivtre": {
        "C24/CLT": {"density": 500.0, "co2": 110.0, "price": 30000.0},
    },
}

# ---------------------------------------------------------------------------
# BREEAM
# ---------------------------------------------------------------------------
BREEAM_LEVELS = ["Ingen", "Pass", "Good", "Very Good", "Excellent", "Outstanding"]

# ---------------------------------------------------------------------------
# Grunnarbeidssystemer
# ---------------------------------------------------------------------------
GROUND_SYSTEM_LIBRARY = {
    "Standard byggegrop": {
        "rigg_pct": 0.08, "excavation_rate": 185.0, "fill_rate": 145.0, "geotextile_rate": 32.0,
        "stripping_depth": 0.30, "stripping_rate": 95.0, "subbase_thickness": 0.25, "subbase_rate": 420.0,
        "transport_cut_rate": 65.0, "import_fill_rate": 55.0, "drain_rate": 180.0, "stormwater_rate": 0.0,
        "reuse_factor": 0.15, "documentation_pct": 0.00, "label": "Standard byggegrop"
    },
    "Boligtomt / lett grunnarbeid": {
        "rigg_pct": 0.08, "excavation_rate": 165.0, "fill_rate": 135.0, "geotextile_rate": 28.0,
        "stripping_depth": 0.25, "stripping_rate": 85.0, "subbase_thickness": 0.20, "subbase_rate": 390.0,
        "transport_cut_rate": 55.0, "import_fill_rate": 50.0, "drain_rate": 150.0, "stormwater_rate": 0.0,
        "reuse_factor": 0.20, "documentation_pct": 0.00, "label": "Boligtomt / lett grunnarbeid"
    },
    "Næring / hardt belastet tomt": {
        "rigg_pct": 0.10, "excavation_rate": 225.0, "fill_rate": 165.0, "geotextile_rate": 36.0,
        "stripping_depth": 0.35, "stripping_rate": 110.0, "subbase_thickness": 0.35, "subbase_rate": 480.0,
        "transport_cut_rate": 75.0, "import_fill_rate": 65.0, "drain_rate": 230.0, "stormwater_rate": 0.0,
        "reuse_factor": 0.12, "documentation_pct": 0.00, "label": "Næring / hardt belastet tomt"
    },
    "Sprengning / vanskelige masser": {
        "rigg_pct": 0.12, "excavation_rate": 315.0, "fill_rate": 175.0, "geotextile_rate": 38.0,
        "stripping_depth": 0.30, "stripping_rate": 110.0, "subbase_thickness": 0.35, "subbase_rate": 510.0,
        "transport_cut_rate": 95.0, "import_fill_rate": 70.0, "drain_rate": 250.0, "stormwater_rate": 0.0,
        "reuse_factor": 0.08, "documentation_pct": 0.00, "label": "Sprengning / vanskelige masser"
    },
}
