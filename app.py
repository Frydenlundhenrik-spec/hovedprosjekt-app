import io
import math
import os
import re
import tempfile
import hashlib
from pathlib import Path
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    import openpyxl
except Exception:
    openpyxl = None

try:
    import ifcopenshell
    import ifcopenshell.geom
    from ifcopenshell.util import element as ifc_element_util
except Exception:
    ifcopenshell = None
    ifc_element_util = None

try:
    from docx import Document
except Exception:
    Document = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
except Exception:
    SimpleDocTemplate = None


# ---------------------------------------------------------------------------
# MODUL-OVERSIKT
# Denne filen er hoved-app. Logikk er delt i:
#   materials.py  – alle databaser og konstanter
#   geometry.py   – terreng og konveks hull
#   ifc_utils.py  – IFC-lesing og materialbytte
#   reports.py    – DOCX/PDF-eksport
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="byggTotal – Mengder, kalkyle, IFC og prosjektering",
    page_icon="🏗️",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Session state – initialiseres før resten av appen kjøres
# ---------------------------------------------------------------------------
if "deck_variant_key" not in st.session_state:
    st.session_state["deck_variant_key"] = "Hulldekke"
if "concrete_variant_key" not in st.session_state:
    st.session_state["concrete_variant_key"] = "Plasstøpt_betong"
if "wall_variant_key" not in st.session_state:
    st.session_state["wall_variant_key"] = "Betong_vegg"
if "breeam_target_level" not in st.session_state:
    st.session_state["breeam_target_level"] = "Ingen"
if "breeam_active" not in st.session_state:
    st.session_state["breeam_active"] = False



st.markdown("""
<style>
    .main { background-color: #f6f7fb; }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; padding-left: 2rem; padding-right: 2rem; }
    h1, h2, h3 { color: #1f2937; }
    .stMetric, div[data-testid="stDataFrame"], div[data-testid="stPlotlyChart"], div[data-testid="stPyplot"] {
        background: white; padding: 14px; border-radius: 16px; border: 1px solid #e5e7eb; box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }
    section[data-testid="stSidebar"] { background-color: #eef2f7; }
    .stButton > button, .stDownloadButton > button {
        border-radius: 12px; border: none; background-color: #1f4e79; color: white; font-weight: 600; padding: 0.6rem 1rem;
    }
    .stButton > button:hover, .stDownloadButton > button:hover { background-color: #163a5a; color: white; }
    .custom-card {
        background: white; padding: 18px 20px; border-radius: 18px; border: 1px solid #e5e7eb;
        box-shadow: 0 4px 14px rgba(0,0,0,0.05); margin-bottom: 1rem;
    }
    .custom-title { font-size: 15px; font-weight: 600; color: #6b7280; margin-bottom: 6px; }
    .custom-value { font-size: 28px; font-weight: 700; color: #111827; }
    .small-muted { color: #6b7280; font-size: 13px; }
</style>
""", unsafe_allow_html=True)

STEEL_DENSITY = 7850.0
GLULAM_DENSITY = 460.0
CLT_DENSITY = 500.0
TIMBER_DENSITY = 450.0
CONCRETE_DENSITY = 2400.0

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


# Priser fra Norsk Prisbok 2024 – oppdater PRICE_VERSION i materials.py ved ny utgave
PRICE_VERSION = "Norsk Prisbok 2024"

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

PROFILE_LIBRARY = {
    "Limtre": ["90x315", "90x405", "115x315", "115x360", "115x405", "140x315", "140x360", "140x405", "140x450", "165x315", "165x360", "165x405", "190x405", "190x450", "215x405", "215x450"],
    "Massivtre": ["100x300", "120x300", "120x400", "140x400", "160x400", "200x400"],
    "Stål": ["KFHUP 120x120x8", "KFHUP 140x140x10", "KFHUP 160x160x10", "KFHUP 180x180x12.5", "KFHUP 200x200x12.5", "KFHUP 220x220x12.5"],
    "Betong": ["200x200", "250x250", "300x300", "350x350", "400x400"],
}


def safe_num(value) -> float:
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def clean_dataframe(df: pd.DataFrame, required_cols=None) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")
    if required_cols:
        for col in required_cols:
            if col in df.columns:
                df = df[df[col].notna()]
    return df.reset_index(drop=True)


def metric_card(title, value):
    st.markdown(f"""
        <div class="custom-card">
            <div class="custom-title">{title}</div>
            <div class="custom-value">{value}</div>
        </div>
    """, unsafe_allow_html=True)


def file_hash(file_bytes: bytes) -> str:
    return hashlib.md5(file_bytes).hexdigest()


def detect_stake_columns(df: pd.DataFrame) -> dict:
    col_map = {}
    normalized = {str(c).strip().lower(): c for c in df.columns}

    candidates = {
        "x": ["x", "øst", "easting", "east", "x-koordinat", "x_koordinat"],
        "y": ["y", "nord", "northing", "north", "y-koordinat", "y_koordinat"],
        "z": ["z", "høyde", "hoyde", "elev", "elevation", "kote", "z-koordinat", "z_koordinat"],
        "kode": ["kode", "code", "type", "punktkode", "pointcode", "point_code"],
        "punkt": ["punkt", "point", "punktnr", "pointid", "id", "nr"],
    }

    for key, names in candidates.items():
        for name in names:
            if name in normalized:
                col_map[key] = normalized[name]
                break

    if not {"x", "y", "z"}.issubset(col_map.keys()):
        numeric_cols = []
        for c in df.columns:
            s = pd.to_numeric(df[c], errors="coerce")
            if s.notna().sum() >= max(3, int(len(df) * 0.5)):
                numeric_cols.append(c)
        if len(numeric_cols) >= 3:
            if "x" not in col_map:
                col_map["x"] = numeric_cols[0]
            if "y" not in col_map:
                col_map["y"] = numeric_cols[1]
            if "z" not in col_map:
                col_map["z"] = numeric_cols[2]

    return col_map


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

BREEAM_LEVELS = ["Ingen", "Pass", "Good", "Very Good", "Excellent", "Outstanding"]


def get_breeam_config(level: str) -> dict:
    level = level or "Ingen"
    config = {
        "level": level,
        "deck_variant": st.session_state.get("deck_variant_key", "Hulldekke"),
        "concrete_variant": st.session_state.get("concrete_variant_key", "Plasstøpt_betong"),
        "wall_variant": st.session_state.get("wall_variant_key", "Betong_vegg"),
        "ground_multiplier": 1.0,
        "reuse_bonus": 0.0,
        "stormwater_rate": 0.0,
        "documentation_pct": 0.0,
        "waste_sorting_pct": 0.0,
        "notes": ["Ingen aktiv BREEAM-scenario."],
    }
    if level == "Pass":
        config.update({"use_epd": True, "notes": ["EPD/prosjektfaktorer brukes som hovedgrunnlag.", "Enkel miljøoppfølging i byggeplassfasen."]})
    elif level == "Good":
        config.update({
            "use_epd": True,
            "concrete_variant": "Plasstøpt_betong_lavCO2",
            "wall_variant": "Betong_vegg_lavCO2",
            "ground_multiplier": 1.02,
            "reuse_bonus": 0.05,
            "documentation_pct": 0.01,
            "notes": ["Lavkarbon betong anbefales i dekker/vegger.", "Enklere miljøoppfølgingsplan og avfallsplan legges inn."]
        })
    elif level == "Very Good":
        config.update({
            "use_epd": True,
            "deck_variant": "Hulldekke_lavCO2",
            "concrete_variant": "Plasstøpt_betong_lavCO2",
            "wall_variant": "Betong_vegg_lavCO2",
            "ground_multiplier": 1.04,
            "reuse_bonus": 0.10,
            "stormwater_rate": 65.0,
            "documentation_pct": 0.015,
            "waste_sorting_pct": 0.005,
            "notes": ["Lavkarbonløsninger aktiveres i betongbaserte bygningsdeler.", "Overvannstiltak og økt masseseparering anbefales i grunnarbeidene.", "Miljøoppfølging og dokumentasjon skjerpes."]
        })
    elif level == "Excellent":
        config.update({
            "use_epd": True,
            "deck_variant": "Hulldekke_lavCO2",
            "concrete_variant": "Plasstøpt_betong_lavCO2",
            "wall_variant": "Betong_vegg_lavCO2",
            "ground_multiplier": 1.07,
            "reuse_bonus": 0.18,
            "stormwater_rate": 95.0,
            "documentation_pct": 0.02,
            "waste_sorting_pct": 0.01,
            "notes": ["Scenarioet antar mer omfattende overvannshåndtering og masseregnskap.", "Høyere krav til dokumentasjon, avfallssortering og logistikk.", "Materialvalg styres mot lavkarbon der det finnes i modellen."]
        })
    elif level == "Outstanding":
        config.update({
            "use_epd": True,
            "deck_variant": "Hulldekke_lavCO2",
            "concrete_variant": "Plasstøpt_betong_lavCO2",
            "wall_variant": "Betong_vegg_lavCO2",
            "ground_multiplier": 1.10,
            "reuse_bonus": 0.25,
            "stormwater_rate": 125.0,
            "documentation_pct": 0.03,
            "waste_sorting_pct": 0.015,
            "notes": ["Scenarioet legger inn svært streng miljøstyring og overvannstiltak.", "Masser gjenbrukes i størst mulig grad, og dokumentasjonskostnader øker.", "Modellen peker mot ytterligere materialsubstitusjon, f.eks. mer tre der det er mulig."]
        })
    return config




def build_ground_pricing_basis_v2(summary: dict, system_key: str = "Standard byggegrop", breeam_level: str = "Ingen") -> pd.DataFrame:
    base = GROUND_SYSTEM_LIBRARY.get(system_key, GROUND_SYSTEM_LIBRARY["Standard byggegrop"]).copy()
    breeam_cfg = get_breeam_config(breeam_level)
    area = safe_num(summary.get("Tomteareal (konveks hull)", 0))
    cut = safe_num(summary.get("Estimert utgraving", 0))
    fill = safe_num(summary.get("Estimert oppfylling", 0))
    stripping_volume = area * base["stripping_depth"]
    reuse_amount = min(fill, cut * (base.get("reuse_factor", 0.0) + breeam_cfg.get("reuse_bonus", 0.0)))
    imported_fill = max(fill - reuse_amount, 0.0)
    exported_cut = max(cut - reuse_amount, 0.0)
    subbase_volume = area * base.get("subbase_thickness", 0.0)
    env_multiplier = breeam_cfg.get("ground_multiplier", 1.0)
    stormwater_rate = max(base.get("stormwater_rate", 0.0), breeam_cfg.get("stormwater_rate", 0.0))

    rows = [
        {"Post": "Rigg og drift", "Enhet": "RS", "Mengde": 1.0, "Enhetspris": 1.0, "Beløp": 0.0, "Merknad": f"{system_key}. Rigg beregnes av variable poster."},
        {"Post": "Avdekking / matjord", "Enhet": "m3", "Mengde": stripping_volume, "Enhetspris": base["stripping_rate"] * env_multiplier, "Beløp": stripping_volume * base["stripping_rate"] * env_multiplier, "Merknad": f"Antatt dybde {base['stripping_depth']:.2f} m"},
        {"Post": "Utgraving", "Enhet": "m3", "Mengde": cut, "Enhetspris": base["excavation_rate"] * env_multiplier, "Beløp": cut * base["excavation_rate"] * env_multiplier, "Merknad": "Basert på stikningspunkter mot valgt prosjektkote"},
        {"Post": "Intern gjenbruk av masser", "Enhet": "m3", "Mengde": reuse_amount, "Enhetspris": max(base["fill_rate"] * 0.45, 1.0), "Beløp": reuse_amount * max(base["fill_rate"] * 0.45, 1.0), "Merknad": "Masser som kan omdisponeres på tomta"},
        {"Post": "Bortkjøring / deponi", "Enhet": "m3", "Mengde": exported_cut, "Enhetspris": base["transport_cut_rate"] * env_multiplier, "Beløp": exported_cut * base["transport_cut_rate"] * env_multiplier, "Merknad": "Overskuddsmasser til depot/deponi"},
        {"Post": "Oppfylling / komprimering", "Enhet": "m3", "Mengde": imported_fill, "Enhetspris": base["fill_rate"] * env_multiplier, "Beløp": imported_fill * base["fill_rate"] * env_multiplier, "Merknad": "Netto importerte fyllmasser"},
        {"Post": "Tilkjøring av fyllmasser", "Enhet": "m3", "Mengde": imported_fill, "Enhetspris": base["import_fill_rate"] * env_multiplier, "Beløp": imported_fill * base["import_fill_rate"] * env_multiplier, "Merknad": "Transport og logistikk for eksterne masser"},
        {"Post": "Geotekstil / separasjonslag", "Enhet": "m2", "Mengde": area, "Enhetspris": base["geotextile_rate"] * env_multiplier, "Beløp": area * base["geotextile_rate"] * env_multiplier, "Merknad": "Lik tomteareal som første anslag"},
        {"Post": "Forsterkningslag / bærelag", "Enhet": "m3", "Mengde": subbase_volume, "Enhetspris": base["subbase_rate"] * env_multiplier, "Beløp": subbase_volume * base["subbase_rate"] * env_multiplier, "Merknad": f"Antatt tykkelse {base.get('subbase_thickness', 0.0):.2f} m"},
        {"Post": "Drenering", "Enhet": "m2", "Mengde": area, "Enhetspris": base["drain_rate"] * 0.10 * env_multiplier, "Beløp": area * base["drain_rate"] * 0.10 * env_multiplier, "Merknad": "Forenklet post for drenering/perimeter"},
    ]
    if stormwater_rate > 0:
        rows.append({"Post": "Overvannstiltak / blågrønne løsninger", "Enhet": "m2", "Mengde": area, "Enhetspris": stormwater_rate, "Beløp": area * stormwater_rate, "Merknad": f"Aktivert av BREEAM {breeam_level}"})
    df = pd.DataFrame(rows)
    variable_sum = df.loc[df["Post"] != "Rigg og drift", "Beløp"].sum()
    df.loc[df["Post"] == "Rigg og drift", "Beløp"] = variable_sum * base["rigg_pct"]
    documentation_pct = base.get("documentation_pct", 0.0) + breeam_cfg.get("documentation_pct", 0.0) + breeam_cfg.get("waste_sorting_pct", 0.0)
    if documentation_pct > 0:
        env_base = df["Beløp"].sum()
        df = pd.concat([df, pd.DataFrame([{
            "Post": "Miljøoppfølging / dokumentasjon", "Enhet": "RS", "Mengde": 1.0, "Enhetspris": env_base * documentation_pct,
            "Beløp": env_base * documentation_pct, "Merknad": f"BREEAM {breeam_level} miljøoppfølging og avfallslogistikk"
        }])], ignore_index=True)
    return df


def compare_ground_scenarios(summary: dict, current_system: str, target_system: str, breeam_level: str = "Ingen") -> pd.DataFrame:
    current_df = build_ground_pricing_basis_v2(summary, current_system, breeam_level)
    target_df = build_ground_pricing_basis_v2(summary, target_system, breeam_level)
    current_total = float(current_df["Beløp"].sum())
    target_total = float(target_df["Beløp"].sum())
    return pd.DataFrame([
        {"Scenario": current_system, "Estimert kostnad [kr]": current_total},
        {"Scenario": target_system, "Estimert kostnad [kr]": target_total},
        {"Scenario": "Endring", "Estimert kostnad [kr]": target_total - current_total},
    ])


def load_stake_data(file) -> pd.DataFrame:
    suffix = Path(file.name).suffix.lower()
    if suffix == ".csv":
        raw = pd.read_csv(file)
    elif suffix in [".xlsx", ".xls"]:
        raw = pd.read_excel(file)
    elif suffix in [".txt", ".pts"]:
        raw = pd.read_csv(file, sep=None, engine="python")
    else:
        raise ValueError("Støttet filformat er CSV, XLSX, XLS, TXT eller PTS.")

    raw = clean_dataframe(raw)
    col_map = detect_stake_columns(raw)
    if not {"x", "y", "z"}.issubset(col_map.keys()):
        raise ValueError("Fant ikke X, Y og Z i stikningsfilen. Gi kolonnene navn som X, Y, Z / Øst, Nord, Høyde.")

    out = pd.DataFrame({
        "Punkt": raw[col_map.get("punkt", raw.columns[0])].astype(str) if col_map.get("punkt") in raw.columns else [f"P{i+1}" for i in range(len(raw))],
        "X": pd.to_numeric(raw[col_map["x"]], errors="coerce"),
        "Y": pd.to_numeric(raw[col_map["y"]], errors="coerce"),
        "Z": pd.to_numeric(raw[col_map["z"]], errors="coerce"),
        "Kode": raw[col_map["kode"]].astype(str) if col_map.get("kode") in raw.columns else "Ukjent",
    }).dropna(subset=["X", "Y", "Z"])

    if out.empty:
        raise ValueError("Ingen gyldige punkt ble funnet i stikningsfilen.")
    return out.reset_index(drop=True)


def _cross(o, a, b):
    return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])


def convex_hull(points):
    pts = sorted(set((float(x), float(y)) for x, y in points))
    if len(pts) <= 1:
        return pts
    lower = []
    for p in pts:
        while len(lower) >= 2 and _cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    upper = []
    for p in reversed(pts):
        while len(upper) >= 2 and _cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]


def polygon_area(poly):
    if len(poly) < 3:
        return 0.0
    area = 0.0
    for i in range(len(poly)):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % len(poly)]
        area += x1 * y2 - x2 * y1
    return abs(area) / 2.0


def build_ground_summary(points_df: pd.DataFrame, target_elevation: float | None = None, mass_factor: float = 1.15):
    hull = convex_hull(points_df[["X", "Y"]].itertuples(index=False, name=None))
    hull_area = polygon_area(hull)
    x_min, x_max = points_df["X"].min(), points_df["X"].max()
    y_min, y_max = points_df["Y"].min(), points_df["Y"].max()
    z_min, z_max = points_df["Z"].min(), points_df["Z"].max()
    z_mean = points_df["Z"].mean()
    target = float(target_elevation) if target_elevation is not None else float(z_mean)
    delta = points_df["Z"] - target
    cut_depth = delta.clip(lower=0)
    fill_depth = (-delta).clip(lower=0)
    point_density = len(points_df) / hull_area if hull_area > 0 else 0.0
    sample_area = hull_area / len(points_df) if len(points_df) > 0 else 0.0
    cut_volume = cut_depth.sum() * sample_area * mass_factor
    fill_volume = fill_depth.sum() * sample_area * mass_factor
    avg_spacing = math.sqrt(sample_area) if sample_area > 0 else math.nan

    summary = {
        "Antall punkt": int(len(points_df)),
        "Tomteareal (konveks hull)": hull_area,
        "Utbredelse X": x_max - x_min,
        "Utbredelse Y": y_max - y_min,
        "Laveste kote": z_min,
        "Høyeste kote": z_max,
        "Middelkote": z_mean,
        "Prosjektkote": target,
        "Estimert utgraving": cut_volume,
        "Estimert oppfylling": fill_volume,
        "Punkttetthet": point_density,
        "Punktavstand ca.": avg_spacing,
    }

    points_out = points_df.copy()
    points_out["Avvik fra prosjektkote [m]"] = points_out["Z"] - target
    points_out["Skjæring [m]"] = cut_depth
    points_out["Fylling [m]"] = fill_depth
    return summary, points_out, hull


def build_ground_pricing_basis(summary: dict, rigg_pct: float = 0.08, excavation_rate: float = 185.0, fill_rate: float = 145.0, geotextile_rate: float = 32.0, stripping_depth: float = 0.3, stripping_rate: float = 95.0):
    return build_ground_pricing_basis_v2(summary, system_key="Standard byggegrop", breeam_level=st.session_state.get("breeam_target_level", "Ingen") if st.session_state.get("breeam_active", False) else "Ingen")


def generate_ground_obj(points_df: pd.DataFrame) -> bytes:
    try:
        import matplotlib.tri as mtri
    except Exception as e:
        raise RuntimeError(f"Kunne ikke laste triangulering: {e}")
    tri = mtri.Triangulation(points_df["X"].to_numpy(), points_df["Y"].to_numpy())
    lines = ["# byggTotal terrengmodell"]
    for row in points_df.itertuples(index=False):
        lines.append(f"v {row.X:.4f} {row.Y:.4f} {row.Z:.4f}")
    for a, b, c in tri.triangles:
        lines.append(f"f {a+1} {b+1} {c+1}")
    return "\n".join(lines).encode("utf-8")


def plot_ground_points(points_df: pd.DataFrame, hull=None):
    fig, ax = plt.subplots(figsize=(7, 5))
    sc = ax.scatter(points_df["X"], points_df["Y"], c=points_df["Z"], s=20)
    if hull and len(hull) >= 3:
        hx = [p[0] for p in hull] + [hull[0][0]]
        hy = [p[1] for p in hull] + [hull[0][1]]
        ax.plot(hx, hy, linewidth=1.5)
    ax.set_xlabel("X")
    ax.set_ylabel("Y")
    ax.set_title("Stikningspunkter / tomteutbredelse")
    ax.axis("equal")
    fig.colorbar(sc, ax=ax, label="Z / kote")
    return fig


@st.cache_data(show_spinner=False)
def load_sheet_df(file_bytes: bytes, sheet_name: str, data_only: bool = True) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl")


def classify_material(material_text):
    text = str(material_text or "").lower()
    if any(x in text for x in ["stål", "steel", "s355", "s235", "kfh", "vfh", "rhs", "shs", "hea", "heb", "ipe", "hup", "kfhu"]):
        return "Stål"
    if any(x in text for x in ["limtre", "glulam", "glt"]):
        return "Limtre"
    if any(x in text for x in ["massivtre", "clt", "cross laminated timber", "krysslaminert"]):
        return "Massivtre"
    if any(x in text for x in ["hulldekke", "hd"]):
        return "Betong"
    if any(x in text for x in ["betong", "concrete", "in-situ", "cast in place", "prefab concrete"]):
        return "Betong"
    if any(x in text for x in ["wood", "tre", "timber"]):
        return "Tre"
    return "Ukjent"


def parse_profile(profile: str):
    text = str(profile or "")
    material = classify_material(text)
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+[\.,]?\d*", text)]
    area_m2 = None
    width_mm = height_mm = thickness_mm = None

    if material == "Stål" and len(nums) >= 3:
        width_mm, height_mm, thickness_mm = nums[-3], nums[-2], nums[-1]
        inner_w = max(width_mm - 2 * thickness_mm, 0)
        inner_h = max(height_mm - 2 * thickness_mm, 0)
        area_mm2 = (width_mm * height_mm) - (inner_w * inner_h)
        area_m2 = area_mm2 / 1_000_000
    elif material in ["Limtre", "Massivtre", "Tre", "Betong"] and len(nums) >= 2:
        width_mm, height_mm = nums[-2], nums[-1]
        area_mm2 = width_mm * height_mm
        area_m2 = area_mm2 / 1_000_000

    return {
        "materiale": material,
        "bredde_mm": width_mm,
        "høyde_mm": height_mm,
        "tykkelse_mm": thickness_mm,
        "areal_m2": area_m2,
    }


def parse_profile_area_from_text(profile_text: str, material_hint: str = "") -> float:
    text = str(profile_text or "")
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+[\.,]?\d*", text)]
    if len(nums) < 2:
        return math.nan

    material_guess = classify_material(material_hint if material_hint else text)
    lower_text = text.lower()

    if material_guess == "Stål" or any(x in lower_text for x in ["kfh", "rhs", "shs", "hup"]):
        if len(nums) >= 3:
            width_mm, height_mm, thickness_mm = nums[-3], nums[-2], nums[-1]
            inner_w = max(width_mm - 2 * thickness_mm, 0)
            inner_h = max(height_mm - 2 * thickness_mm, 0)
            area_mm2 = (width_mm * height_mm) - (inner_w * inner_h)
            return area_mm2 / 1_000_000
        return math.nan

    width_mm, height_mm = nums[-2], nums[-1]
    return (width_mm * height_mm) / 1_000_000


def map_ifc_type(entity_name):
    mapping = {
        "IfcBeam": "Bjelke",
        "IfcColumn": "Søyle",
        "IfcSlab": "Dekke",
        "IfcWall": "Vegg",
        "IfcWallStandardCase": "Vegg",
        "IfcRoof": "Tak",
        "IfcMember": "Medlem",
        "IfcFooting": "Fundament",
    }
    return mapping.get(entity_name, entity_name.replace("Ifc", ""))


def material_color(materiale: str, is_changed: bool = False):
    if is_changed:
        return "#ff66cc"
    mapping = {
        "Stål": "#4F81BD", "Betong": "#A6A6A6", "Limtre": "#C58C4B",
        "Massivtre": "#8CBF3F", "Tre": "#B97A57", "Ukjent": "#D9D9D9",
    }
    return mapping.get(materiale, "#D9D9D9")


def detect_product_key(row, deck_variant, concrete_variant, wall_variant):
    row_type = str(row.get("Type", "") or "")
    materiale = str(row.get("materiale", "") or "")
    profile = str(row.get("Material / Tverrsnitt", "") or "").lower()

    if row_type == "Vegg":
        if materiale == "Massivtre":
            return "Massivtre_vegg"
        if materiale == "Betong":
            return wall_variant
        return materiale if materiale in MATERIAL_DATABASE else "Ukjent"

    if row_type == "Dekke":
        if "hulldekke" in profile or re.search(r"\bhd\b", profile):
            return deck_variant
        if materiale == "Betong":
            return concrete_variant
        return materiale if materiale in MATERIAL_DATABASE else "Ukjent"

    if materiale == "Stål":
        return "Stål"
    if materiale == "Limtre":
        return "Limtre"
    if materiale == "Massivtre":
        return "Massivtre"
    if materiale == "Tre":
        return "Tre"
    if materiale == "Betong":
        return "Betong_volum"
    return "Ukjent"


def get_quantity_for_product(row, product_key):
    product = MATERIAL_DATABASE.get(product_key, MATERIAL_DATABASE["Ukjent"])
    unit = product["unit"]
    if unit == "kg":
        return safe_num(row.get("Vekt [kg]", 0))
    if unit == "m3":
        return safe_num(row.get("Volum [m3]", 0))
    if unit == "m2":
        return safe_num(row.get("Areal [m2]", 0))
    return 0.0


def cost_for_row(row, deck_variant, concrete_variant, wall_variant):
    key = detect_product_key(row, deck_variant, concrete_variant, wall_variant)
    return get_quantity_for_product(row, key) * MATERIAL_DATABASE.get(key, MATERIAL_DATABASE["Ukjent"])["price"]


def co2_for_row(row, deck_variant, concrete_variant, wall_variant, use_epd=True):
    key = detect_product_key(row, deck_variant, concrete_variant, wall_variant)
    qty = get_quantity_for_product(row, key)
    if use_epd and key in EPD_DATABASE:
        return qty * EPD_DATABASE[key]["co2"]
    return qty * MATERIAL_DATABASE.get(key, MATERIAL_DATABASE["Ukjent"])["co2"]


def map_ns3420_code(row) -> str:
    row_type = str(row.get("Type", "") or "")
    material = str(row.get("materiale", "") or "")
    profile = str(row.get("Material / Tverrsnitt", "") or "").lower()

    if row_type == "Søyle" and material == "Stål":
        return "NS3420: K / stålsøyle"
    if row_type == "Bjelke" and material == "Stål":
        return "NS3420: K / stålbjelke"
    if row_type == "Søyle" and material == "Limtre":
        return "NS3420: K / limtresøyle"
    if row_type == "Bjelke" and material == "Limtre":
        return "NS3420: K / limtrebjelke"
    if row_type == "Dekke" and "hulldekke" in profile:
        return "NS3420: L / hulldekke"
    if row_type == "Dekke" and material == "Betong":
        return "NS3420: L / betongdekke"
    if row_type == "Vegg" and material == "Betong":
        return "NS3420: M / betongvegg"
    if row_type == "Vegg" and material == "Massivtre":
        return "NS3420: M / massivtrevegg"
    if row_type == "Fundament":
        return "NS3420: L / fundament"
    return "NS3420: ikke klassifisert"


def build_dataset_from_excel(file_bytes: bytes):
    mengder = clean_dataframe(load_sheet_df(file_bytes, "MENGDER"), ["Segment"])
    segmenter = clean_dataframe(load_sheet_df(file_bytes, "Segmenter"), ["Navn"])
    knutepunkter = clean_dataframe(load_sheet_df(file_bytes, "Knutepunkter"), ["Navn"])
    forside = clean_dataframe(load_sheet_df(file_bytes, "FORSIDE"))

    merged = mengder.merge(
        segmenter[["Navn", "Material / Tverrsnitt"]],
        left_on="Segment",
        right_on="Navn",
        how="left",
    ).drop(columns=[c for c in ["Navn"] if c in mengder.columns], errors="ignore")

    profile_df = merged["Material / Tverrsnitt"].apply(parse_profile).apply(pd.Series)
    merged = pd.concat([merged, profile_df], axis=1)

    if "Lengde [m]" not in merged.columns:
        merged["Lengde [m]"] = math.nan
    if "Areal [m2]" not in merged.columns:
        merged["Areal [m2]"] = math.nan
    for col in ["Lengde [m]", "Areal [m2]", "Volum [m3]"]:
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce")

    if "Volum [m3]" not in merged.columns:
        merged["Volum [m3]"] = merged["Lengde [m]"] * merged["areal_m2"]

    def calc_weight(row):
        volume = safe_num(row.get("Volum [m3]"))
        mat = row.get("materiale")
        if mat == "Stål":
            return volume * STEEL_DENSITY
        if mat == "Limtre":
            return volume * GLULAM_DENSITY
        if mat == "Massivtre":
            return volume * CLT_DENSITY
        if mat == "Betong":
            return volume * CONCRETE_DENSITY
        if mat == "Tre":
            return volume * TIMBER_DENSITY
        return math.nan

    merged["Vekt [kg]"] = merged.apply(calc_weight, axis=1)
    merged["Mengdegrunnlag"] = merged.apply(
        lambda row: "Excel" if any(pd.notna(row.get(c)) and safe_num(row.get(c)) > 0 for c in ["Lengde [m]", "Areal [m2]", "Volum [m3]"]) else "Manglende mengder",
        axis=1,
    )
    merged["Endret IFC"] = False
    return merged, knutepunkter, forside


def get_ifc_material_name(element):
    if ifc_element_util is None:
        return "Ukjent"
    try:
        material = ifc_element_util.get_material(element, should_skip_usage=True)
        if material is None:
            return "Ukjent"
        if hasattr(material, "Name") and material.Name:
            return str(material.Name)
        if material.is_a("IfcMaterialLayerSet"):
            names = [layer.Material.Name for layer in material.MaterialLayers if layer.Material]
            return ", ".join([n for n in names if n]) or "Ukjent"
        if material.is_a("IfcMaterialProfileSet"):
            names = []
            for prof in material.MaterialProfiles:
                if getattr(prof, "Material", None) and prof.Material.Name:
                    names.append(prof.Material.Name)
            return ", ".join(names) or "Ukjent"
    except Exception:
        pass
    return "Ukjent"


def get_property_from_pset(element, pset_name: str, prop_name: str):
    try:
        for rel in getattr(element, "IsDefinedBy", []) or []:
            pdef = getattr(rel, "RelatingPropertyDefinition", None)
            if not pdef or not pdef.is_a("IfcPropertySet"):
                continue
            if getattr(pdef, "Name", "") != pset_name:
                continue
            for prop in getattr(pdef, "HasProperties", []) or []:
                if getattr(prop, "Name", "") == prop_name:
                    nominal = getattr(prop, "NominalValue", None)
                    if nominal is None:
                        return None
                    return getattr(nominal, "wrappedValue", nominal)
    except Exception:
        pass
    return None


def is_ifc_element_changed(element) -> bool:
    return bool(get_property_from_pset(element, BYGGTOTAL_PSET_NAME, BYGGTOTAL_CHANGED_PROP))


def get_ifc_quantity_smart(element):
    quantity_map = {
        "length": ["Length", "NetLength", "GrossLength", "Height", "Depth", "OverallLength"],
        "area": ["Area", "NetArea", "GrossArea", "NetSideArea", "GrossSideArea", "OuterSurfaceArea", "FootprintArea", "CrossSectionArea"],
        "volume": ["Volume", "NetVolume", "GrossVolume"],
        "weight": ["Weight", "GrossWeight", "NetWeight"],
    }
    results = {"length": None, "area": None, "volume": None, "weight": None}
    try:
        for rel in getattr(element, "IsDefinedBy", []) or []:
            definition = getattr(rel, "RelatingPropertyDefinition", None)
            if not definition or not definition.is_a("IfcElementQuantity"):
                continue
            for qty in getattr(definition, "Quantities", []) or []:
                qname = getattr(qty, "Name", "")
                for key, names in quantity_map.items():
                    if qname not in names:
                        continue
                    for attr in ["LengthValue", "AreaValue", "VolumeValue", "WeightValue"]:
                        if hasattr(qty, attr):
                            val = getattr(qty, attr)
                            if val is not None and safe_num(val) > 0:
                                results[key] = val
    except Exception:
        pass
    return results


def estimate_dimensions_from_mesh(verts):
    if not verts:
        return None
    x = verts[0::3]
    y = verts[1::3]
    z = verts[2::3]
    if not x or not y or not z:
        return None
    dims = sorted([abs(max(x) - min(x)), abs(max(y) - min(y)), abs(max(z) - min(z))], reverse=True)
    return dims


def estimate_quantities_from_geometry(element, settings):
    try:
        shape = ifcopenshell.geom.create_shape(settings, element)
        geom = shape.geometry
        dims = estimate_dimensions_from_mesh(geom.verts)
        if not dims:
            return {"length": None, "area": None, "volume": None, "weight": None, "method": None}
        d1, d2, d3 = dims
        return {
            "length": d1 if d1 > 0 else None,
            "area": d1 * d2 if d1 * d2 > 0 else None,
            "volume": d1 * d2 * d3 if d1 * d2 * d3 > 0 else None,
            "weight": None,
            "method": "Geometriestimat",
        }
    except Exception:
        return {"length": None, "area": None, "volume": None, "weight": None, "method": None}


@st.cache_data(show_spinner="Leser IFC og bygger datasett...")
def build_dataset_from_ifc(ifc_bytes: bytes, use_geometry_fallback: bool = True, fast_mode: bool = False):
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert.")
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp:
            tmp.write(ifc_bytes)
            temp_path = tmp.name
        model = ifcopenshell.open(temp_path)
        settings = ifcopenshell.geom.settings()
        settings.set(settings.USE_WORLD_COORDS, True)
        rows = []
        for type_name in SUPPORTED_IFC_TYPES:
            for el in model.by_type(type_name):
                global_id = getattr(el, "GlobalId", None)
                name = getattr(el, "Name", None) or global_id or "Ukjent"
                object_type = getattr(el, "ObjectType", None) or ""
                predefined = getattr(el, "PredefinedType", None) or ""
                material_raw = get_ifc_material_name(el)
                materiale = classify_material(material_raw)
                changed_flag = is_ifc_element_changed(el)
                q = get_ifc_quantity_smart(el)
                length_m = pd.to_numeric(q["length"], errors="coerce")
                area_m2 = pd.to_numeric(q["area"], errors="coerce")
                volume_m3 = pd.to_numeric(q["volume"], errors="coerce")
                weight_kg = pd.to_numeric(q["weight"], errors="coerce")
                quantity_method = "IFC quantities"

                should_use_geometry_fallback = (
                    use_geometry_fallback
                    and not fast_mode
                    and type_name in IFC_GEOMETRY_FALLBACK_TYPES
                    and all(pd.isna(v) or safe_num(v) == 0 for v in [length_m, area_m2, volume_m3])
                )

                if should_use_geometry_fallback:
                    geo_q = estimate_quantities_from_geometry(el, settings)
                    length_m = pd.to_numeric(geo_q["length"], errors="coerce")
                    area_m2 = pd.to_numeric(geo_q["area"], errors="coerce")
                    volume_m3 = pd.to_numeric(geo_q["volume"], errors="coerce")
                    quantity_method = geo_q["method"] or "Ikke funnet"

                if (pd.isna(area_m2) or safe_num(area_m2) == 0) and pd.notna(volume_m3) and pd.notna(length_m) and safe_num(length_m) > 0:
                    area_m2 = safe_num(volume_m3) / safe_num(length_m)

                if pd.notna(weight_kg) and safe_num(weight_kg) > 0:
                    vekt_kg = weight_kg
                elif materiale == "Stål":
                    vekt_kg = safe_num(volume_m3) * STEEL_DENSITY
                elif materiale == "Limtre":
                    vekt_kg = safe_num(volume_m3) * GLULAM_DENSITY
                elif materiale == "Massivtre":
                    vekt_kg = safe_num(volume_m3) * CLT_DENSITY
                elif materiale == "Betong":
                    vekt_kg = safe_num(volume_m3) * CONCRETE_DENSITY
                elif materiale == "Tre":
                    vekt_kg = safe_num(volume_m3) * TIMBER_DENSITY
                else:
                    vekt_kg = math.nan

                rows.append({
                    "Segment": name,
                    "Type": map_ifc_type(type_name),
                    "Knutepunkter": "",
                    "Material / Tverrsnitt": object_type if object_type else predefined,
                    "Lengde [m]": length_m,
                    "Areal [m2]": area_m2,
                    "Volum [m3]": volume_m3,
                    "Vekt [kg]": vekt_kg,
                    "materiale": materiale,
                    "IFC Type": type_name,
                    "IFC GlobalId": global_id,
                    "Kilde": "IFC",
                    "Mengdegrunnlag": quantity_method,
                    "Endret IFC": changed_flag,
                })

        data = pd.DataFrame(rows)
        nodes = pd.DataFrame()
        forside = pd.DataFrame([["Kilde", "IFC"], ["Antall elementer", len(data)], ["Filtype", "IFC"]], columns=["Parameter", "Verdi"])
        if data.empty:
            raise ValueError("Fant ingen støttede elementer i IFC-filen.")
        return data, nodes, forside
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


@st.cache_data(show_spinner=False)
def extract_ifc_meshes_filtered(ifc_bytes: bytes, visible_ids_tuple=None, max_elements=1000):
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert.")
    temp_path = None
    meshes = []
    visible_ids = set(visible_ids_tuple) if visible_ids_tuple not in [None, (), []] else None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp:
            tmp.write(ifc_bytes)
            temp_path = tmp.name
        model = ifcopenshell.open(temp_path)
        settings = ifcopenshell.geom.settings()
        settings.set(settings.USE_WORLD_COORDS, True)
        count = 0
        for type_name in SUPPORTED_IFC_TYPES:
            for el in model.by_type(type_name):
                gid = getattr(el, "GlobalId", "")
                if visible_ids is not None and gid not in visible_ids:
                    continue
                try:
                    shape = ifcopenshell.geom.create_shape(settings, el)
                    geom = shape.geometry
                    verts = geom.verts
                    faces = geom.faces
                    if not verts or not faces:
                        continue
                    meshes.append({
                        "global_id": gid,
                        "name": getattr(el, "Name", "") or gid or "Ukjent",
                        "type": map_ifc_type(type_name),
                        "ifc_type": type_name,
                        "materiale": classify_material(get_ifc_material_name(el)),
                        "changed": is_ifc_element_changed(el),
                        "x": verts[0::3],
                        "y": verts[1::3],
                        "z": verts[2::3],
                        "i": faces[0::3],
                        "j": faces[1::3],
                        "k": faces[2::3],
                    })
                    count += 1
                    if count >= max_elements:
                        return meshes
                except Exception:
                    continue
        return meshes
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def build_ifc_3d_figure(meshes, preview_ids=None, show_only_preview=False, preview_material=None):
    preview_ids = set(preview_ids or [])
    fig = go.Figure()
    for mesh in meshes:
        is_preview = mesh["global_id"] in preview_ids
        if show_only_preview and preview_ids and not is_preview:
            continue
        if is_preview:
            color = "#ff66cc"
            opacity = 1.0
            display_material = preview_material if preview_material else f"{mesh['materiale']} → ny"
            preview_text = "Ja"
        else:
            color = material_color(mesh["materiale"], mesh.get("changed", False))
            opacity = 0.95 if not preview_ids else 0.12
            display_material = mesh["materiale"]
            preview_text = "Nei"
        fig.add_trace(go.Mesh3d(
            x=mesh["x"], y=mesh["y"], z=mesh["z"], i=mesh["i"], j=mesh["j"], k=mesh["k"],
            color=color, opacity=opacity, flatshading=True,
            name=f"{mesh['type']} – {display_material}",
            hovertext=(
                f"Navn: {mesh['name']}<br>Type: {mesh['type']}<br>IFC-type: {mesh['ifc_type']}<br>"
                f"Materiale: {mesh['materiale']}<br>Forhåndsvisning: {preview_text}<br>GlobalId: {mesh['global_id']}"
            ),
            hoverinfo="text", showscale=False
        ))
    fig.update_layout(
        margin=dict(l=0, r=0, t=20, b=0),
        scene=dict(xaxis_title="X", yaxis_title="Y", zaxis_title="Z", aspectmode="data", bgcolor="rgba(0,0,0,0)"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        height=760
    )
    return fig


def get_owner_history(model):
    owners = model.by_type("IfcOwnerHistory")
    return owners[0] if owners else None


def get_or_create_ifc_material(model, material_name: str):
    for mat in model.by_type("IfcMaterial"):
        if getattr(mat, "Name", None) == material_name:
            return mat
    return model.create_entity("IfcMaterial", Name=material_name)


def remove_direct_material_relations(model, element):
    rels_to_remove = []
    for rel in model.by_type("IfcRelAssociatesMaterial"):
        related_objects = getattr(rel, "RelatedObjects", None) or []
        if element in related_objects:
            if len(related_objects) <= 1:
                rels_to_remove.append(rel)
            else:
                rel.RelatedObjects = [obj for obj in related_objects if obj != element]
    for rel in rels_to_remove:
        try:
            model.remove(rel)
        except Exception:
            pass


def assign_simple_material_to_element(model, element, material_name: str):
    material_entity = get_or_create_ifc_material(model, material_name)
    owner_history = get_owner_history(model)
    remove_direct_material_relations(model, element)
    model.create_entity(
        "IfcRelAssociatesMaterial",
        GlobalId=ifcopenshell.guid.new(),
        OwnerHistory=owner_history,
        Name=f"Material assignment {material_name}",
        Description=None,
        RelatedObjects=[element],
        RelatingMaterial=material_entity,
    )


def _make_ifc_text(model, text: str):
    try:
        return model.create_entity("IfcText", str(text))
    except Exception:
        return str(text)


def _make_ifc_boolean(model, value: bool):
    try:
        return model.create_entity("IfcBoolean", bool(value))
    except Exception:
        return bool(value)


def _set_or_create_pset_property(model, element, pset_name: str, prop_name: str, value, prop_type="text"):
    owner_history = get_owner_history(model)
    existing_pset = None
    for rel in getattr(element, "IsDefinedBy", []) or []:
        pdef = getattr(rel, "RelatingPropertyDefinition", None)
        if pdef and pdef.is_a("IfcPropertySet") and getattr(pdef, "Name", "") == pset_name:
            existing_pset = pdef
            break

    nominal = _make_ifc_text(model, value) if prop_type == "text" else _make_ifc_boolean(model, bool(value))

    if existing_pset is None:
        prop = model.create_entity("IfcPropertySingleValue", Name=prop_name, Description=None, NominalValue=nominal, Unit=None)
        pset = model.create_entity("IfcPropertySet", GlobalId=ifcopenshell.guid.new(), OwnerHistory=owner_history, Name=pset_name, Description=None, HasProperties=[prop])
        model.create_entity("IfcRelDefinesByProperties", GlobalId=ifcopenshell.guid.new(), OwnerHistory=owner_history, Name=f"{pset_name} relation", Description=None, RelatedObjects=[element], RelatingPropertyDefinition=pset)
        return

    props = list(getattr(existing_pset, "HasProperties", []) or [])
    for prop in props:
        if getattr(prop, "Name", "") == prop_name:
            prop.NominalValue = nominal
            return
    props.append(model.create_entity("IfcPropertySingleValue", Name=prop_name, Description=None, NominalValue=nominal, Unit=None))
    existing_pset.HasProperties = props


def get_swap_target_options(selected_type: str):
    if selected_type in ["Søyle", "Bjelke"]:
        return ["Stål", "Limtre", "Betong"]
    if selected_type == "Vegg":
        return ["Betong_vegg_150", "Betong_vegg_180", "Betong_vegg_200", "Massivtre_vegg_100", "Massivtre_vegg_120", "Massivtre_vegg_140", "Massivtre_vegg_160", "Massivtre_vegg_200", "Massivtre_vegg_240"]
    if selected_type == "Dekke":
        return ["Plasstopt_dekke_180", "Plasstopt_dekke_200", "Plasstopt_dekke_220", "Plasstopt_dekke_250", "Plasstopt_dekke_300", "Plasstopt_dekke_350", "Plasstopt_dekke_lavCO2", "Hulldekke_200", "Hulldekke_220", "Hulldekke_265", "Hulldekke_265_lavCO2", "Hulldekke_290", "Hulldekke_320", "Hulldekke_320_lavCO2", "Hulldekke_340", "Hulldekke_400", "Hulldekke_420", "Hulldekke_500", "Massivtre_dekke_160", "Massivtre_dekke_180", "Massivtre_dekke_200", "Massivtre_dekke_220", "Massivtre_dekke_240", "Massivtre_dekke_260", "Massivtre_dekke_280"]
    return ["Stål", "Limtre", "Betong"]


def format_swap_target_option(option_key: str) -> str:
    if option_key in NORSK_PRISBOK_DATABASE:
        item = NORSK_PRISBOK_DATABASE[option_key]
        return f"{item['label']} ({item['npb_code']})"
    if option_key in MATERIAL_DATABASE:
        return MATERIAL_DATABASE[option_key]["label"]
    return option_key


def get_swap_target_defaults(target_key: str):
    if target_key in NORSK_PRISBOK_DATABASE:
        db = NORSK_PRISBOK_DATABASE[target_key]
        return {
            "density": 0.0, "price": db["price"], "price_unit": db["unit"], "co2": db["co2"], "label": db["label"],
            "target_key": target_key, "source": db["source"], "npb_code": db["npb_code"], "ak": db.get("ak", 0.0),
            "thickness_mm": db.get("thickness_mm"),
        }
    if target_key in MATERIAL_DATABASE:
        db = MATERIAL_DATABASE[target_key]
        return {
            "density": db.get("density", 0.0), "price": db.get("price", 0.0), "price_unit": db.get("unit", ""),
            "co2": EPD_DATABASE.get(target_key, {}).get("co2", db.get("co2", 0.0)), "label": db.get("label", target_key),
            "target_key": target_key, "source": "Materialdatabase", "npb_code": "", "ak": 0.0, "thickness_mm": None,
        }
    if target_key in ["Stål", "Limtre", "Massivtre", "Betong"]:
        base = "Betong_volum" if target_key == "Betong" else target_key
        db = MATERIAL_DATABASE[base]
        return {
            "density": db["density"], "price": db["price"], "price_unit": db["unit"], "co2": EPD_DATABASE.get(base, {}).get("co2", db["co2"]),
            "label": "Betong" if target_key == "Betong" else db["label"], "target_key": target_key,
            "source": "Materialdatabase", "npb_code": "", "ak": 0.0, "thickness_mm": None,
        }
    return {"density": 0.0, "price": 0.0, "price_unit": "", "co2": 0.0, "label": target_key, "target_key": target_key, "source": "Ukjent", "npb_code": "", "ak": 0.0, "thickness_mm": None}


def is_area_based_swap_target(target_key: str) -> bool:
    return target_key in NORSK_PRISBOK_DATABASE and NORSK_PRISBOK_DATABASE[target_key]["unit"] == "m2"


def infer_swap_length_for_row(row: pd.Series) -> float:
    length_m = safe_num(row.get("Lengde [m]", math.nan))
    old_volume_m3 = safe_num(row.get("Volum [m3]", math.nan))
    old_profile = row.get("Material / Tverrsnitt", "")
    old_material = row.get("materiale", "")
    if 0 < length_m <= 100:
        return length_m
    old_area_m2 = parse_profile_area_from_text(old_profile, old_material)
    if pd.notna(old_area_m2) and old_area_m2 > 0 and old_volume_m3 > 0:
        derived_length = old_volume_m3 / old_area_m2
        if 0 < derived_length <= 100:
            return derived_length
    return math.nan


def calculate_material_swap(source_df: pd.DataFrame, selected_type: str, from_material: str, target_key: str, new_profile_text: str):
    matched = source_df[(source_df["Type"] == selected_type) & (source_df["materiale"] == from_material)].copy()
    if matched.empty:
        return matched
    defaults = get_swap_target_defaults(target_key)
    matched["Gammelt materiale"] = matched["materiale"]
    matched["Nytt materiale"] = defaults["label"]
    matched["Nytt systemvalg"] = target_key
    matched["Nytt tverrsnitt"] = new_profile_text
    matched["Gammel kostnad [kr]"] = matched["Kostnad [kr]"]
    matched["Gammel vekt [kg]"] = matched["Vekt [kg]"]
    matched["Gammelt volum [m3]"] = matched["Volum [m3]"]
    matched["Gammel CO2 [kgCO2e]"] = matched["CO2 [kgCO2e]"]
    matched["Byttelengde [m]"] = matched.apply(infer_swap_length_for_row, axis=1)

    if is_area_based_swap_target(target_key):
        matched["Nytt volum [m3]"] = matched["Gammelt volum [m3]"]
        matched["Ny vekt [kg]"] = matched["Gammel vekt [kg]"]
        matched["Ny kostnad [kr]"] = matched["Areal [m2]"].fillna(0) * defaults["price"]
        matched["Ny CO2 [kgCO2e]"] = matched["Areal [m2]"].fillna(0) * defaults["co2"]
        matched["Byttemetode"] = "Areal × Norsk Prisbok-post"
        matched["Nytt tverrsnittsareal [m2]"] = math.nan
    else:
        new_area_m2 = parse_profile_area_from_text(new_profile_text, target_key)
        matched["Nytt volum [m3]"] = matched.apply(
            lambda row: safe_num(row["Byttelengde [m]"]) * new_area_m2 if pd.notna(row["Byttelengde [m]"]) and pd.notna(new_area_m2) and new_area_m2 > 0 else safe_num(row["Gammelt volum [m3]"]),
            axis=1
        )
        matched["Ny vekt [kg]"] = matched["Nytt volum [m3]"] * defaults["density"]
        if defaults["price_unit"] == "kg":
            matched["Ny kostnad [kr]"] = matched["Ny vekt [kg]"] * defaults["price"]
            matched["Ny CO2 [kgCO2e]"] = matched["Ny vekt [kg]"] * defaults["co2"]
        elif defaults["price_unit"] == "m3":
            matched["Ny kostnad [kr]"] = matched["Nytt volum [m3]"] * defaults["price"]
            matched["Ny CO2 [kgCO2e]"] = matched["Nytt volum [m3]"] * defaults["co2"]
        else:
            matched["Ny kostnad [kr]"] = matched["Areal [m2]"].fillna(0) * defaults["price"]
            matched["Ny CO2 [kgCO2e]"] = matched["Areal [m2]"].fillna(0) * defaults["co2"]
        matched["Byttemetode"] = matched.apply(
            lambda row: "Utledet lengde × nytt tverrsnitt" if pd.notna(row["Byttelengde [m]"]) and pd.notna(new_area_m2) and new_area_m2 > 0 else "Fallback til eksisterende volum",
            axis=1
        )
        matched["Nytt tverrsnittsareal [m2]"] = new_area_m2

    matched["Kostnadsendring [kr]"] = matched["Ny kostnad [kr]"] - matched["Gammel kostnad [kr]"]
    matched["Vektendring [kg]"] = matched["Ny vekt [kg]"] - matched["Gammel vekt [kg]"]
    matched["CO2-endring [kgCO2e]"] = matched["Ny CO2 [kgCO2e]"] - matched["Gammel CO2 [kgCO2e]"]
    matched["Prisgrunnlag"] = f"{defaults['label']} ({defaults['price_unit']})"
    matched["Tetthet brukt [kg/m3]"] = defaults["density"]
    matched["CO2-faktor brukt"] = defaults["co2"]
    matched["Norsk Prisbok-kode"] = defaults.get("npb_code", "")
    matched["ÅK/enh"] = defaults.get("ak", 0.0)
    return matched


def export_ifc_material_swap(ifc_bytes: bytes, source_df: pd.DataFrame, selected_type: str, from_material: str, target_key: str, new_profile_text: str = ""):
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert.")
    temp_in = None
    temp_out = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_in:
            tmp_in.write(ifc_bytes)
            temp_in = tmp_in.name
        model = ifcopenshell.open(temp_in)
        matched = source_df[(source_df["Type"] == selected_type) & (source_df["materiale"] == from_material)].copy()
        if matched.empty:
            return None, pd.DataFrame()
        defaults = get_swap_target_defaults(target_key)
        target_label = defaults.get("label", target_key)
        target_ids = set(matched["IFC GlobalId"].dropna().astype(str).tolist())
        changed_rows = []

        for type_name in SUPPORTED_IFC_TYPES:
            for el in model.by_type(type_name):
                gid = getattr(el, "GlobalId", "")
                if gid not in target_ids:
                    continue
                old_name = getattr(el, "Name", "") or ""
                old_object_type = getattr(el, "ObjectType", "") or ""
                old_material = get_ifc_material_name(el)
                assign_simple_material_to_element(model, el, target_label)
                try:
                    if new_profile_text:
                        el.ObjectType = new_profile_text
                    elif target_key in NORSK_PRISBOK_DATABASE:
                        el.ObjectType = target_label
                    else:
                        el.ObjectType = str(target_key)
                except Exception:
                    pass
                try:
                    desc_text = f"Materialbytte: {from_material} -> {target_label}"
                    if defaults.get("npb_code"):
                        desc_text += f" | Norsk Prisbok: {defaults['npb_code']}"
                    if new_profile_text:
                        desc_text += f" | Profil: {new_profile_text}"
                    el.Description = desc_text
                except Exception:
                    pass

                _set_or_create_pset_property(model, el, BYGGTOTAL_PSET_NAME, BYGGTOTAL_CHANGED_PROP, True, prop_type="bool")
                _set_or_create_pset_property(model, el, BYGGTOTAL_PSET_NAME, BYGGTOTAL_OLD_MATERIAL_PROP, str(old_material), prop_type="text")
                _set_or_create_pset_property(model, el, BYGGTOTAL_PSET_NAME, BYGGTOTAL_NEW_MATERIAL_PROP, str(target_label), prop_type="text")
                _set_or_create_pset_property(model, el, BYGGTOTAL_PSET_NAME, BYGGTOTAL_PROFILE_PROP, str(new_profile_text or ""), prop_type="text")

                changed_rows.append({
                    "IFC GlobalId": gid, "Navn": old_name, "Type": map_ifc_type(type_name),
                    "Gammelt materiale": old_material, "Nytt materiale": target_label,
                    "Norsk Prisbok-kode": defaults.get("npb_code", ""), "Nytt tverrsnitt": new_profile_text,
                    "Gammel ObjectType": old_object_type, "Ny ObjectType": getattr(el, "ObjectType", "") or "",
                })

        if not changed_rows:
            return None, pd.DataFrame()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_out:
            temp_out = tmp_out.name
        model.write(temp_out)
        with open(temp_out, "rb") as f:
            out_bytes = f.read()
        return out_bytes, pd.DataFrame(changed_rows)
    finally:
        for p in [temp_in, temp_out]:
            if p and os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass


# -------------------------
# PROSJEKTERINGSMODUL
# -------------------------

def safe_bool_ja_nei(value) -> str:
    text = str(value or "").strip().upper()
    return "JA" if text == "JA" else "NEI"


def generate_plan_geometry(params: dict) -> dict:
    fx = max(int(round(safe_num(params.get("fag_x_r1", 1)))), 1)
    fy = max(int(round(safe_num(params.get("fag_y_r1", 1)))), 1)
    dx = safe_num(params.get("faglengde_x_mm", 12000)) / 1000.0
    dy = safe_num(params.get("faglengde_y_mm", 16000)) / 1000.0

    main_width = fx * dx
    main_height = fy * dy
    rectangles = [{
        "name": "Rektangel 1",
        "x": 0.0, "y": 0.0, "width": main_width, "height": main_height,
    }]

    if safe_bool_ja_nei(params.get("rektangel2_aktiv", "NEI")) == "JA":
        r2_fx = max(int(round(safe_num(params.get("fag_x_r2", 0)))), 0)
        r2_fy = max(int(round(safe_num(params.get("fag_y_r2", 0)))), 0)
        if r2_fx > 0 and r2_fy > 0:
            rectangles.append({
                "name": "Rektangel 2",
                "x": safe_num(params.get("r2_offset_x_fag", 0)) * dx,
                "y": safe_num(params.get("r2_offset_y_fag", 0)) * dy,
                "width": r2_fx * dx,
                "height": r2_fy * dy,
            })

    opening = None
    ow = max(int(round(safe_num(params.get("opening_width_fag", 0)))), 0)
    oh = max(int(round(safe_num(params.get("opening_height_fag", 0)))), 0)
    if ow > 0 and oh > 0:
        opening = {
            "x": safe_num(params.get("opening_offset_x_fag", 0)) * dx,
            "y": safe_num(params.get("opening_offset_y_fag", 0)) * dy,
            "width": ow * dx,
            "height": oh * dy,
        }

    union_area = sum(r["width"] * r["height"] for r in rectangles)
    if len(rectangles) == 2:
        r1, r2 = rectangles[0], rectangles[1]
        overlap_w = max(0.0, min(r1["x"] + r1["width"], r2["x"] + r2["width"]) - max(r1["x"], r2["x"]))
        overlap_h = max(0.0, min(r1["y"] + r1["height"], r2["y"] + r2["height"]) - max(r1["y"], r2["y"]))
        union_area -= overlap_w * overlap_h

    opening_area = 0.0
    if opening:
        opening_area = opening["width"] * opening["height"]

    active_area = max(union_area - opening_area, 0.0)

    if len(rectangles) == 1 and not opening:
        planformkode = "R1"
    elif len(rectangles) == 1 and opening:
        planformkode = "R1_MED_APNING"
    elif len(rectangles) == 2 and not opening:
        planformkode = "L_FORM"
    else:
        planformkode = "L_FORM_MED_APNING"

    return {
        "rectangles": rectangles,
        "opening": opening,
        "active_area_m2": active_area,
        "gross_area_m2": union_area,
        "opening_area_m2": opening_area,
        "planformkode": planformkode,
        "dx": dx,
        "dy": dy,
        "dekkenivaer_m": [safe_num(params.get("dekke_tykkelse_mm", 0)) * 0 + (i + 1) * 4 for i in range(max(int(round(safe_num(params.get("dekker_i_modell", 1)))), 1))],
    }


def rectangle_inside(rect, outer):
    return (
        rect["x"] >= outer["x"]
        and rect["y"] >= outer["y"]
        and rect["x"] + rect["width"] <= outer["x"] + outer["width"]
        and rect["y"] + rect["height"] <= outer["y"] + outer["height"]
    )


def format_point(x, y):
    return f"{int(round(x * 1000))}, {int(round(y * 1000))}"


def generate_slab_export(params: dict) -> pd.DataFrame:
    geom = generate_plan_geometry(params)
    n_decks = max(int(round(safe_num(params.get("dekker_i_modell", 1)))), 1)
    rows = []
    rectangles = geom["rectangles"]
    opening = geom["opening"]

    etasjeh_mm = safe_num(params.get("etasjehoyde_mm", 3000))
    if etasjeh_mm <= 0:
        etasjeh_mm = 3000

    for i in range(n_decks):
        z_mm = int(round((i + 1) * etasjeh_mm))
        if len(rectangles) == 1 and not opening:
            r = rectangles[0]
            pts = [(r["x"], r["y"]), (r["x"] + r["width"], r["y"]), (r["x"] + r["width"], r["y"] + r["height"]), (r["x"], r["y"] + r["height"])]
        else:
            xmin = min(r["x"] for r in rectangles)
            ymin = min(r["y"] for r in rectangles)
            xmax = max(r["x"] + r["width"] for r in rectangles)
            ymax = max(r["y"] + r["height"] for r in rectangles)
            pts = [(xmin, ymin), (xmax, ymin), (xmax, ymax), (xmin, ymax)]

        row = {
            "DeckID": f"D{i+1}",
            "Nivå": i + 1,
            "Aktiv": 1 if safe_bool_ja_nei(params.get("dekker_aktiv", "JA")) == "JA" else 0,
            "Z [mm]": z_mm,
            "Skalltype": params.get("skalltype", "Platt skall"),
            "Materiale": params.get("dekke_materiale", "B35, Betong"),
            "Tykkelse [mm]": safe_num(params.get("dekke_tykkelse_mm", 300)),
            "Areal [m²]": geom["active_area_m2"],
        }
        for idx in range(8):
            row[f"P{idx+1} (X,Y)"] = format_point(*pts[idx]) if idx < len(pts) else ""
        rows.append(row)

    return pd.DataFrame(rows)


def run_project_qa(params: dict, frame_df: pd.DataFrame, slab_df: pd.DataFrame) -> pd.DataFrame:
    geom = generate_plan_geometry(params)
    rectangles = geom["rectangles"]
    opening = geom["opening"]

    checks = []
    checks.append({
        "Kontroll": "Grunnparametre",
        "Status": "OK" if safe_num(params.get("fag_x_r1")) >= 1 and safe_num(params.get("fag_y_r1")) >= 1 and safe_num(params.get("antall_etasjer")) >= 1 else "FEIL",
        "Melding": "Fag og etasjer er satt opp." if safe_num(params.get("fag_x_r1")) >= 1 and safe_num(params.get("fag_y_r1")) >= 1 and safe_num(params.get("antall_etasjer")) >= 1 else "Mangler eller ugyldige inputverdier.",
        "Anbefaling": "Sett minimum 1 fag i hver retning og minst 1 etasje.",
    })
    checks.append({
        "Kontroll": "Dekker aktivert",
        "Status": "OK" if safe_bool_ja_nei(params.get("dekker_aktiv", "JA")) == "JA" else "INFO",
        "Melding": "Dekker er aktive." if safe_bool_ja_nei(params.get("dekker_aktiv", "JA")) == "JA" else "Dekker er deaktivert.",
        "Anbefaling": "Aktiver dekker hvis du ønsker skalleksport.",
    })

    opening_status = "Ingen åpning"
    opening_msg = "Ingen åpning definert."
    if opening:
        inside_main = rectangle_inside(opening, rectangles[0])
        opening_status = "OK" if inside_main else "SJEKK"
        opening_msg = "Åpningen ligger innenfor hovedrektangelet." if inside_main else "Åpningen ligger utenfor hovedrektangelet."
    checks.append({"Kontroll": "Åpning innenfor hovedrektangel", "Status": opening_status, "Melding": opening_msg, "Anbefaling": "Juster åpningens plassering eller størrelse ved behov."})

    thickness_ok = safe_num(params.get("dekke_tykkelse_mm", 0)) > 0
    checks.append({"Kontroll": "Tykkelse > 0", "Status": "OK" if thickness_ok else "SJEKK", "Melding": f"Tykkelse: {safe_num(params.get('dekke_tykkelse_mm', 0)):.0f} mm", "Anbefaling": "Bruk positiv tykkelse."})

    active_area_ok = geom["active_area_m2"] > 0
    checks.append({"Kontroll": "Aktivt dekkeareal", "Status": "OK" if active_area_ok else "SJEKK", "Melding": f"Aktivt areal: {geom['active_area_m2']:.1f} m²", "Anbefaling": "Kontroller geometri og åpninger."})

    export_ok = active_area_ok and thickness_ok and len(frame_df) > 0 and len(slab_df) > 0
    checks.append({"Kontroll": "Eksportstatus", "Status": "KLAR" if export_ok else "MÅ SJEKKES", "Melding": "Eksportgrunnlaget er klart." if export_ok else "En eller flere kontroller feilet.", "Anbefaling": "Løs eventuelle røde/sjekk-punkter før eksport."})

    checks.append({"Kontroll": "Planformkode", "Status": "OK", "Melding": geom["planformkode"], "Anbefaling": "Brukes som intern kode for planformen."})
    checks.append({"Kontroll": "Antall eksportlinjer ramme", "Status": "OK", "Melding": str(len(frame_df)), "Anbefaling": "Kontroller at antallet virker realistisk."})
    checks.append({"Kontroll": "Antall eksportlinjer dekker", "Status": "OK", "Melding": str(len(slab_df)), "Anbefaling": "Kontroller at antallet dekker samsvarer med modellen."})

    return pd.DataFrame(checks)


def plot_plan_geometry(geom: dict):
    fig, ax = plt.subplots(figsize=(8, 6))
    for rect in geom["rectangles"]:
        patch = plt.Rectangle((rect["x"], rect["y"]), rect["width"], rect["height"], fill=False, linewidth=2)
        ax.add_patch(patch)
        ax.text(rect["x"] + rect["width"] / 2, rect["y"] + rect["height"] / 2, rect["name"], ha="center", va="center")

    if geom["opening"]:
        o = geom["opening"]
        patch = plt.Rectangle((o["x"], o["y"]), o["width"], o["height"], fill=False, linestyle="--", linewidth=2)
        ax.add_patch(patch)
        ax.text(o["x"] + o["width"] / 2, o["y"] + o["height"] / 2, "Åpning", ha="center", va="center")

    xmax = max(r["x"] + r["width"] for r in geom["rectangles"])
    ymax = max(r["y"] + r["height"] for r in geom["rectangles"])
    ax.set_xlim(-1, xmax + 1)
    ax.set_ylim(-1, ymax + 1)
    ax.set_aspect("equal")
    ax.set_xlabel("X [m]")
    ax.set_ylabel("Y [m]")
    ax.set_title("2D planform")
    ax.grid(True, alpha=0.3)
    return fig


from reports import build_docx_report, build_pdf_report, make_report_summary_dict




# -------------------------
# BYGGGENERATOR – PARAMETRISK MODELL FRA EXCEL-LOGIKK
# -------------------------

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


def quality_options(material: str):
    return list(QUALITY_LIBRARY.get(material, {"Standard": {}}).keys())


def make_material_profile_label(material: str, quality: str, profile: str) -> str:
    if material == "Betong":
        return f"{quality}, Betong / {profile}"
    if material == "Stål":
        return f"{quality}, Stål / {profile}"
    if material == "Limtre":
        return f"{quality}, Limtre / {profile}"
    if material == "Massivtre":
        return f"{quality}, Massivtre / {profile}"
    return f"{quality}, {material} / {profile}"


def material_quality_values(material: str, quality: str) -> dict:
    return QUALITY_LIBRARY.get(material, {}).get(quality, {"density": 1000.0, "co2": 200.0, "price": 1000.0})


def generate_frame_export_parametric(params: dict) -> pd.DataFrame:
    """Genererer søyle-/bjelkesystem fra valg i appen. Støtter rektangel 1 og valgfritt rektangel 2."""
    geom = generate_plan_geometry(params)
    dx = geom["dx"]
    dy = geom["dy"]
    etasjeh = safe_num(params.get("etasjehoyde_mm", 3000)) / 1000.0
    if etasjeh <= 0:
        etasjeh = 3.0
    n_levels = max(int(round(safe_num(params.get("antall_etasjer", 1)))), 1)
    rows = []
    col_id = 1
    beam_id = 1
    seen_columns = set()
    seen_beams = set()

    for level in range(1, n_levels + 1):
        z0 = (level - 1) * etasjeh
        z1 = level * etasjeh
        for rect in geom["rectangles"]:
            fx = max(int(round(rect["width"] / dx)), 1) if dx > 0 else 1
            fy = max(int(round(rect["height"] / dy)), 1) if dy > 0 else 1
            x0 = rect["x"]
            y0 = rect["y"]
            for ix in range(fx + 1):
                for iy in range(fy + 1):
                    x = round(x0 + ix * dx, 6)
                    y = round(y0 + iy * dy, 6)
                    ckey = (level, x, y, round(z0, 6), round(z1, 6))
                    if ckey not in seen_columns:
                        rows.append({"Type": "Søyle", "ID": f"C{col_id}", "Nivå": level, "X1 [m]": x, "Y1 [m]": y, "Z1 [m]": z0, "X2 [m]": x, "Y2 [m]": y, "Z2 [m]": z1})
                        seen_columns.add(ckey)
                        col_id += 1
            for iy in range(fy + 1):
                y = round(y0 + iy * dy, 6)
                for ix in range(fx):
                    x1 = round(x0 + ix * dx, 6)
                    x2 = round(x0 + (ix + 1) * dx, 6)
                    bkey = (level, min(x1, x2), y, z1, max(x1, x2), y, z1)
                    if bkey not in seen_beams:
                        rows.append({"Type": "Bjelke", "ID": f"B{beam_id}", "Nivå": level, "X1 [m]": x1, "Y1 [m]": y, "Z1 [m]": z1, "X2 [m]": x2, "Y2 [m]": y, "Z2 [m]": z1})
                        seen_beams.add(bkey)
                        beam_id += 1
            for ix in range(fx + 1):
                x = round(x0 + ix * dx, 6)
                for iy in range(fy):
                    y1 = round(y0 + iy * dy, 6)
                    y2 = round(y0 + (iy + 1) * dy, 6)
                    bkey = (level, x, min(y1, y2), z1, x, max(y1, y2), z1)
                    if bkey not in seen_beams:
                        rows.append({"Type": "Bjelke", "ID": f"B{beam_id}", "Nivå": level, "X1 [m]": x, "Y1 [m]": y1, "Z1 [m]": z1, "X2 [m]": x, "Y2 [m]": y2, "Z2 [m]": z1})
                        seen_beams.add(bkey)
                        beam_id += 1
    return pd.DataFrame(rows)


def frame_to_quantity_dataset(frame_df: pd.DataFrame, slab_df: pd.DataFrame, params: dict) -> pd.DataFrame:
    rows = []
    beam_mat = params.get("bjelkemateriale", "Stål")
    beam_quality = params.get("bjelkekvalitet", "S355")
    beam_profile = params.get("bjelkeprofil", "KFHUP 200x200x12.5")
    col_mat = params.get("søylemateriale", "Stål")
    col_quality = params.get("søylekvalitet", "S355")
    col_profile = params.get("søyleprofil", "KFHUP 200x200x12.5")

    for _, r in frame_df.iterrows():
        typ = str(r["Type"])
        if typ == "Bjelke":
            mat, qual, profile = beam_mat, beam_quality, beam_profile
        else:
            mat, qual, profile = col_mat, col_quality, col_profile
        length = math.sqrt((safe_num(r["X2 [m]"]) - safe_num(r["X1 [m]"])) ** 2 + (safe_num(r["Y2 [m]"]) - safe_num(r["Y1 [m]"])) ** 2 + (safe_num(r["Z2 [m]"]) - safe_num(r["Z1 [m]"])) ** 2)
        area = parse_profile_area_from_text(profile, mat)
        if pd.isna(area) or area <= 0:
            area = 0.04 if mat == "Stål" else 0.09
        volume = length * area
        qv = material_quality_values(mat, qual)
        weight = volume * qv["density"]
        if mat == "Stål":
            cost = weight * qv["price"]
            co2 = weight * qv["co2"]
        else:
            cost = volume * qv["price"]
            co2 = volume * qv["co2"]
        rows.append({
            "Segment": r["ID"], "Type": typ, "Nivå": r["Nivå"],
            "Knutepunkter": "", "Material / Tverrsnitt": make_material_profile_label(mat, qual, profile),
            "Lengde [m]": length, "Areal [m2]": area, "Volum [m3]": volume, "Vekt [kg]": weight,
            "materiale": mat, "Materialkvalitet": qual, "Mengdegrunnlag": "Bygggenerator", "Endret IFC": False,
            "Kostnad [kr]": cost, "CO2 [kgCO2e]": co2,
        })

    deck_mat = params.get("dekke_materialtype", "Betong")
    deck_quality = params.get("dekke_kvalitet", "B35")
    deck_profile = f"t={safe_num(params.get('dekke_tykkelse_mm', 300)):.0f} mm"
    qv = material_quality_values(deck_mat, deck_quality)
    for _, r in slab_df.iterrows():
        area = safe_num(r.get("Areal [m²]", r.get("Areal [m2]", 0)))
        thickness = safe_num(r.get("Tykkelse [mm]", params.get("dekke_tykkelse_mm", 300))) / 1000.0
        volume = area * thickness
        weight = volume * qv["density"]
        cost = volume * qv["price"]
        co2 = volume * qv["co2"]
        rows.append({
            "Segment": r.get("DeckID", "D"), "Type": "Dekke", "Nivå": r.get("Nivå", 1),
            "Knutepunkter": "", "Material / Tverrsnitt": make_material_profile_label(deck_mat, deck_quality, deck_profile),
            "Lengde [m]": math.nan, "Areal [m2]": area, "Volum [m3]": volume, "Vekt [kg]": weight,
            "materiale": deck_mat, "Materialkvalitet": deck_quality, "Mengdegrunnlag": "Bygggenerator", "Endret IFC": False,
            "Kostnad [kr]": cost, "CO2 [kgCO2e]": co2,
        })
    return pd.DataFrame(rows)


def plot_frame_3d(frame_df: pd.DataFrame, slab_df: pd.DataFrame | None = None):
    fig = go.Figure()
    if frame_df is not None and not frame_df.empty:
        for typ, group in frame_df.groupby("Type"):
            xs, ys, zs = [], [], []
            for _, r in group.iterrows():
                xs += [r["X1 [m]"], r["X2 [m]"], None]
                ys += [r["Y1 [m]"], r["Y2 [m]"], None]
                zs += [r["Z1 [m]"], r["Z2 [m]"], None]
            fig.add_trace(go.Scatter3d(x=xs, y=ys, z=zs, mode="lines", name=typ, line=dict(width=6 if typ == "Søyle" else 4)))
    if slab_df is not None and not slab_df.empty:
        for _, r in slab_df.iterrows():
            pts = []
            for i in range(1, 9):
                val = str(r.get(f"P{i} (X,Y)", "") or "")
                nums = [float(x.strip())/1000.0 for x in val.split(",") if x.strip().replace("-", "").isdigit()]
                if len(nums) == 2:
                    pts.append(nums)
            if len(pts) >= 3:
                z = safe_num(r.get("Z [mm]", 0)) / 1000.0
                fig.add_trace(go.Mesh3d(x=[p[0] for p in pts], y=[p[1] for p in pts], z=[z]*len(pts), opacity=0.18, name=f"Dekke {r.get('Nivå','')}", color="#808080"))
    fig.update_layout(height=650, margin=dict(l=0, r=0, t=20, b=0), scene=dict(xaxis_title="X [m]", yaxis_title="Y [m]", zaxis_title="Z [m]", aspectmode="data"))
    return fig


# -------------------------
# IFC-EKSPORT FOR BYGGGENERATOR
# -------------------------

def _ifc_guid():
    return ifcopenshell.guid.new()


def _ifc_point(model, x=0.0, y=0.0, z=0.0):
    return model.create_entity("IfcCartesianPoint", Coordinates=(float(x), float(y), float(z)))


def _ifc_dir(model, x=0.0, y=0.0, z=1.0):
    return model.create_entity("IfcDirection", DirectionRatios=(float(x), float(y), float(z)))


def _ifc_axis3d(model, location=None, axis=None, ref_direction=None):
    return model.create_entity(
        "IfcAxis2Placement3D",
        Location=location or _ifc_point(model, 0, 0, 0),
        Axis=axis or _ifc_dir(model, 0, 0, 1),
        RefDirection=ref_direction or _ifc_dir(model, 1, 0, 0),
    )


def _ifc_local_placement(model, relative_to=None, x=0.0, y=0.0, z=0.0, axis=None, ref_direction=None):
    return model.create_entity(
        "IfcLocalPlacement",
        PlacementRelTo=relative_to,
        RelativePlacement=_ifc_axis3d(model, _ifc_point(model, x, y, z), axis, ref_direction),
    )


def _ifc_product_shape(model, context, width: float, height: float, depth: float):
    """Lager enkel rektangulær SweptSolid-geometri. Lokalt ekstruderes profilen i Z-retningen."""
    profile = model.create_entity(
        "IfcRectangleProfileDef",
        ProfileType="AREA",
        ProfileName="byggTotal rektangelprofil",
        Position=model.create_entity("IfcAxis2Placement2D", Location=model.create_entity("IfcCartesianPoint", Coordinates=(0.0, 0.0))),
        XDim=max(float(width), 0.001),
        YDim=max(float(height), 0.001),
    )
    solid = model.create_entity(
        "IfcExtrudedAreaSolid",
        SweptArea=profile,
        Position=_ifc_axis3d(model),
        ExtrudedDirection=_ifc_dir(model, 0, 0, 1),
        Depth=max(float(depth), 0.001),
    )
    body = model.create_entity(
        "IfcShapeRepresentation",
        ContextOfItems=context,
        RepresentationIdentifier="Body",
        RepresentationType="SweptSolid",
        Items=[solid],
    )
    return model.create_entity("IfcProductDefinitionShape", Representations=[body])


def _get_profile_dims_m(profile_text: str, material_hint: str, fallback=(0.2, 0.2)):
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+[\.,]?\d*", str(profile_text or ""))]
    if len(nums) >= 2:
        # For stål HUP/KFHUP med tre tall brukes bredde/høyde som de to første av de tre siste tallene.
        if classify_material(material_hint if material_hint else profile_text) == "Stål" and len(nums) >= 3:
            return max(nums[-3] / 1000.0, 0.001), max(nums[-2] / 1000.0, 0.001)
        return max(nums[-2] / 1000.0, 0.001), max(nums[-1] / 1000.0, 0.001)
    return fallback


def _beam_orientation(model, x1, y1, z1, x2, y2, z2):
    dx, dy, dz = float(x2 - x1), float(y2 - y1), float(z2 - z1)
    length = math.sqrt(dx * dx + dy * dy + dz * dz)
    if length <= 0:
        return _ifc_dir(model, 0, 0, 1), _ifc_dir(model, 1, 0, 0), 0.001
    axis = _ifc_dir(model, dx / length, dy / length, dz / length)
    # RefDirection må ikke være parallell med Axis. Global Z fungerer for horisontale bjelker.
    if abs(dz / length) > 0.95:
        ref = _ifc_dir(model, 1, 0, 0)
    else:
        ref = _ifc_dir(model, 0, 0, 1)
    return axis, ref, length


def generate_building_ifc_bytes(frame_df: pd.DataFrame, slab_df: pd.DataFrame, params: dict, project_name: str = "byggTotal generert bygg") -> bytes:
    """Genererer en enkel IFC4-fil fra Bygggeneratoren.

    IFC-en inneholder etasjer, søyler, bjelker og dekker med enkel rektangulær geometri,
    materialnavn og spatial containment. Den er laget for tidligfasevisning i Solibri/BIM-viewer.
    """
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert. Legg til ifcopenshell i requirements.txt for IFC-eksport.")

    model = ifcopenshell.file(schema="IFC4")

    origin = _ifc_axis3d(model)
    context = model.create_entity(
        "IfcGeometricRepresentationContext",
        ContextIdentifier="Model",
        ContextType="Model",
        CoordinateSpaceDimension=3,
        Precision=1e-5,
        WorldCoordinateSystem=origin,
    )
    units = model.create_entity("IfcUnitAssignment", Units=[
        model.create_entity("IfcSIUnit", UnitType="LENGTHUNIT", Name="METRE"),
        model.create_entity("IfcSIUnit", UnitType="AREAUNIT", Name="SQUARE_METRE"),
        model.create_entity("IfcSIUnit", UnitType="VOLUMEUNIT", Name="CUBIC_METRE"),
    ])

    project = model.create_entity("IfcProject", GlobalId=_ifc_guid(), Name=project_name, RepresentationContexts=[context], UnitsInContext=units)
    site_placement = _ifc_local_placement(model)
    site = model.create_entity("IfcSite", GlobalId=_ifc_guid(), Name="Tomt", ObjectPlacement=site_placement)
    building_placement = _ifc_local_placement(model, site_placement)
    building = model.create_entity("IfcBuilding", GlobalId=_ifc_guid(), Name="Generert råbygg", ObjectPlacement=building_placement)
    model.create_entity("IfcRelAggregates", GlobalId=_ifc_guid(), RelatingObject=project, RelatedObjects=[site])
    model.create_entity("IfcRelAggregates", GlobalId=_ifc_guid(), RelatingObject=site, RelatedObjects=[building])

    n_levels = max(int(round(safe_num(params.get("antall_etasjer", 1)))), 1)
    etasjeh = safe_num(params.get("etasjehoyde_mm", 3000)) / 1000.0
    if etasjeh <= 0:
        etasjeh = 3.0

    storeys = {}
    storey_children = {i: [] for i in range(1, n_levels + 1)}
    storey_list = []
    for level in range(1, n_levels + 1):
        z = (level - 1) * etasjeh
        sp = _ifc_local_placement(model, building_placement, 0, 0, z)
        storey = model.create_entity("IfcBuildingStorey", GlobalId=_ifc_guid(), Name=f"Etasje {level}", ObjectPlacement=sp, Elevation=z)
        storeys[level] = storey
        storey_list.append(storey)
    model.create_entity("IfcRelAggregates", GlobalId=_ifc_guid(), RelatingObject=building, RelatedObjects=storey_list)

    material_entities = {}
    def get_material(name):
        name = str(name or "Ukjent")
        if name not in material_entities:
            material_entities[name] = model.create_entity("IfcMaterial", Name=name)
        return material_entities[name]

    def assign_material(product, mat_name):
        model.create_entity("IfcRelAssociatesMaterial", GlobalId=_ifc_guid(), RelatedObjects=[product], RelatingMaterial=get_material(mat_name))

    beam_mat = params.get("bjelkemateriale", "Stål")
    beam_quality = params.get("bjelkekvalitet", "S355")
    beam_profile = params.get("bjelkeprofil", "KFHUP 200x200x12.5")
    col_mat = params.get("søylemateriale", "Stål")
    col_quality = params.get("søylekvalitet", "S355")
    col_profile = params.get("søyleprofil", "KFHUP 200x200x12.5")
    deck_mat = params.get("dekke_materialtype", "Betong")
    deck_quality = params.get("dekke_kvalitet", "B35")
    deck_thk = safe_num(params.get("dekke_tykkelse_mm", 300)) / 1000.0
    if deck_thk <= 0:
        deck_thk = 0.3

    if frame_df is not None and not frame_df.empty:
        for _, r in frame_df.iterrows():
            typ = str(r.get("Type", ""))
            level = max(int(round(safe_num(r.get("Nivå", 1)))), 1)
            x1, y1, z1 = safe_num(r.get("X1 [m]")), safe_num(r.get("Y1 [m]")), safe_num(r.get("Z1 [m]"))
            x2, y2, z2 = safe_num(r.get("X2 [m]")), safe_num(r.get("Y2 [m]")), safe_num(r.get("Z2 [m]"))
            axis, ref, length = _beam_orientation(model, x1, y1, z1, x2, y2, z2)
            if typ == "Søyle":
                w, h = _get_profile_dims_m(col_profile, col_mat, fallback=(0.3, 0.3))
                mat_label = f"{col_mat} {col_quality}"
                shape = _ifc_product_shape(model, context, w, h, length)
                placement = _ifc_local_placement(model, None, x1, y1, z1, axis, ref)
                product = model.create_entity("IfcColumn", GlobalId=_ifc_guid(), Name=str(r.get("ID", "Søyle")), ObjectPlacement=placement, Representation=shape)
            else:
                w, h = _get_profile_dims_m(beam_profile, beam_mat, fallback=(0.2, 0.3))
                mat_label = f"{beam_mat} {beam_quality}"
                shape = _ifc_product_shape(model, context, w, h, length)
                placement = _ifc_local_placement(model, None, x1, y1, z1, axis, ref)
                product = model.create_entity("IfcBeam", GlobalId=_ifc_guid(), Name=str(r.get("ID", "Bjelke")), ObjectPlacement=placement, Representation=shape)
            assign_material(product, mat_label)
            if level in storey_children:
                storey_children[level].append(product)

    if slab_df is not None and not slab_df.empty:
        for _, r in slab_df.iterrows():
            level = max(int(round(safe_num(r.get("Nivå", 1)))), 1)
            pts = []
            for i in range(1, 9):
                raw = str(r.get(f"P{i} (X,Y)", "") or "")
                nums = [float(x.strip()) / 1000.0 for x in raw.split(",") if x.strip().replace("-", "").replace(".", "").isdigit()]
                if len(nums) == 2:
                    pts.append(tuple(nums))
            if len(pts) >= 3:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                xmin, xmax = min(xs), max(xs)
                ymin, ymax = min(ys), max(ys)
                width = max(xmax - xmin, 0.001)
                depth = max(ymax - ymin, 0.001)
                z_top = safe_num(r.get("Z [mm]", level * etasjeh * 1000)) / 1000.0
                shape = _ifc_product_shape(model, context, width, depth, deck_thk)
                placement = _ifc_local_placement(model, None, xmin + width / 2.0, ymin + depth / 2.0, z_top - deck_thk, _ifc_dir(model, 0, 0, 1), _ifc_dir(model, 1, 0, 0))
                slab = model.create_entity("IfcSlab", GlobalId=_ifc_guid(), Name=str(r.get("DeckID", f"Dekke {level}")), ObjectPlacement=placement, Representation=shape, PredefinedType="FLOOR")
                assign_material(slab, f"{deck_mat} {deck_quality}")
                if level in storey_children:
                    storey_children[level].append(slab)

    for level, children in storey_children.items():
        if children:
            model.create_entity("IfcRelContainedInSpatialStructure", GlobalId=_ifc_guid(), RelatedElements=children, RelatingStructure=storeys[level])

    with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp:
        temp_path = tmp.name
    try:
        model.write(temp_path)
        with open(temp_path, "rb") as f:
            return f.read()
    finally:
        try:
            os.remove(temp_path)
        except Exception:
            pass


def generate_complete_ifc_bytes(
    frame_df: pd.DataFrame,
    slab_df: pd.DataFrame,
    params: dict,
    ground_layers: list | None = None,
    gwl_depth_m: float | None = None,
    foundation_key: str | None = None,
    foundation_area_m2: float = 0.0,
    pile_lm: float = 0.0,
    n_piles: int = 0,
    pile_length_m: float = 0.0,
    site_area_m2: float = 0.0,
    project_name: str = "byggTotal – komplett modell",
    building_offset_x: float = 0.0,
    building_offset_y: float = 0.0,
    building_rotation_deg: float = 0.0,
) -> bytes:
    """Genererer en komplett IFC4-fil med konstruksjon OG grunnforhold.

    Grunnforhold modelleres som:
    - IfcGeographicElement (terreng/tomt)
    - IfcFooting (fundamentplate eller stripe)
    - IfcPile (peler)
    - IfcBuildingElementProxy (geotekniske lag)
    Konstruksjon er identisk med generate_building_ifc_bytes.
    """
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert.")

    from ground_module import SOIL_DATABASE, FOUNDATION_DATABASE

    model = ifcopenshell.file(schema="IFC4")
    origin = _ifc_axis3d(model)
    context = model.create_entity(
        "IfcGeometricRepresentationContext",
        ContextIdentifier="Model", ContextType="Model",
        CoordinateSpaceDimension=3, Precision=1e-5,
        WorldCoordinateSystem=origin,
    )
    units = model.create_entity("IfcUnitAssignment", Units=[
        model.create_entity("IfcSIUnit", UnitType="LENGTHUNIT", Name="METRE"),
        model.create_entity("IfcSIUnit", UnitType="AREAUNIT", Name="SQUARE_METRE"),
        model.create_entity("IfcSIUnit", UnitType="VOLUMEUNIT", Name="CUBIC_METRE"),
    ])
    project = model.create_entity("IfcProject", GlobalId=_ifc_guid(), Name=project_name,
                                   RepresentationContexts=[context], UnitsInContext=units)
    import math as _ifc_math

    site_placement = _ifc_local_placement(model)
    site = model.create_entity("IfcSite", GlobalId=_ifc_guid(), Name="Tomt",
                                ObjectPlacement=site_placement)

    # Bygg plasseres med offset og rotasjon relativt til tomten
    _rot_rad = _ifc_math.radians(building_rotation_deg)
    _cos_r   = _ifc_math.cos(_rot_rad)
    _sin_r   = _ifc_math.sin(_rot_rad)
    # IFC-rotasjon: RefDirection angir X-aksen, Axis angir Z-aksen
    _rot_axis      = _ifc_dir(model, 0.0, 0.0, 1.0)
    _rot_ref_dir   = _ifc_dir(model, _cos_r, _sin_r, 0.0)
    building_placement = _ifc_local_placement(
        model, site_placement,
        building_offset_x, building_offset_y, 0.0,
        _rot_axis, _rot_ref_dir,
    )
    building = model.create_entity("IfcBuilding", GlobalId=_ifc_guid(), Name="Generert råbygg",
                                    ObjectPlacement=building_placement)
    model.create_entity("IfcRelAggregates", GlobalId=_ifc_guid(),
                         RelatingObject=project, RelatedObjects=[site])
    model.create_entity("IfcRelAggregates", GlobalId=_ifc_guid(),
                         RelatingObject=site, RelatedObjects=[building])

    n_levels = max(int(round(safe_num(params.get("antall_etasjer", 1)))), 1)
    etasjeh = safe_num(params.get("etasjehoyde_mm", 3000)) / 1000.0
    if etasjeh <= 0:
        etasjeh = 3.0

    # Beregn byggets faktiske fotavtrykk fra frame_df-koordinater
    # Dette sikrer at fundament og konstruksjon alltid bruker samme dimensjoner
    if frame_df is not None and not frame_df.empty:
        _all_x = []
        _all_y = []
        for _c in ["X1 [m]", "X2 [m]"]:
            if _c in frame_df.columns:
                _all_x += frame_df[_c].tolist()
        for _c in ["Y1 [m]", "Y2 [m]"]:
            if _c in frame_df.columns:
                _all_y += frame_df[_c].tolist()
        _bx = max(_all_x) - min(_all_x) if _all_x else 10.0
        _by = max(_all_y) - min(_all_y) if _all_y else 10.0
        _x0 = min(_all_x) if _all_x else 0.0
        _y0 = min(_all_y) if _all_y else 0.0
    else:
        _fag_x = safe_num(params.get("fag_x_r1", 4))
        _fag_y = safe_num(params.get("fag_y_r1", 2))
        _dx    = safe_num(params.get("faglengde_x_mm", 8000)) / 1000.0
        _dy    = safe_num(params.get("faglengde_y_mm", 12000)) / 1000.0
        _bx = _fag_x * _dx if _fag_x > 0 and _dx > 0 else max(foundation_area_m2 ** 0.5, 1.0)
        _by = _fag_y * _dy if _fag_y > 0 and _dy > 0 else max(foundation_area_m2 ** 0.5, 1.0)
        _x0, _y0 = 0.0, 0.0

    # Beregn total grunnlagsdybde så bygget kan løftes opp
    total_ground_depth = 0.0
    if ground_layers:
        total_ground_depth = sum(float(l["thickness_m"]) for l in ground_layers)
    fd_thk_lift = 0.3 if (foundation_key and foundation_area_m2 > 0) else 0.0
    building_z_offset = fd_thk_lift  # bygget starter på toppen av fundamentplaten

    storeys = {}
    storey_children = {i: [] for i in range(1, n_levels + 1)}
    storey_list = []
    for level in range(1, n_levels + 1):
        z = building_z_offset + (level - 1) * etasjeh
        sp = _ifc_local_placement(model, building_placement, 0, 0, z)
        storey = model.create_entity("IfcBuildingStorey", GlobalId=_ifc_guid(),
                                      Name=f"Etasje {level}", ObjectPlacement=sp, Elevation=z)
        storeys[level] = storey
        storey_list.append(storey)
    model.create_entity("IfcRelAggregates", GlobalId=_ifc_guid(),
                         RelatingObject=building, RelatedObjects=storey_list)

    material_entities = {}
    def get_material(name):
        name = str(name or "Ukjent")
        if name not in material_entities:
            material_entities[name] = model.create_entity("IfcMaterial", Name=name)
        return material_entities[name]
    def assign_material(product, mat_name):
        model.create_entity("IfcRelAssociatesMaterial", GlobalId=_ifc_guid(),
                             RelatedObjects=[product], RelatingMaterial=get_material(mat_name))

    beam_mat = params.get("bjelkemateriale", "Stål")
    beam_quality = params.get("bjelkekvalitet", "S355")
    beam_profile = params.get("bjelkeprofil", "KFHUP 200x200x12.5")
    col_mat = params.get("søylemateriale", "Stål")
    col_quality = params.get("søylekvalitet", "S355")
    col_profile = params.get("søyleprofil", "KFHUP 200x200x12.5")
    deck_mat = params.get("dekke_materialtype", "Betong")
    deck_quality = params.get("dekke_kvalitet", "B35")
    deck_thk = safe_num(params.get("dekke_tykkelse_mm", 300)) / 1000.0
    if deck_thk <= 0:
        deck_thk = 0.3

    def _world_xy(lx, ly):
        """Transformer lokal (lx,ly) i byggets koordinatsystem til verdenskoordinater."""
        return (
            building_offset_x + lx * _cos_r - ly * _sin_r,
            building_offset_y + lx * _sin_r + ly * _cos_r,
        )

    if frame_df is not None and not frame_df.empty:
        for _, r in frame_df.iterrows():
            typ = str(r.get("Type", ""))
            level = max(int(round(safe_num(r.get("Nivå", 1)))), 1)
            # Lokal posisjon i bygget
            lx1, ly1 = safe_num(r.get("X1 [m]")), safe_num(r.get("Y1 [m]"))
            lx2, ly2 = safe_num(r.get("X2 [m]")), safe_num(r.get("Y2 [m]"))
            z1 = safe_num(r.get("Z1 [m]")) + building_z_offset
            z2 = safe_num(r.get("Z2 [m]")) + building_z_offset
            # Verdenskoordinater
            wx1, wy1 = _world_xy(lx1, ly1)
            wx2, wy2 = _world_xy(lx2, ly2)
            axis, ref, length = _beam_orientation(model, wx1, wy1, z1, wx2, wy2, z2)
            if typ == "Søyle":
                w, h = _get_profile_dims_m(col_profile, col_mat, fallback=(0.3, 0.3))
                shape = _ifc_product_shape(model, context, w, h, length)
                placement = _ifc_local_placement(model, site_placement, wx1, wy1, z1, axis, ref)
                product = model.create_entity("IfcColumn", GlobalId=_ifc_guid(),
                                               Name=str(r.get("ID", "Søyle")),
                                               ObjectPlacement=placement, Representation=shape)
                assign_material(product, f"{col_mat} {col_quality}")
            else:
                w, h = _get_profile_dims_m(beam_profile, beam_mat, fallback=(0.2, 0.3))
                shape = _ifc_product_shape(model, context, w, h, length)
                placement = _ifc_local_placement(model, site_placement, wx1, wy1, z1, axis, ref)
                product = model.create_entity("IfcBeam", GlobalId=_ifc_guid(),
                                               Name=str(r.get("ID", "Bjelke")),
                                               ObjectPlacement=placement, Representation=shape)
                assign_material(product, f"{beam_mat} {beam_quality}")
            if level in storey_children:
                storey_children[level].append(product)

    if slab_df is not None and not slab_df.empty:
        for _, r in slab_df.iterrows():
            level = max(int(round(safe_num(r.get("Nivå", 1)))), 1)
            pts = []
            for i in range(1, 9):
                raw = str(r.get(f"P{i} (X,Y)", "") or "")
                nums = [float(x.strip()) / 1000.0 for x in raw.split(",") if x.strip().replace("-", "").replace(".", "").isdigit()]
                if len(nums) == 2:
                    pts.append(tuple(nums))
            if len(pts) >= 3:
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                xmin, xmax = min(xs), max(xs)
                ymin, ymax = min(ys), max(ys)
                width = max(xmax - xmin, 0.001)
                depth_slab = max(ymax - ymin, 0.001)
                z_top = safe_num(r.get("Z [mm]", level * etasjeh * 1000)) / 1000.0 + building_z_offset
                # Senter av dekke i lokal -> verdenskoordinater
                s_lx = xmin + width / 2.0
                s_ly = ymin + depth_slab / 2.0
                s_wx, s_wy = _world_xy(s_lx, s_ly)
                shape = _ifc_product_shape(model, context, width, depth_slab, deck_thk)
                placement = _ifc_local_placement(model, site_placement,
                                                  s_wx, s_wy, z_top - deck_thk,
                                                  _ifc_dir(model, 0, 0, 1),
                                                  _ifc_dir(model, _cos_r, _sin_r, 0.0))
                slab = model.create_entity("IfcSlab", GlobalId=_ifc_guid(),
                                            Name=str(r.get("DeckID", f"Dekke {level}")),
                                            ObjectPlacement=placement, Representation=shape,
                                            PredefinedType="FLOOR")
                assign_material(slab, f"{deck_mat} {deck_quality}")
                if level in storey_children:
                    storey_children[level].append(slab)

    for level, children in storey_children.items():
        if children:
            model.create_entity("IfcRelContainedInSpatialStructure", GlobalId=_ifc_guid(),
                                 RelatedElements=children, RelatingStructure=storeys[level])

    # -----------------------------------------------------------------------
    # GRUNNFORHOLD
    # -----------------------------------------------------------------------
    ground_elements = []

    # Grunnvannsnivå – sentrert under bygget, samme plassering som jordlag
    if gwl_depth_m is not None:
        gwl_z = -gwl_depth_m
        _gwl_local_cx = _x0 + _bx / 2.0
        _gwl_local_cy = _y0 + _by / 2.0
        _gwl_cx = building_offset_x + _gwl_local_cx * _cos_r - _gwl_local_cy * _sin_r
        _gwl_cy = building_offset_y + _gwl_local_cx * _sin_r + _gwl_local_cy * _cos_r
        gwl_placement = _ifc_local_placement(
            model, site_placement,
            _gwl_cx, _gwl_cy, gwl_z,
            _ifc_dir(model, 0.0, 0.0, 1.0),
            _ifc_dir(model, _cos_r, _sin_r, 0.0),
        )
        gwl_shape = _ifc_product_shape(model, context, _bx, _by, 0.05)
        gwl_elem = model.create_entity(
            "IfcBuildingElementProxy", GlobalId=_ifc_guid(),
            Name=f"Grunnvannsnivå (GVN) -{gwl_depth_m:.1f} m",
            ObjectPlacement=gwl_placement, Representation=gwl_shape,
        )
        assign_material(gwl_elem, "Grunnvann")
        ground_elements.append(gwl_elem)

    # Geotekniske lag med farge per jordart
    def _hex_to_ifc_rgb(hex_color: str):
        """Konverterer hex-farge (#RRGGBB) til IFC IfcColourRgb (0.0–1.0)."""
        h = hex_color.lstrip("#")
        r = int(h[0:2], 16) / 255.0
        g = int(h[2:4], 16) / 255.0
        b = int(h[4:6], 16) / 255.0
        return r, g, b

    def _assign_colored_material(product, mat_name: str, hex_color: str):
        """Tilordner materiale med farge (IfcStyledItem) til et produkt."""
        mat = get_material(mat_name)
        try:
            r, g, b = _hex_to_ifc_rgb(hex_color)
            colour = model.create_entity("IfcColourRgb", Name=mat_name, Red=r, Green=g, Blue=b)
            surface_style_rendering = model.create_entity(
                "IfcSurfaceStyleRendering",
                SurfaceColour=colour,
                Transparency=0.15,
                ReflectanceMethod="FLAT",
            )
            surface_style = model.create_entity(
                "IfcSurfaceStyle",
                Name=mat_name,
                Side="BOTH",
                Styles=[surface_style_rendering],
            )
            presentation_style = model.create_entity(
                "IfcPresentationStyleAssignment",
                Styles=[surface_style],
            )
            rep = product.Representation
            if rep:
                for item in rep.Representations:
                    for shape_item in item.Items:
                        model.create_entity(
                            "IfcStyledItem",
                            Item=shape_item,
                            Styles=[presentation_style],
                            Name=mat_name,
                        )
        except Exception:
            pass  # Farge er valgfritt – produktet vises uansett
        model.create_entity("IfcRelAssociatesMaterial", GlobalId=_ifc_guid(),
                             RelatedObjects=[product], RelatingMaterial=mat)

    if ground_layers:
        depth = 0.0
        # Jordlag sentreres under byggets faktiske fotavtrykk i verdenskoordinater
        _local_cx = _x0 + _bx / 2.0
        _local_cy = _y0 + _by / 2.0
        _layer_cx = building_offset_x + _local_cx * _cos_r - _local_cy * _sin_r
        _layer_cy = building_offset_y + _local_cx * _sin_r + _local_cy * _cos_r
        for i, layer in enumerate(ground_layers):
            soil_info = SOIL_DATABASE.get(layer["soil_type"], SOIL_DATABASE["Ukjent"])
            thickness = float(layer["thickness_m"])
            z_top = -depth
            layer_placement = _ifc_local_placement(
                model, site_placement,
                _layer_cx, _layer_cy, z_top - thickness,
                _ifc_dir(model, 0.0, 0.0, 1.0),
                _ifc_dir(model, _cos_r, _sin_r, 0.0),
            )
            layer_shape = _ifc_product_shape(model, context, _bx, _by, thickness)
            soil_elem = model.create_entity(
                "IfcBuildingElementProxy", GlobalId=_ifc_guid(),
                Name=f"Jordlag {i+1}: {soil_info['label']} ({thickness:.1f} m)",
                ObjectPlacement=layer_placement, Representation=layer_shape,
            )
            hex_col = soil_info.get("color", "#AAAAAA")
            _assign_colored_material(soil_elem, soil_info["label"], hex_col)
            ground_elements.append(soil_elem)
            depth += thickness

    # Fundamentplate og peler – beregn verdenskoordinater manuelt
    # slik at de IKKE arver byggets rotasjon og blir skjeve
    if foundation_key and foundation_area_m2 > 0:
        fd_info = FOUNDATION_DATABASE.get(foundation_key, {})
        fd_label = fd_info.get("label", foundation_key)
        fd_thk = 0.3
        # Transformer senter av byggets faktiske fotavtrykk til verdenskoordinater
        _local_cx = _x0 + _bx / 2.0
        _local_cy = _y0 + _by / 2.0
        _fd_cx = building_offset_x + _local_cx * _cos_r - _local_cy * _sin_r
        _fd_cy = building_offset_y + _local_cx * _sin_r + _local_cy * _cos_r
        fd_placement = _ifc_local_placement(
            model, site_placement,
            _fd_cx, _fd_cy, -fd_thk,
            _ifc_dir(model, 0.0, 0.0, 1.0),
            _ifc_dir(model, _cos_r, _sin_r, 0.0),
        )
        fd_shape = _ifc_product_shape(model, context, _bx, _by, fd_thk)
        footing = model.create_entity(
            "IfcFooting", GlobalId=_ifc_guid(),
            Name=fd_label,
            ObjectPlacement=fd_placement, Representation=fd_shape,
            PredefinedType="PAD_FOOTING",
        )
        assign_material(footing, "Betong B35")
        ground_elements.append(footing)

    # Peler – transformer hvert punkt fra byggets lokale system til verdenskoordinater
    if n_piles > 0 and pile_length_m > 0:
        grid_n = max(int(n_piles ** 0.5), 1)
        sp_x = _bx / max(grid_n, 1)
        sp_y = _by / max(grid_n, 1)
        pile_count = 0
        for ix in range(grid_n):
            for iy in range(grid_n):
                if pile_count >= n_piles:
                    break
                _lx = _x0 + (ix + 0.5) * sp_x
                _ly = _y0 + (iy + 0.5) * sp_y
                _wx = building_offset_x + _lx * _cos_r - _ly * _sin_r
                _wy = building_offset_y + _lx * _sin_r + _ly * _cos_r
                pile_placement = _ifc_local_placement(
                    model, site_placement, _wx, _wy, 0.0,
                    _ifc_dir(model, 0.0, 0.0, -1.0),
                    _ifc_dir(model, 1.0, 0.0, 0.0),
                )
                pile_shape = _ifc_product_shape(model, context, 0.3, 0.3, pile_length_m)
                pile_elem = model.create_entity(
                    "IfcPile", GlobalId=_ifc_guid(),
                    Name=f"Pel {pile_count + 1}",
                    ObjectPlacement=pile_placement, Representation=pile_shape,
                    PredefinedType="COHESION",
                )
                assign_material(pile_elem, "Stålpel")
                ground_elements.append(pile_elem)
                pile_count += 1
            if pile_count >= n_piles:
                break

    if ground_elements:
        model.create_entity("IfcRelContainedInSpatialStructure", GlobalId=_ifc_guid(),
                             RelatedElements=ground_elements, RelatingStructure=site)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp:
        temp_path = tmp.name
    try:
        model.write(temp_path)
        with open(temp_path, "rb") as f:
            return f.read()
    finally:
        try:
            os.remove(temp_path)
        except Exception:
            pass



def _nav_to(page):
    st.session_state["_page"] = page
    st.rerun()

if "_nav" in st.session_state:
    st.session_state["_page"] = st.session_state.pop("_nav")

valg = st.session_state.get("_page", "Hjem")

# Alle innstillinger leses fra session state med standardverdier
st.session_state.setdefault("deck_variant_key",     "Hulldekke")
st.session_state.setdefault("concrete_variant_key", "Plasstøpt_betong")
st.session_state.setdefault("wall_variant_key",     "Betong_vegg")
st.session_state.setdefault("_glulam_density",      460.0)
st.session_state.setdefault("_clt_density",         500.0)
st.session_state.setdefault("_use_epd",             True)
st.session_state.setdefault("_show_raw",            False)
st.session_state.setdefault("_fast_mode",           True)
st.session_state.setdefault("_use_geom_fallback",   False)
st.session_state.setdefault("_profile_limit",       float(MAX_PROFILE_OPTIONS_DEFAULT))
st.session_state.setdefault("_lazy_3d",             True)

deck_variant          = st.session_state["deck_variant_key"]
concrete_variant      = st.session_state["concrete_variant_key"]
wall_variant          = st.session_state["wall_variant_key"]
glulam_density        = st.session_state["_glulam_density"]
clt_density           = st.session_state["_clt_density"]
use_epd               = st.session_state["_use_epd"]
show_raw              = st.session_state["_show_raw"]
fast_mode             = st.session_state["_fast_mode"]
use_geometry_fallback = st.session_state["_use_geom_fallback"]
profile_option_limit  = st.session_state["_profile_limit"]
lazy_load_3d          = st.session_state["_lazy_3d"]

uploaded_ifc   = None
uploaded_excel = None

GLULAM_DENSITY = glulam_density
CLT_DENSITY    = clt_density
MATERIAL_DATABASE["Limtre"]["density"]       = glulam_density
MATERIAL_DATABASE["Massivtre"]["density"]    = clt_density
MATERIAL_DATABASE["Massivtre_vegg"]["density"] = clt_density

if valg == "Hjem":
    st.markdown("""
    <style>
    .bt-hero {
        background: #1f4e79;
        border-radius: 16px;
        padding: 2.8rem 2.5rem 2.2rem;
        margin-bottom: 2rem;
    }
    .bt-hero-title { font-size: 2.8rem; font-weight: 600; color: #e6f1fb; margin: 0 0 0.3rem; letter-spacing: -1px; }
    .bt-hero-sub   { font-size: 1.1rem; color: #b5d4f4; margin: 0 0 1.2rem; }
    .bt-badge { display: inline-block; background: rgba(255,255,255,0.13); border: 0.5px solid rgba(255,255,255,0.22); border-radius: 20px; padding: 3px 13px; font-size: 0.8rem; color: #b5d4f4; margin-right: 6px; }
    .bt-navcard {
        background: var(--color-background-primary);
        border: 0.5px solid var(--color-border-tertiary);
        border-radius: 14px;
        padding: 1.3rem 1.2rem 0.6rem;
        height: 100%;
    }
    .bt-navcard-icon { font-size: 2rem; margin-bottom: 0.5rem; }
    .bt-navcard-title { font-size: 1rem; font-weight: 600; color: var(--color-text-primary); margin: 0 0 0.4rem; }
    .bt-navcard-desc { font-size: 0.84rem; color: var(--color-text-secondary); line-height: 1.55; margin: 0 0 0.9rem; }
    .bt-step-row { display: flex; gap: 10px; align-items: flex-start; padding: 0.55rem 0; border-bottom: 0.5px solid var(--color-border-tertiary); }
    .bt-step-num { background: #e6f1fb; color: #0c447c; border-radius: 50%; width: 22px; height: 22px; display:flex; align-items:center; justify-content:center; font-size: 0.75rem; font-weight: 600; flex-shrink: 0; margin-top: 2px; }
    .bt-step-text { font-size: 0.88rem; color: var(--color-text-primary); line-height: 1.5; }
    .bt-step-text b { color: #0c447c; font-weight: 500; }
    </style>

    <div class="bt-hero">
        <div class="bt-hero-title">byggTotal</div>
        <div class="bt-hero-sub">Komplett verktøy for mengdeuttak, kostnad, CO₂ og grunnarbeider</div>
        <span class="bt-badge">Norsk Prisbok 2024</span>
        <span class="bt-badge">NS3420</span>
        <span class="bt-badge">IFC / BIM</span>
        <span class="bt-badge">EPD-faktorer</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### Velg verktøy")
    _r1c1, _r1c2, _r1c3 = st.columns(3)
    _r2c1, _r2c2, _r2c3 = st.columns(3)

    with _r1c1:
        st.markdown('<div class="bt-navcard"><div class="bt-navcard-icon">📊</div><div class="bt-navcard-title">Mengder</div><div class="bt-navcard-desc">Automatisk mengdeuttak fra Excel eller IFC med kostnad og CO₂ per element, type og materiale.</div></div>', unsafe_allow_html=True)
        if st.button("Åpne Mengder", key="nav_mengder", use_container_width=True):
            _nav_to("Mengder")

    with _r1c2:
        st.markdown('<div class="bt-navcard"><div class="bt-navcard-icon">🏗️</div><div class="bt-navcard-title">Bygggenerator</div><div class="bt-navcard-desc">Generer et parametrisk bygg fra scratch — definer etasjer, fagverk og materialer, eksporter til IFC.</div></div>', unsafe_allow_html=True)
        if st.button("Åpne Bygggenerator", key="nav_bg", use_container_width=True):
            _nav_to("Bygggenerator")

    with _r1c3:
        st.markdown('<div class="bt-navcard"><div class="bt-navcard-icon">🌍</div><div class="bt-navcard-title">Grunn</div><div class="bt-navcard-desc">Terrenganalyse, geoteknikk, fundamentering og peling med CO₂-regnskap og stikningsdata-import.</div></div>', unsafe_allow_html=True)
        if st.button("Åpne Grunn", key="nav_grunn", use_container_width=True):
            _nav_to("Grunn")

    st.markdown("")

    with _r2c1:
        st.markdown('<div class="bt-navcard"><div class="bt-navcard-icon">🔁</div><div class="bt-navcard-title">Materialbytte</div><div class="bt-navcard-desc">Sammenlign kostnad og CO₂-konsekvens ved materialbytte direkte i modellen, med IFC-eksport.</div></div>', unsafe_allow_html=True)
        if st.button("Åpne Materialbytte", key="nav_mb", use_container_width=True):
            _nav_to("Materialbytte")

    with _r2c2:
        st.markdown('<div class="bt-navcard"><div class="bt-navcard-icon">🧊</div><div class="bt-navcard-title">3D-modell</div><div class="bt-navcard-desc">Interaktiv 3D-visning av IFC-modellen i nettleseren. Filtrer og fremhev elementer fritt.</div></div>', unsafe_allow_html=True)
        if st.button("Åpne 3D-modell", key="nav_3d", use_container_width=True):
            _nav_to("3D-modell")

    with _r2c3:
        st.markdown('<div class="bt-navcard"><div class="bt-navcard-icon">📝</div><div class="bt-navcard-title">Rapport</div><div class="bt-navcard-desc">Komplett prosjektrapport med alle nøkkeltall fra konstruksjon og grunnarbeider. Word eller PDF.</div></div>', unsafe_allow_html=True)
        if st.button("Åpne Rapport", key="nav_rapport", use_container_width=True):
            _nav_to("Rapport")

    st.markdown("")
    _r3c1, _r3c2, _r3c3 = st.columns(3)
    with _r3c1:
        st.markdown('<div class="bt-navcard"><div class="bt-navcard-icon">⚙️</div><div class="bt-navcard-title">Innstillinger</div><div class="bt-navcard-desc">Juster materialtetthet, CO₂-kilde, ytelsesalternativer og standardprodukter fra Norsk Prisbok.</div></div>', unsafe_allow_html=True)
        if st.button("Åpne Innstillinger", key="nav_innstillinger", use_container_width=True):
            _nav_to("Innstillinger")

    st.markdown("---")

    _qs1, _qs2 = st.columns([1.2, 1])
    with _qs1:
        st.markdown("#### Last opp prosjektfil")
        _hjem_file = st.file_uploader(
            "Dra inn Excel- eller IFC-fil her",
            type=["xlsx", "ifc"],
            label_visibility="collapsed",
            help="Støtter Excel (.xlsx) og IFC (.ifc)",
        )
        if _hjem_file is not None:
            _bytes = _hjem_file.getvalue()
            st.session_state["_filename"] = _hjem_file.name
            if _hjem_file.name.lower().endswith(".ifc"):
                st.session_state["_ifc_bytes"]   = _bytes
                st.session_state["_excel_bytes"] = None
            else:
                st.session_state["_excel_bytes"] = _bytes
                st.session_state["_ifc_bytes"]   = None
            _nav_to("Mengder")

        _active_fname = st.session_state.get("_filename")
        if _active_fname:
            st.success(f"Aktiv fil: **{_active_fname}**")
            _hjem_c1, _hjem_c2 = st.columns(2)
            with _hjem_c1:
                if st.button("→ Gå til Mengder", type="primary", use_container_width=True):
                    _nav_to("Mengder")
            with _hjem_c2:
                if st.button("🗑️ Fjern fil", use_container_width=True):
                    for _k in ["_filename", "_ifc_bytes", "_excel_bytes"]:
                        st.session_state.pop(_k, None)
                    st.rerun()
        else:
            st.caption("Støtter Excel (.xlsx) og IFC (.ifc). Eller start med Bygggenerator uten fil.")

    with _qs2:
        st.markdown("#### Kom i gang")
        st.markdown("""
        <div class="bt-step-row"><div class="bt-step-num">1</div><div class="bt-step-text">Last opp en <b>Excel</b>- eller <b>IFC</b>-fil til venstre</div></div>
        <div class="bt-step-row"><div class="bt-step-num">2</div><div class="bt-step-text">Klikk <b>Mengder</b> for mengdeuttak, kostnad og CO₂-oversikt</div></div>
        <div class="bt-step-row"><div class="bt-step-num">3</div><div class="bt-step-text">Bruk <b>Materialbytte</b> for å sammenligne materialalternativer</div></div>
        <div class="bt-step-row"><div class="bt-step-num">4</div><div class="bt-step-text">Analyser grunnforhold i <b>Grunn</b></div></div>
        <div class="bt-step-row" style="border-bottom:none"><div class="bt-step-num">5</div><div class="bt-step-text">Last ned <b>Rapport</b> som Word eller PDF</div></div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.caption("byggTotal · Norsk Prisbok 2024 · EPD-faktorer · NS3420 · IFC/BIM")
    st.stop()

data = None
nodes = pd.DataFrame()
forside = pd.DataFrame()
ifc_bytes   = st.session_state.get("_ifc_bytes")
excel_bytes = st.session_state.get("_excel_bytes")
filename    = st.session_state.get("_filename")

try:
    if ifc_bytes is not None:
        with st.spinner("Behandler IFC-fil..."):
            data, nodes, forside = build_dataset_from_ifc(
                ifc_bytes,
                use_geometry_fallback=use_geometry_fallback,
                fast_mode=fast_mode,
            )
    elif excel_bytes is not None:
        try:
            data, nodes, forside = build_dataset_from_excel(excel_bytes)
        except Exception:
            data = pd.DataFrame(columns=["Segment", "Type", "Knutepunkter", "Material / Tverrsnitt", "Lengde [m]", "Areal [m2]", "Volum [m3]", "Vekt [kg]", "materiale", "Endret IFC", "Mengdegrunnlag"])
            nodes = pd.DataFrame()
            forside = pd.DataFrame()
    else:
        _empty_cols_base = ["Segment", "Type", "Knutepunkter", "Material / Tverrsnitt", "Lengde [m]", "Areal [m2]", "Volum [m3]", "Vekt [kg]", "materiale", "Endret IFC", "Mengdegrunnlag"]
        if valg == "Bygggenerator":
            filename = "Generert bygg"
            data = pd.DataFrame(columns=_empty_cols_base + ["Materialkvalitet", "Kostnad [kr]", "CO2 [kgCO2e]"])
        else:
            filename = filename or "Ingen fil lastet"
            data = pd.DataFrame(columns=_empty_cols_base)
        nodes = pd.DataFrame()
        forside = pd.DataFrame()
except Exception as e:
    st.error(f"Kunne ikke lese filen: {e}")
    st.stop()

if filename and filename not in ("Ingen fil lastet", "Generert bygg"):
    st.success(f"Aktiv fil: **{filename}**")

for col in ["Segment", "Type", "Knutepunkter", "Material / Tverrsnitt", "Lengde [m]", "Areal [m2]", "Volum [m3]", "Vekt [kg]", "materiale", "Endret IFC", "Mengdegrunnlag"]:
    if col not in data.columns:
        data[col] = pd.NA

if not data.empty:
    data["Produktnøkkel"] = data.apply(lambda row: detect_product_key(row, deck_variant, concrete_variant, wall_variant), axis=1)
    data["Produktnavn"] = data["Produktnøkkel"].apply(lambda key: MATERIAL_DATABASE.get(key, MATERIAL_DATABASE["Ukjent"])["label"])
    data["NS3420-kode"] = data.apply(map_ns3420_code, axis=1)
    data["Kostnad [kr]"] = data.apply(lambda row: cost_for_row(row, deck_variant, concrete_variant, wall_variant), axis=1)
    data["CO2 [kgCO2e]"] = data.apply(lambda row: co2_for_row(row, deck_variant, concrete_variant, wall_variant, use_epd=use_epd), axis=1)
else:
    for col in ["Produktnøkkel", "Produktnavn", "NS3420-kode", "Kostnad [kr]", "CO2 [kgCO2e]"]:
        data[col] = []

param = {}
if not forside.empty:
    for _, row in forside.iterrows():
        if len(row) >= 2 and pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
            param[str(row.iloc[0]).strip()] = row.iloc[1]

type_options = sorted([x for x in data["Type"].dropna().astype(str).unique().tolist()]) if not data.empty else []
mat_options = sorted([x for x in data["materiale"].dropna().astype(str).unique().tolist()]) if not data.empty else []
all_profile_options = sorted([x for x in data["Material / Tverrsnitt"].dropna().astype(str).unique().tolist()]) if not data.empty else []
profile_options = all_profile_options[:int(profile_option_limit)]
length_series = pd.to_numeric(data.get("Lengde [m]", pd.Series(dtype=float)), errors="coerce")
length_series = length_series[(length_series.notna()) & (length_series >= 0)]
reasonable_lengths = length_series[length_series <= 1000]
slider_source = reasonable_lengths if not reasonable_lengths.empty else length_series
max_length = float(slider_source.max() or 0) if not slider_source.empty else 0.0

if valg in ["Mengder", "Materialbytte"]:
    with st.container():
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            selected_types = st.multiselect("Type", type_options, default=type_options)
        with c2:
            selected_materials = st.multiselect("Materiale", mat_options, default=mat_options)
        with c3:
            selected_profiles = st.multiselect("Profil / tverrsnitt", profile_options, default=profile_options[:8] if len(profile_options) > 8 else profile_options, help=f"Viser de første {int(profile_option_limit)} unike profilene for raskere lasting.")
        with c4:
            length_range = st.slider("Lengdeintervall [m]", 0.0, max(1.0, max_length), (0.0, max(1.0, max_length)))
    st.caption("Tips: Profilfilteret viser nå alle profiler som standard, og lengdeslideren ignorerer urimelig store verdier når maksgrensen settes.")
else:
    selected_types = type_options
    selected_materials = mat_options
    selected_profiles = profile_options
    length_range = (0.0, max(1.0, max_length))

filtered = data.copy()
if not filtered.empty:
    if selected_types:
        filtered = filtered[filtered["Type"].isin(selected_types)]
    if selected_materials:
        filtered = filtered[filtered["materiale"].isin(selected_materials)]
    if selected_profiles:
        filtered = filtered[filtered["Material / Tverrsnitt"].isin(selected_profiles)]
    length_values = pd.to_numeric(filtered["Lengde [m]"], errors="coerce")
    filtered = filtered[(length_values.fillna(0) >= length_range[0]) & (length_values.fillna(0) <= length_range[1])]
else:
    filtered = pd.DataFrame(columns=data.columns)

summary = (
    filtered.groupby(["Type", "materiale"], dropna=False)
    .agg(antall=("Segment", "count"), areal_m2=("Areal [m2]", "sum"), lengde_m=("Lengde [m]", "sum"), volum_m3=("Volum [m3]", "sum"), vekt_kg=("Vekt [kg]", "sum"), kostnad_kr=("Kostnad [kr]", "sum"), co2_kg=("CO2 [kgCO2e]", "sum"))
    .reset_index()
    .sort_values(["Type", "materiale"])
    if not filtered.empty else pd.DataFrame()
)

material_summary = (
    filtered.groupby(["materiale", "Produktnavn", "NS3420-kode"], dropna=False)
    .agg(antall=("Segment", "count"), areal_m2=("Areal [m2]", "sum"), lengde_m=("Lengde [m]", "sum"), volum_m3=("Volum [m3]", "sum"), vekt_kg=("Vekt [kg]", "sum"), kostnad_kr=("Kostnad [kr]", "sum"), co2_kg=("CO2 [kgCO2e]", "sum"))
    .reset_index()
    .sort_values("kostnad_kr", ascending=False)
    if not filtered.empty else pd.DataFrame()
)

swap_df = pd.DataFrame()

if valg == "Mengder":
    if st.button("← Hjem", key="hjem_mengder"): _nav_to("Hjem")
    st.header("📊 Mengder")
    with st.expander("🛒 Produktvalg fra Norsk Prisbok", expanded=False):
        _mc1, _mc2, _mc3 = st.columns(3)
        with _mc1:
            _dv2 = st.selectbox("Dekkeløsning", ["Hulldekke","Hulldekke_lavCO2"],
                index=["Hulldekke","Hulldekke_lavCO2"].index(st.session_state["deck_variant_key"]),
                format_func=lambda x: MATERIAL_DATABASE[x]["label"], key="deck_mengder")
            if _dv2 != st.session_state["deck_variant_key"]:
                st.session_state["deck_variant_key"] = _dv2; st.rerun()
        with _mc2:
            _cv2 = st.selectbox("Plasstøpt betong", ["Plasstøpt_betong","Plasstøpt_betong_lavCO2"],
                index=["Plasstøpt_betong","Plasstøpt_betong_lavCO2"].index(st.session_state["concrete_variant_key"]),
                format_func=lambda x: MATERIAL_DATABASE[x]["label"], key="concrete_mengder")
            if _cv2 != st.session_state["concrete_variant_key"]:
                st.session_state["concrete_variant_key"] = _cv2; st.rerun()
        with _mc3:
            _wv2 = st.selectbox("Betongvegg", ["Betong_vegg","Betong_vegg_lavCO2"],
                index=["Betong_vegg","Betong_vegg_lavCO2"].index(st.session_state["wall_variant_key"]),
                format_func=lambda x: MATERIAL_DATABASE[x]["label"], key="wall_mengder")
            if _wv2 != st.session_state["wall_variant_key"]:
                st.session_state["wall_variant_key"] = _wv2; st.rerun()
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1:
        metric_card("Elementer", f"{len(filtered):,}".replace(",", " "))
    with k2:
        metric_card("Total lengde", f"{pd.to_numeric(filtered['Lengde [m]'], errors='coerce').fillna(0).sum():,.1f} m".replace(",", " "))
    with k3:
        metric_card("Total areal", f"{pd.to_numeric(filtered['Areal [m2]'], errors='coerce').fillna(0).sum():,.1f} m²".replace(",", " "))
    with k4:
        steel_sum = filtered.loc[filtered["materiale"] == "Stål", "Vekt [kg]"].sum() if not filtered.empty else 0
        metric_card("Stålvekt", f"{steel_sum:,.0f} kg".replace(",", " "))
    with k5:
        metric_card("Estimert kostnad", f"{pd.to_numeric(filtered['Kostnad [kr]'], errors='coerce').fillna(0).sum():,.0f} kr".replace(",", " "))
    with k6:
        metric_card("CO₂-avtrykk", f"{pd.to_numeric(filtered['CO2 [kgCO2e]'], errors='coerce').fillna(0).sum():,.0f} kgCO₂e".replace(",", " "))

    left, right = st.columns([1.2, 1])
    with left:
        st.subheader("Oppsummering per type og materiale")
        st.dataframe(summary, use_container_width=True, hide_index=True)
        st.subheader("Oppsummering per profil / tverrsnitt")
        profiles = (
            filtered.groupby(["Material / Tverrsnitt", "Produktnavn", "Mengdegrunnlag"], dropna=False)
            .agg(antall=("Segment", "count"), areal_m2=("Areal [m2]", "sum"), lengde_m=("Lengde [m]", "sum"), kostnad_kr=("Kostnad [kr]", "sum"), co2_kg=("CO2 [kgCO2e]", "sum"))
            .reset_index()
            .sort_values("kostnad_kr", ascending=False)
        ) if not filtered.empty else pd.DataFrame()
        st.dataframe(profiles, use_container_width=True, hide_index=True)

    with right:
        st.subheader("Kostnad og CO₂ per produkt")
        pie_data = summary[summary["kostnad_kr"] > 0].copy() if not summary.empty else pd.DataFrame()
        co2_data = filtered.groupby("Produktnavn", dropna=False)["CO2 [kgCO2e]"].sum().reset_index() if not filtered.empty else pd.DataFrame()
        co2_data = co2_data[co2_data["CO2 [kgCO2e]"] > 0] if not co2_data.empty else co2_data

        if not pie_data.empty or not co2_data.empty:
            import plotly.subplots as ps
            fig_combo = ps.make_subplots(
                rows=1, cols=2,
                specs=[[{"type": "pie"}, {"type": "bar"}]],
                subplot_titles=["Kostnadsfordeling", "CO₂ per produkt"],
            )
            if not pie_data.empty:
                pie_data["navn"] = pie_data["Type"].fillna("Ukjent") + " – " + pie_data["materiale"].fillna("Ukjent")
                fig_combo.add_trace(go.Pie(labels=pie_data["navn"], values=pie_data["kostnad_kr"], textinfo="percent", hovertemplate="%{label}<br>%{value:,.0f} kr<extra></extra>"), row=1, col=1)
            if not co2_data.empty:
                fig_combo.add_trace(go.Bar(x=co2_data["Produktnavn"].fillna("Ukjent"), y=co2_data["CO2 [kgCO2e]"], marker_color="#2ecc71", hovertemplate="%{x}<br>%{y:,.0f} kgCO₂e<extra></extra>"), row=1, col=2)
            fig_combo.update_layout(height=420, showlegend=False, margin=dict(t=40, b=10, l=10, r=10))
            fig_combo.update_xaxes(tickangle=-30, row=1, col=2)
            st.plotly_chart(fig_combo, use_container_width=True)
        else:
            st.info("Ingen data er tilgjengelige for valgt utvalg.")

    show_cols = [c for c in ["Segment", "Type", "Knutepunkter", "Material / Tverrsnitt", "materiale", "Produktnøkkel", "Produktnavn", "NS3420-kode", "Mengdegrunnlag", "Endret IFC", "Lengde [m]", "Areal [m2]", "Volum [m3]", "Vekt [kg]", "Kostnad [kr]", "CO2 [kgCO2e]", "IFC Type", "IFC GlobalId"] if c in filtered.columns]
    st.subheader("Filtrerte elementer")
    st.dataframe(filtered[show_cols], use_container_width=True, hide_index=True)
    st.download_button("Last ned filtrerte data som CSV", filtered[show_cols].to_csv(index=False).encode("utf-8-sig"), file_name="filtrerte_mengder.csv", mime="text/csv")
    if show_raw:
        with st.expander("Rådata"):
            st.dataframe(data, use_container_width=True)

elif valg == "Grunn":
    from ground_module import (
        SOIL_DATABASE, FOUNDATION_DATABASE, GROUND_CO2_DATABASE,
        recommend_foundation, estimate_pile_length, calculate_foundation_cost,
        calculate_ground_co2, load_geojson, load_dxf_points,
        plot_soil_profile, plot_bearing_capacity_chart, groundwater_risk_assessment,
        build_ground_report_df,
    )

    if st.button("← Hjem", key="hjem_grunn"): _nav_to("Hjem")
    st.header("🌍 Grunn og georeferering")
    active_breeam_level = st.session_state.get("breeam_target_level", "Ingen") if st.session_state.get("breeam_active", False) else "Ingen"

    ground_tab1, ground_tab2, ground_tab3, ground_tab4, ground_tab5 = st.tabs([
        "📍 Stikningsdata / terreng",
        "🗺️ Georeferering",
        "🪨 Geoteknikk og profil",
        "🏗️ Fundamentering og peling",
        "🌿 CO₂-regnskap grunn",
    ])

    # -----------------------------------------------------------------------
    # TAB 1 – STIKNINGSDATA OG TERRENG
    # -----------------------------------------------------------------------
    with ground_tab1:
        st.subheader("Stikningsdata og terrengmodell")
        st.caption("Last opp stikningspunkter (CSV, Excel, TXT, PTS). Appen beregner terrengvolum, tomteareal og første anslag for masser.")

        stake_file = st.file_uploader("Last opp stikningsdata (.csv, .xlsx, .txt, .pts)", type=["csv", "xlsx", "xls", "txt", "pts"], key="stake_file_v2")

        if stake_file is None:
            st.info("Last opp stikningsdata for å aktivere terrengmodulen.")
        else:
            try:
                stake_df = load_stake_data(stake_file)
            except Exception as e:
                st.error(f"Kunne ikke lese stikningsdata: {e}")
                stake_df = pd.DataFrame()

            if not stake_df.empty:
                default_target = float(stake_df["Z"].mean())
                c1, c2, c3 = st.columns(3)
                with c1:
                    use_mean_level = st.toggle("Bruk middelkote som prosjektkote", value=True, key="ground_use_mean_v2")
                with c2:
                    target_elevation = st.number_input("Prosjektkote [m]", value=default_target, step=0.10, format="%.2f", disabled=use_mean_level)
                with c3:
                    mass_factor = st.number_input("Usikkerhets-/massefaktor", min_value=0.50, max_value=2.00, value=1.15, step=0.05)
                if use_mean_level:
                    target_elevation = default_target

                summary_ground, evaluated_points, hull = build_ground_summary(stake_df, target_elevation=target_elevation, mass_factor=mass_factor)

                p1, p2, p3, p4, p5, p6 = st.columns(6)
                with p1: metric_card("Punkt", f"{summary_ground['Antall punkt']:,}".replace(",", " "))
                with p2: metric_card("Tomteareal", f"{summary_ground['Tomteareal (konveks hull)']:,.1f} m²".replace(",", " "))
                with p3: metric_card("Prosjektkote", f"{summary_ground['Prosjektkote']:.2f} m")
                with p4: metric_card("Utgraving", f"{summary_ground['Estimert utgraving']:,.1f} m³".replace(",", " "))
                with p5: metric_card("Oppfylling", f"{summary_ground['Estimert oppfylling']:,.1f} m³".replace(",", " "))
                with p6: metric_card("Punktavstand", f"{summary_ground['Punktavstand ca.']:.2f} m")
                # Lagre til rapport
                st.session_state["rapport_tomteareal"] = f"{summary_ground['Tomteareal (konveks hull)']:,.1f} m²".replace(",", " ")
                st.session_state["rapport_utgraving"] = f"{summary_ground['Estimert utgraving']:,.1f} m³".replace(",", " ")
                st.session_state["rapport_oppfylling"] = f"{summary_ground['Estimert oppfylling']:,.1f} m³".replace(",", " ")
                st.session_state["rapport_prosjektkote"] = f"{summary_ground['Prosjektkote']:.2f} m"
                st.session_state["rapport_antall_punkt"] = str(summary_ground['Antall punkt'])

                left, right = st.columns([1.1, 1])
                with left:
                    st.subheader("Tomteutbredelse")
                    st.pyplot(plot_ground_points(evaluated_points, hull))
                    summary_df = pd.DataFrame({"Parameter": list(summary_ground.keys()), "Verdi": list(summary_ground.values())})
                    st.dataframe(summary_df, use_container_width=True, hide_index=True)
                with right:
                    st.subheader("Punktfordeling per kode")
                    code_df = evaluated_points.groupby("Kode", dropna=False).size().reset_index(name="Antall")
                    st.dataframe(code_df, use_container_width=True, hide_index=True)

                st.subheader("Grunnjobbscenario")
                g1, g2, g3 = st.columns(3)
                with g1:
                    current_ground_system = st.selectbox("Basis scenario", list(GROUND_SYSTEM_LIBRARY.keys()), index=0)
                with g2:
                    target_ground_system = st.selectbox("Alternativt scenario", list(GROUND_SYSTEM_LIBRARY.keys()), index=min(1, len(GROUND_SYSTEM_LIBRARY)-1))
                with g3:
                    ground_breeam_level = st.selectbox("BREEAM-nivå", BREEAM_LEVELS, index=BREEAM_LEVELS.index(active_breeam_level) if active_breeam_level in BREEAM_LEVELS else 0)

                current_price_df = build_ground_pricing_basis_v2(summary_ground, current_ground_system, ground_breeam_level)
                target_price_df = build_ground_pricing_basis_v2(summary_ground, target_ground_system, ground_breeam_level)
                ground_compare_df = compare_ground_scenarios(summary_ground, current_ground_system, target_ground_system, ground_breeam_level)

                m1, m2, m3, m4 = st.columns(4)
                with m1: metric_card("Basis kostnad", f"{current_price_df['Beløp'].sum():,.0f} kr".replace(",", " "))
                with m2: metric_card("Alternativt", f"{target_price_df['Beløp'].sum():,.0f} kr".replace(",", " "))
                with m3: metric_card("Endring", f"{(target_price_df['Beløp'].sum() - current_price_df['Beløp'].sum()):,.0f} kr".replace(",", " "))
                with m4: metric_card("BREEAM", ground_breeam_level)

                cl, cr = st.columns(2)
                with cl:
                    st.markdown("**Prisgrunnlag – basis**")
                    st.dataframe(current_price_df, use_container_width=True, hide_index=True)
                with cr:
                    st.markdown("**Prisgrunnlag – alternativt**")
                    st.dataframe(target_price_df, use_container_width=True, hide_index=True)

                st.dataframe(ground_compare_df, use_container_width=True, hide_index=True)
                st.dataframe(evaluated_points, use_container_width=True, hide_index=True, height=320)

                dl1, dl2, dl3 = st.columns(3)
                with dl1:
                    st.download_button("Last ned punktgrunnlag CSV", evaluated_points.to_csv(index=False).encode("utf-8-sig"), file_name="grunn_punktgrunnlag.csv", mime="text/csv")
                with dl2:
                    st.download_button("Last ned prisgrunnlag CSV", target_price_df.to_csv(index=False).encode("utf-8-sig"), file_name="grunn_prisgrunnlag.csv", mime="text/csv")
                with dl3:
                    try:
                        obj_bytes = generate_ground_obj(evaluated_points)
                        st.download_button("Last ned terrengmodell OBJ", obj_bytes, file_name="terrengmodell.obj", mime="text/plain")
                    except Exception as e:
                        st.info(f"OBJ-modell ikke tilgjengelig: {e}")

    # -----------------------------------------------------------------------
    # TAB 2 – GEOREFERERING
    # -----------------------------------------------------------------------
    with ground_tab2:
        st.subheader("Georeferering – importer punktdata")
        st.caption("Støtter GeoJSON, DXF (AutoCAD) og stikningsdata (CSV/Excel).")

        geo_format = st.radio("Velg format", ["GeoJSON", "DXF (AutoCAD)", "Stikningsdata (CSV/Excel)"], horizontal=True)

        geo_df = pd.DataFrame()
        geo_meta = {}

        if geo_format == "GeoJSON":
            geo_file = st.file_uploader("Last opp GeoJSON-fil", type=["geojson", "json"], key="geo_geojson")
            if geo_file:
                try:
                    geo_df, geo_meta = load_geojson(geo_file)
                    st.success(f"Leste {geo_meta['antall_punkt']} punkter fra GeoJSON.")
                    if geo_meta.get("crs") and geo_meta["crs"] != "Ukjent CRS":
                        st.info(f"CRS: {geo_meta['crs']}")
                except Exception as e:
                    st.error(f"Feil ved lesing av GeoJSON: {e}")

        elif geo_format == "DXF (AutoCAD)":
            dxf_file = st.file_uploader("Last opp DXF-fil", type=["dxf"], key="geo_dxf")
            if dxf_file:
                try:
                    geo_df = load_dxf_points(dxf_file)
                    st.success(f"Leste {len(geo_df)} punkter fra DXF.")
                except ImportError:
                    st.warning("DXF-støtte krever pakken `ezdxf`. Legg til `ezdxf` i requirements.txt og installer på nytt med `py -m pip install ezdxf`.")
                except Exception as e:
                    st.error(f"Feil ved lesing av DXF: {e}")

        elif geo_format == "Stikningsdata (CSV/Excel)":
            stake_geo_file = st.file_uploader("Last opp stikningsdata", type=["csv", "xlsx", "xls", "txt", "pts"], key="geo_stake")
            if stake_geo_file:
                try:
                    geo_df = load_stake_data(stake_geo_file)
                    st.success(f"Leste {len(geo_df)} punkter.")
                except Exception as e:
                    st.error(f"Feil: {e}")

        if not geo_df.empty:
            st.subheader("Importerte punkter")
            g1, g2, g3, g4 = st.columns(4)
            with g1: metric_card("Punkter", str(len(geo_df)))
            with g2: metric_card("X-spenn", f"{geo_df['X'].max() - geo_df['X'].min():.1f} m")
            with g3: metric_card("Y-spenn", f"{geo_df['Y'].max() - geo_df['Y'].min():.1f} m")
            with g4: metric_card("Z-min / maks", f"{geo_df['Z'].min():.1f} / {geo_df['Z'].max():.1f} m")

            fig_geo, ax_geo = plt.subplots(figsize=(7, 5))
            sc = ax_geo.scatter(geo_df["X"], geo_df["Y"], c=geo_df["Z"], s=15, cmap="terrain")
            fig_geo.colorbar(sc, ax=ax_geo, label="Z / kote [m]")
            ax_geo.set_xlabel("X (øst)")
            ax_geo.set_ylabel("Y (nord)")
            ax_geo.set_title("Georefererte punkter")
            ax_geo.axis("equal")
            st.pyplot(fig_geo)
            plt.close(fig_geo)

            st.dataframe(geo_df.head(200), use_container_width=True, hide_index=True)
            st.download_button("Last ned som CSV", geo_df.to_csv(index=False).encode("utf-8-sig"), file_name="geo_punkter.csv", mime="text/csv")

    # -----------------------------------------------------------------------
    # TAB 3 – GEOTEKNIKK OG PROFIL
    # -----------------------------------------------------------------------
    with ground_tab3:
        st.subheader("Geoteknisk lagprofil")
        st.caption("Definer jordlag fra terreng og ned. Appen tegner geoteknisk profil og beregner bæreevne.")

        n_layers = st.number_input("Antall jordlag", min_value=1, max_value=8, value=3, step=1)
        layers = []
        layer_cols = st.columns(min(int(n_layers), 4))
        for i in range(int(n_layers)):
            col = layer_cols[i % 4]
            with col:
                st.markdown(f"**Lag {i+1}**")
                soil_type = st.selectbox(f"Jordart", list(SOIL_DATABASE.keys()), key=f"soil_type_{i}", index=min(i, len(SOIL_DATABASE)-1))
                thickness = st.number_input(f"Tykkelse [m]", min_value=0.1, max_value=50.0, value=2.0 + i * 1.0, step=0.5, key=f"soil_thick_{i}")
                layers.append({"soil_type": soil_type, "thickness_m": thickness})

        gw_col1, gw_col2 = st.columns(2)
        with gw_col1:
            gwl_active = st.toggle("Angi grunnvannsnivå", value=True, key="gwl_active")
        with gw_col2:
            gwl_depth = st.number_input("Grunnvannsnivå – dybde fra terreng [m]", min_value=0.0, max_value=50.0, value=2.5, step=0.25, disabled=not gwl_active)

        total_depth = sum(l["thickness_m"] for l in layers)

        prof_col, prop_col = st.columns([1, 2])
        with prof_col:
            st.subheader("Geoteknisk profil")
            fig_prof = plot_soil_profile(layers, gwl_depth if gwl_active else None)
            st.pyplot(fig_prof)
            plt.close(fig_prof)

        with prop_col:
            st.subheader("Laginformasjon")
            report_df = build_ground_report_df(layers, gwl_depth if gwl_active else None,
                                               "Platefundament_betong", 0, 0, 0, 0, 0)
            st.dataframe(report_df, use_container_width=True, hide_index=True)

            st.subheader("Bæreevne per jordart")
            fig_bc = plot_bearing_capacity_chart(list(SOIL_DATABASE.keys()), 100.0)
            st.pyplot(fig_bc)
            plt.close(fig_bc)

        if gwl_active:
            st.subheader("Grunnvannsrisikovurdering")
            foundation_depth_input = st.number_input("Planlagt fundamentdybde [m]", min_value=0.1, max_value=20.0, value=1.5, step=0.25)
            gwl_risk = groundwater_risk_assessment(gwl_depth, foundation_depth_input)
            r1, r2, r3 = st.columns(3)
            with r1: metric_card("Grunnvannsnivå", f"{gwl_risk['grunnvannsnivå_m']:.1f} m")
            with r2: metric_card("Fundamentdybde", f"{gwl_risk['fundamentdybde_m']:.1f} m")
            with r3: metric_card("Margin", f"{gwl_risk['margin_m']:.2f} m")
            color_map = {"Lav": "success", "Middels": "warning", "Høy": "error", "Kritisk": "error"}
            getattr(st, color_map.get(gwl_risk["risikonivå"], "info"))(f"**Risiko: {gwl_risk['risikonivå']}** – {gwl_risk['merknad']}")

        st.download_button("Last ned lagprofil CSV", report_df.to_csv(index=False).encode("utf-8-sig"), file_name="geoteknikk_profil.csv", mime="text/csv")

    # -----------------------------------------------------------------------
    # TAB 4 – FUNDAMENTERING OG PELING
    # -----------------------------------------------------------------------
    with ground_tab4:
        st.subheader("Fundamentering og peling")
        st.caption("Velg jordart, last og byggets areal. Appen anbefaler fundamenteringstype og estimerer kostnader.")

        fa1, fa2, fa3 = st.columns(3)
        with fa1:
            found_soil = st.selectbox("Dominerende jordart", list(SOIL_DATABASE.keys()), key="found_soil")
        with fa2:
            building_area = st.number_input("Bygningens grunnflate [m²]", min_value=10.0, max_value=10000.0, value=200.0, step=10.0)
        with fa3:
            total_load_kN = st.number_input("Estimert total last [kN]", min_value=100.0, max_value=500000.0, value=building_area * 10.0, step=100.0,
                                             help="Ca. 8–15 kN/m² for boligbygg, 15–25 kN/m² for næringsbygg")

        rec = recommend_foundation(found_soil, building_area, total_load_kN)
        # Lagre til rapport
        st.session_state["rapport_jordart"] = SOIL_DATABASE[found_soil]["label"]
        st.session_state["rapport_bæreevne"] = f"{rec['bæreevne_kPa']:.0f} kPa"
        st.session_state["rapport_fundament_anbefalt"] = rec["label"]
        st.session_state["rapport_bygningsareal"] = f"{building_area:,.0f} m²".replace(",", " ")
        st.session_state["rapport_total_last"] = f"{total_load_kN:,.0f} kN".replace(",", " ")

        r1, r2, r3 = st.columns(3)
        with r1: metric_card("Anbefalt", rec["label"])
        with r2: metric_card("Bæreevne", f"{rec['bæreevne_kPa']:.0f} kPa")
        with r3: metric_card("Nødv. areal", f"{rec['nødvendig_areal_m2']} m²")
        st.info(f"**Begrunnelse:** {rec['begrunnelse']}")

        st.subheader("Beregn fundamentkostnad")
        f1, f2 = st.columns(2)
        with f1:
            chosen_foundation = st.selectbox("Fundamenteringstype", list(FOUNDATION_DATABASE.keys()),
                                              format_func=lambda k: FOUNDATION_DATABASE[k]["label"])
        with f2:
            fd_info = FOUNDATION_DATABASE[chosen_foundation]
            found_qty = st.number_input(f"Mengde [{fd_info['unit']}]", min_value=1.0, max_value=100000.0, value=building_area, step=10.0)

        found_result = calculate_foundation_cost(chosen_foundation, found_qty)
        fc1, fc2, fc3 = st.columns(3)
        with fc1: metric_card("Kostnad", f"{found_result['total_kostnad_kr']:,.0f} kr".replace(",", " "))
        with fc2: metric_card("CO₂", f"{found_result['total_co2_kgCO2e']:,.0f} kgCO₂e".replace(",", " "))
        with fc3: metric_card("Enhetspris", f"{found_result['enhetspris_kr']:,.0f} kr/{fd_info['unit']}".replace(",", " "))
        st.caption(fd_info["description"])

        st.subheader("Pelekalkulator")
        needs_piling = found_soil in ["Leire", "Kvikkleire", "Torv", "Silt"]
        if needs_piling:
            st.warning(f"Jordart **{SOIL_DATABASE[found_soil]['label']}** har lav bæreevne. Peling til fjell eller fast lag er sannsynlig.")

        p1, p2, p3 = st.columns(3)
        with p1:
            depth_to_rock = st.number_input("Dybde til fjell / fast lag [m]", min_value=0.5, max_value=80.0, value=8.0, step=0.5)
        with p2:
            n_piles = st.number_input("Antall peler", min_value=1, max_value=500, value=max(4, int(building_area / 16)), step=1)
        with p3:
            pile_soil = st.selectbox("Jordart rundt peler", list(SOIL_DATABASE.keys()), key="pile_soil", index=list(SOIL_DATABASE.keys()).index(found_soil))

        pile_est = estimate_pile_length(depth_to_rock, pile_soil)
        total_lm = pile_est["anbefalt_pellengde_m"] * n_piles
        # Lagre til rapport
        st.session_state["rapport_peler_antall"] = str(int(n_piles))
        st.session_state["rapport_peler_lengde"] = f"{pile_est['anbefalt_pellengde_m']:.1f} m"
        st.session_state["rapport_peler_total_lm"] = f"{total_lm:.0f} lm"
        st.session_state["rapport_peler_kostnad_staal"] = f"{pile_est['kostnad_stålpel_kr'] * n_piles:,.0f} kr".replace(",", " ")
        st.session_state["rapport_peler_kostnad_betong"] = f"{pile_est['kostnad_betongpel_kr'] * n_piles:,.0f} kr".replace(",", " ")

        st.subheader(f"Estimat: {n_piles} peler à {pile_est['anbefalt_pellengde_m']} m = {total_lm:.0f} lm")
        pe1, pe2, pe3, pe4 = st.columns(4)
        with pe1: metric_card("Pellengde per pel", f"{pile_est['anbefalt_pellengde_m']:.1f} m")
        with pe2: metric_card("Total lm peler", f"{total_lm:.0f} lm")
        with pe3: metric_card("Kostnad stålpel", f"{pile_est['kostnad_stålpel_kr'] * n_piles:,.0f} kr".replace(",", " "))
        with pe4: metric_card("Kostnad betongpel", f"{pile_est['kostnad_betongpel_kr'] * n_piles:,.0f} kr".replace(",", " "))

        pile_df = pd.DataFrame([{
            "Type": "Stålpel", "Antall": n_piles, "Lengde per pel [m]": pile_est["anbefalt_pellengde_m"],
            "Total [lm]": total_lm,
            "Kostnad [kr]": pile_est["kostnad_stålpel_kr"] * n_piles,
            "CO₂ [kgCO2e]": pile_est["co2_stålpel_kgCO2e"] * n_piles,
        }, {
            "Type": "Betongpel", "Antall": n_piles, "Lengde per pel [m]": pile_est["anbefalt_pellengde_m"],
            "Total [lm]": total_lm,
            "Kostnad [kr]": pile_est["kostnad_betongpel_kr"] * n_piles,
            "CO₂ [kgCO2e]": pile_est["co2_betongpel_kgCO2e"] * n_piles,
        }])
        st.dataframe(pile_df, use_container_width=True, hide_index=True)
        st.download_button("Last ned fundamenteringsdata CSV", pile_df.to_csv(index=False).encode("utf-8-sig"), file_name="fundamentering.csv", mime="text/csv")

    # -----------------------------------------------------------------------
    # TAB 5 – CO₂-REGNSKAP GRUNN
    # -----------------------------------------------------------------------
    with ground_tab5:
        st.subheader("CO₂-regnskap for grunnarbeider")
        st.caption("Beregner klimagassutslipp fra utgraving, transport, fyllmasser, fundamenter og peler.")

        co2_c1, co2_c2 = st.columns(2)
        with co2_c1:
            co2_soil = st.selectbox("Dominerende jordart (for CO₂-faktor)", list(SOIL_DATABASE.keys()), key="co2_soil")
            co2_cut = st.number_input("Estimert utgraving [m³]", min_value=0.0, value=500.0, step=50.0)
            co2_fill = st.number_input("Estimert oppfylling [m³]", min_value=0.0, value=200.0, step=50.0)
        with co2_c2:
            co2_area = st.number_input("Fundamentareal [m²]", min_value=0.0, value=200.0, step=10.0)
            co2_pile_lm = st.number_input("Peler – totalt løpemeter [lm]", min_value=0.0, value=0.0, step=10.0)
            co2_found_type = st.selectbox("Fundamenteringstype (CO₂)", list(FOUNDATION_DATABASE.keys()),
                                           format_func=lambda k: FOUNDATION_DATABASE[k]["label"], key="co2_found_type")

        co2_df = calculate_ground_co2(co2_cut, co2_fill, co2_area, co2_soil, co2_pile_lm, co2_found_type)
        total_co2 = co2_df["CO₂ [kgCO2e]"].sum()
        # Lagre til rapport
        st.session_state["rapport_grunn_co2"] = f"{total_co2:,.0f} kgCO₂e".replace(",", " ")
        st.session_state["rapport_grunn_co2_m2"] = f"{total_co2 / max(co2_area, 1):,.1f} kgCO₂e/m²".replace(",", " ")
        st.session_state["rapport_grunn_cut"] = f"{co2_cut:,.1f} m³".replace(",", " ")
        st.session_state["rapport_grunn_fill"] = f"{co2_fill:,.1f} m³".replace(",", " ")
        st.session_state["rapport_grunn_soil"] = co2_soil
        st.session_state["rapport_grunn_foundation"] = co2_found_type
        st.session_state["rapport_grunn_co2_df"] = co2_df.copy()

        tc1, tc2, tc3 = st.columns(3)
        with tc1: metric_card("Total CO₂", f"{total_co2:,.0f} kgCO₂e".replace(",", " "))
        with tc2: metric_card("CO₂ per m²", f"{total_co2 / max(co2_area, 1):,.1f} kgCO₭e/m²".replace(",", " "))
        with tc3: metric_card("CO₂ per m³ utgraving", f"{total_co2 / max(co2_cut, 1):,.1f} kgCO₂e/m³".replace(",", " "))

        st.dataframe(co2_df, use_container_width=True, hide_index=True)

        fig_co2, ax_co2 = plt.subplots(figsize=(7, 4))
        co2_pos = co2_df[co2_df["CO₂ [kgCO2e]"] > 0]
        ax_co2.barh(co2_pos["Post"], co2_pos["CO₂ [kgCO2e]"], color="#1f4e79")
        ax_co2.set_xlabel("CO₂ [kgCO2e]")
        ax_co2.set_title("CO₂-fordeling grunnarbeider")
        ax_co2.invert_yaxis()
        plt.tight_layout()
        st.pyplot(fig_co2)
        plt.close(fig_co2)

        st.download_button("Last ned CO₂-regnskap CSV", co2_df.to_csv(index=False).encode("utf-8-sig"), file_name="grunn_co2.csv", mime="text/csv")


elif valg == "Materialbytte":
    if st.button("← Hjem", key="hjem_mb"): _nav_to("Hjem")
    st.header("🔁 Materialbytte")
    if data.empty:
        st.info("Materialbytte krever mengde- eller IFC-data i datasettet.")
    else:
        col1, col2, col3 = st.columns(3)
        with col1:
            available_types = sorted(data["Type"].dropna().astype(str).unique().tolist())
            selected_swap_type = st.selectbox("Elementtype som skal byttes", available_types)
        with col2:
            available_materials = sorted(data.loc[data["Type"] == selected_swap_type, "materiale"].dropna().astype(str).unique().tolist())
            from_material = st.selectbox("Nåværende materiale", available_materials)
        with col3:
            target_key = st.selectbox("Nytt system / materiale", get_swap_target_options(selected_swap_type), format_func=format_swap_target_option)

        defaults = get_swap_target_defaults(target_key)
        area_based_target = is_area_based_swap_target(target_key)
        new_profile_text = ""
        if not area_based_target:
            profile_material = target_key if target_key in PROFILE_LIBRARY else classify_material(target_key)
            options = PROFILE_LIBRARY.get(profile_material, [])
            new_profile_text = st.selectbox("Nytt tverrsnitt", options if options else ["115x360"])

        swap_df = calculate_material_swap(data, selected_swap_type, from_material, target_key, new_profile_text)
        if swap_df.empty:
            st.warning("Ingen elementer samsvarer med valgt elementtype og materiale.")
        else:
            m1, m2, m3, m4 = st.columns(4)
            with m1: metric_card("Antall elementer", f"{len(swap_df):,}".replace(",", " "))
            with m2: metric_card("Gammel kostnad", f"{swap_df['Gammel kostnad [kr]'].sum():,.0f} kr".replace(",", " "))
            with m3: metric_card("Ny kostnad", f"{swap_df['Ny kostnad [kr]'].sum():,.0f} kr".replace(",", " "))
            _kost_delta = swap_df['Kostnadsendring [kr]'].sum()
            with m4: metric_card("Kostnadsendring", f"{_kost_delta:,.0f} kr".replace(",", " "))

            _co2_old = swap_df["Gammel CO₂ [kgCO2e]"].sum() if "Gammel CO₂ [kgCO2e]" in swap_df.columns else None
            _co2_new = swap_df["Ny CO₂ [kgCO2e]"].sum() if "Ny CO₂ [kgCO2e]" in swap_df.columns else None
            _co2_delta = (_co2_new - _co2_old) if (_co2_old is not None and _co2_new is not None) else None

            _kost_color = "#d4edda" if _kost_delta <= 0 else "#f8d7da"
            _kost_icon = "✅" if _kost_delta <= 0 else "⚠️"
            _banner_parts = [f"{_kost_icon} Kostnadsendring: **{_kost_delta:+,.0f} kr**".replace(",", " ")]
            if _co2_delta is not None:
                _co2_color = "#d4edda" if _co2_delta <= 0 else "#f8d7da"
                _co2_icon = "🌿" if _co2_delta <= 0 else "⚠️"
                _banner_parts.append(f"{_co2_icon} CO₂-endring: **{_co2_delta:+,.0f} kgCO₂e**".replace(",", " "))
                _banner_bg = "#d4edda" if (_kost_delta <= 0 and _co2_delta <= 0) else "#f8d7da"
            else:
                _banner_bg = _kost_color
            st.markdown(f"<div style='background:{_banner_bg};padding:0.6rem 1rem;border-radius:6px;margin-bottom:0.5rem;'>{' &nbsp;|&nbsp; '.join(_banner_parts)}</div>", unsafe_allow_html=True)

            st.dataframe(swap_df, use_container_width=True, hide_index=True)

            if uploaded_ifc is not None and "IFC GlobalId" in swap_df.columns:
                preview_ids = tuple(sorted(set(swap_df["IFC GlobalId"].dropna().astype(str).tolist())))
                try:
                    preview_meshes = extract_ifc_meshes_filtered(ifc_bytes, visible_ids_tuple=None, max_elements=500)
                    fig_preview = build_ifc_3d_figure(preview_meshes, preview_ids=preview_ids, show_only_preview=False, preview_material=defaults["label"])
                    st.plotly_chart(fig_preview, use_container_width=True)
                except Exception as e:
                    st.warning(f"Kunne ikke generere 3D-forhåndsvisning: {e}")

                if st.button("Generer ny IFC-fil"):
                    try:
                        new_ifc_bytes, ifc_change_log = export_ifc_material_swap(ifc_bytes, data, selected_swap_type, from_material, target_key, new_profile_text)
                        if new_ifc_bytes is None or ifc_change_log.empty:
                            st.warning("Ingen elementer ble oppdatert i IFC-filen.")
                        else:
                            out_name = Path(uploaded_ifc.name).stem + f"_materialbytte_{re.sub(r'[^A-Za-z0-9_-]+', '_', str(target_key))}.ifc"
                            st.download_button("Last ned ny IFC-fil", data=new_ifc_bytes, file_name=out_name, mime="application/octet-stream")
                            st.dataframe(ifc_change_log, use_container_width=True, hide_index=True)
                    except Exception as e:
                        st.error(f"Kunne ikke generere IFC-fil: {e}")

elif valg == "3D-modell":
    if st.button("← Hjem", key="hjem_3d"): _nav_to("Hjem")
    st.header("🧊 3D-modellvisning")
    if uploaded_ifc is None:
        st.info("3D-modellvisning er tilgjengelig når en IFC-fil er lastet opp.")
    else:
        visning = st.radio("Visning", ["Kun filtrerte elementer", "Alle elementer"], horizontal=True)
        max_elements_3d = st.slider("Maks antall elementer i 3D-visning", 100, 5000, 300, 100)
        visible_ids = tuple(sorted(set(filtered["IFC GlobalId"].dropna().astype(str).tolist()))) if visning == "Kun filtrerte elementer" and "IFC GlobalId" in filtered.columns else None
        should_load_3d = st.button("Last 3D-modell") if lazy_load_3d else True
        if should_load_3d:
            try:
                with st.spinner("Laster 3D-modell..."):
                    meshes = extract_ifc_meshes_filtered(ifc_bytes, visible_ids_tuple=visible_ids, max_elements=max_elements_3d)
                if meshes:
                    fig3d = build_ifc_3d_figure(meshes)
                    st.plotly_chart(fig3d, use_container_width=True)
                else:
                    st.warning("Ingen 3D-geometri ble funnet for valgt utvalg.")
            except Exception as e:
                st.error(f"Kunne ikke generere 3D-visning: {e}")
        else:
            st.info("Klikk på knappen for å laste 3D-modellen.")

elif valg == "Bygggenerator":
    if st.button("← Hjem", key="hjem_bg"): _nav_to("Hjem")
    st.header("🏗️ Bygggenerator")
    st.caption("Generer et råbygg direkte i appen. Modulen bruker samme logikk som Excel-arket: grid/fag, etasjer, geometri, dekker og materialkvaliteter.")

    with st.expander("Hva gjør denne modulen?", expanded=True):
        st.markdown("""
        - Velg **geometri**: rektangel, L-form/tilbygg og eventuelt åpning i dekke.  
        - Velg **etasjer og høyder**.  
        - Velg **materiale og kvalitet** for bjelker, søyler og dekker.  
        - Appen genererer ramme, dekker, mengder, vekt, kostnad og CO₂ som et tidligfaseforslag.
        """)

    st.subheader("1. Geometri og etasjer")
    g1, g2, g3, g4 = st.columns(4)
    bg_params = {}
    with g1:
        planvalg = st.selectbox("Geometri", ["Rektangel", "L-form / tilbygg", "Rektangel med åpning", "L-form med åpning"])
        bg_params["fag_x_r1"] = st.number_input("Fag X", min_value=1, max_value=20, value=4, step=1, key="bg_fx")
        bg_params["fag_y_r1"] = st.number_input("Fag Y", min_value=1, max_value=10, value=2, step=1, key="bg_fy")
    with g2:
        bg_params["faglengde_x_mm"] = st.number_input("Faglengde X [mm]", min_value=1000, max_value=20000, value=8000, step=500, key="bg_dx")
        bg_params["faglengde_y_mm"] = st.number_input("Faglengde Y [mm]", min_value=1000, max_value=20000, value=12000, step=500, key="bg_dy")
    with g3:
        bg_params["antall_etasjer"] = st.number_input("Antall etasjer", min_value=1, max_value=10, value=3, step=1, key="bg_levels")
        bg_params["dekker_i_modell"] = bg_params["antall_etasjer"]
        bg_params["etasjehoyde_mm"] = st.number_input("Etasjehøyde [mm]", min_value=2200, max_value=6000, value=3000, step=100, key="bg_floor_height")
    with g4:
        bg_params["dekke_tykkelse_mm"] = st.number_input("Dekketøykkelse [mm]", min_value=100, max_value=600, value=300, step=10, key="bg_slab_thk")
        bg_params["dekker_aktiv"] = "JA" if st.toggle("Generer dekker", value=True, key="bg_slabs_on") else "NEI"

    use_r2 = planvalg in ["L-form / tilbygg", "L-form med åpning"]
    use_opening = planvalg in ["Rektangel med åpning", "L-form med åpning"]
    bg_params["rektangel2_aktiv"] = "JA" if use_r2 else "NEI"
    if use_r2:
        r1, r2, r3, r4 = st.columns(4)
        with r1:
            bg_params["fag_x_r2"] = st.number_input("Tilbygg fag X", min_value=1, max_value=20, value=2, step=1, key="bg_r2_fx")
        with r2:
            bg_params["fag_y_r2"] = st.number_input("Tilbygg fag Y", min_value=1, max_value=10, value=2, step=1, key="bg_r2_fy")
        with r3:
            bg_params["r2_offset_x_fag"] = st.number_input("Tilbygg offset X [fag]", min_value=0, max_value=20, value=2, step=1, key="bg_r2_ox")
        with r4:
            bg_params["r2_offset_y_fag"] = st.number_input("Tilbygg offset Y [fag]", min_value=0, max_value=10, value=0, step=1, key="bg_r2_oy")
    else:
        bg_params.update({"fag_x_r2": 0, "fag_y_r2": 0, "r2_offset_x_fag": 0, "r2_offset_y_fag": 0})

    if use_opening:
        o1, o2, o3, o4 = st.columns(4)
        with o1:
            bg_params["opening_width_fag"] = st.number_input("Åpning bredde [fag]", min_value=1, max_value=10, value=1, step=1, key="bg_ow")
        with o2:
            bg_params["opening_height_fag"] = st.number_input("Åpning høyde [fag]", min_value=1, max_value=10, value=1, step=1, key="bg_oh")
        with o3:
            bg_params["opening_offset_x_fag"] = st.number_input("Åpning offset X [fag]", min_value=0, max_value=20, value=1, step=1, key="bg_oox")
        with o4:
            bg_params["opening_offset_y_fag"] = st.number_input("Åpning offset Y [fag]", min_value=0, max_value=10, value=1, step=1, key="bg_ooy")
    else:
        bg_params.update({"opening_width_fag": 0, "opening_height_fag": 0, "opening_offset_x_fag": 0, "opening_offset_y_fag": 0})

    st.subheader("2. Materiale og kvalitet")
    m1, m2, m3 = st.columns(3)
    with m1:
        bg_params["bjelkemateriale"] = st.selectbox("Bjelkemateriale", ["Stål", "Limtre"], key="bg_beam_mat")
        bg_params["bjelkekvalitet"] = st.selectbox("Bjelkekvalitet", quality_options(bg_params["bjelkemateriale"]), key="bg_beam_qual")
        beam_profiles = PROFILE_LIBRARY.get(bg_params["bjelkemateriale"], PROFILE_LIBRARY["Stål"])
        bg_params["bjelkeprofil"] = st.selectbox("Bjelkeprofil", beam_profiles, key="bg_beam_prof")
    with m2:
        bg_params["søylemateriale"] = st.selectbox("Søylemateriale", ["Stål", "Limtre", "Betong"], key="bg_col_mat")
        bg_params["søylekvalitet"] = st.selectbox("Søylekvalitet", quality_options(bg_params["søylemateriale"]), key="bg_col_qual")
        col_profiles = PROFILE_LIBRARY.get(bg_params["søylemateriale"], PROFILE_LIBRARY["Betong"])
        bg_params["søyleprofil"] = st.selectbox("Søyleprofil", col_profiles, key="bg_col_prof")
    with m3:
        bg_params["dekke_materialtype"] = st.selectbox("Dekkemateriale", ["Betong", "Massivtre"], key="bg_deck_mat")
        bg_params["dekke_kvalitet"] = st.selectbox("Dekkekvalitet", quality_options(bg_params["dekke_materialtype"]), key="bg_deck_qual")
        bg_params["skalltype"] = st.selectbox("Skalltype", ["Platt skall", "Hulldekke-prinsipp", "Massivtredekke"], key="bg_shell")
        bg_params["dekke_materiale"] = make_material_profile_label(bg_params["dekke_materialtype"], bg_params["dekke_kvalitet"], f"t={bg_params['dekke_tykkelse_mm']:.0f} mm")

    geom = generate_plan_geometry(bg_params)
    frame_df = generate_frame_export_parametric(bg_params)
    slab_df = generate_slab_export(bg_params) if bg_params["dekker_aktiv"] == "JA" else pd.DataFrame()
    # Lagre i session state slik at Rapport-siden kan bruke dem til IFC-eksport og visning
    st.session_state["bg_params_last"] = bg_params
    st.session_state["bg_frame_df_last"] = frame_df
    st.session_state["bg_slab_df_last"] = slab_df
    qty_df = frame_to_quantity_dataset(frame_df, slab_df, bg_params)
    qa_df = run_project_qa(bg_params, frame_df, slab_df if not slab_df.empty else pd.DataFrame(columns=["DeckID"]))

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1: metric_card("Planform", geom["planformkode"])
    with k2: metric_card("Aktivt areal", f"{geom['active_area_m2']:,.1f} m²".replace(",", " "))
    with k3: metric_card("Elementer", f"{len(qty_df):,}".replace(",", " "))
    with k4: metric_card("Kostnad", f"{qty_df['Kostnad [kr]'].sum():,.0f} kr".replace(",", " "))
    with k5: metric_card("CO₂", f"{qty_df['CO2 [kgCO2e]'].sum():,.0f} kgCO₂e".replace(",", " "))

    p_left, p_right = st.columns([1, 1.15])
    with p_left:
        st.subheader("3. 2D-plan")
        st.pyplot(plot_plan_geometry(geom))
    with p_right:
        st.subheader("4. 3D-prinsippmodell")
        st.plotly_chart(plot_frame_3d(frame_df, slab_df), use_container_width=True)

    # Lagre nøkkeltall til rapport-siden
    st.session_state["rapport_bg_elementer"] = str(len(qty_df))
    st.session_state["rapport_bg_kostnad"] = f"{pd.to_numeric(qty_df['Kostnad [kr]'], errors='coerce').fillna(0).sum():,.0f} kr".replace(",", " ")
    st.session_state["rapport_bg_co2"] = f"{pd.to_numeric(qty_df['CO2 [kgCO2e]'], errors='coerce').fillna(0).sum():,.0f} kgCO₂e".replace(",", " ")
    st.session_state["rapport_bg_vekt"] = f"{pd.to_numeric(qty_df['Vekt [kg]'], errors='coerce').fillna(0).sum():,.0f} kg".replace(",", " ")
    st.session_state["rapport_bg_volum"] = f"{pd.to_numeric(qty_df['Volum [m3]'], errors='coerce').fillna(0).sum():,.2f} m³".replace(",", " ")
    st.session_state["rapport_bg_etasjer"] = str(bg_params.get("antall_etasjer", "–"))
    st.session_state["rapport_bg_bredde"] = str(bg_params.get("bredde_mm", "–"))
    st.session_state["rapport_bg_lengde"] = str(bg_params.get("lengde_mm", "–"))
    st.session_state["rapport_bg_hoyde"] = str(bg_params.get("etasjehoyde_mm", "–"))
    st.session_state["rapport_bg_qty_df"] = qty_df.copy()

    st.subheader("5. Mengder, vekt, kostnad og CO₂")
    st.dataframe(qty_df, use_container_width=True, hide_index=True, height=430)

    sum_df = qty_df.groupby(["Type", "materiale", "Materialkvalitet"], dropna=False).agg(
        Antall=("Segment", "count"),
        Lengde_m=("Lengde [m]", "sum"),
        Areal_m2=("Areal [m2]", "sum"),
        Volum_m3=("Volum [m3]", "sum"),
        Vekt_kg=("Vekt [kg]", "sum"),
        Kostnad_kr=("Kostnad [kr]", "sum"),
        CO2_kg=("CO2 [kgCO2e]", "sum"),
    ).reset_index()
    st.subheader("6. Oppsummering")
    st.dataframe(sum_df, use_container_width=True, hide_index=True)

    with st.expander("QA-kontroll"):
        st.dataframe(qa_df, use_container_width=True, hide_index=True)

    dl1, dl2, dl3, dl4 = st.columns(4)
    with dl1:
        st.download_button("Last ned mengder CSV", qty_df.to_csv(index=False).encode("utf-8-sig"), file_name="bygggenerator_mengder.csv", mime="text/csv")
    with dl2:
        st.download_button("Last ned ramme CSV", frame_df.to_csv(index=False).encode("utf-8-sig"), file_name="bygggenerator_ramme.csv", mime="text/csv")
    with dl3:
        if not slab_df.empty:
            st.download_button("Last ned dekker CSV", slab_df.to_csv(index=False).encode("utf-8-sig"), file_name="bygggenerator_dekker.csv", mime="text/csv")
    with dl4:
        if ifcopenshell is None:
            st.warning("IFC-eksport krever at `ifcopenshell` ligger i requirements.txt.")
        else:
            try:
                ifc_out = generate_building_ifc_bytes(frame_df, slab_df, bg_params)
                st.download_button("Last ned bygg IFC", data=ifc_out, file_name="bygggenerator_generert_bygg.ifc", mime="application/octet-stream")
            except Exception as e:
                st.error(f"Kunne ikke generere IFC: {e}")

elif valg == "Rapport":
    import matplotlib.pyplot as plt
    from ground_module import SOIL_DATABASE, FOUNDATION_DATABASE, recommend_foundation, estimate_pile_length

    ss = st.session_state

    # Beregn defaultverdier første gang rapporten åpnes
    if "rapport_bg_elementer" not in ss:
        try:
            _def_params = {
                "fag_x_r1": 4, "fag_y_r1": 2,
                "faglengde_x_mm": 8000, "faglengde_y_mm": 12000,
                "antall_etasjer": 3, "dekker_i_modell": 3,
                "etasjehoyde_mm": 3000, "dekke_tykkelse_mm": 300,
                "dekker_aktiv": "JA", "rektangel2_aktiv": "NEI",
                "opening_aktiv": "NEI",
                "bjelkemateriale": "Stål", "bjelkekvalitet": "S355",
                "bjelkeprofil": "KFHUP 200x200x12.5",
                "søylemateriale": "Stål", "søylekvalitet": "S355",
                "søyleprofil": "KFHUP 200x200x12.5",
                "dekke_materialtype": "Betong", "dekke_kvalitet": "B35",
                "skalltype": "Platt skall",
                "dekke_materiale": "Betong B35, t=300 mm",
            }
            _frame = generate_frame_export_parametric(_def_params)
            _slab  = generate_slab_export(_def_params)
            _qty   = frame_to_quantity_dataset(_frame, _slab, _def_params)
            ss["bg_params_last"]       = _def_params
            ss["bg_frame_df_last"]     = _frame
            ss["bg_slab_df_last"]      = _slab
            ss["rapport_bg_elementer"] = str(len(_qty))
            ss["rapport_bg_kostnad"]   = f"{_qty['Kostnad [kr]'].sum():,.0f} kr".replace(",", " ")
            ss["rapport_bg_co2"]       = f"{_qty['CO2 [kgCO2e]'].sum():,.0f} kgCO₂e".replace(",", " ")
            ss["rapport_bg_vekt"]      = f"{_qty['Vekt [kg]'].sum():,.0f} kg".replace(",", " ")
            ss["rapport_bg_volum"]     = f"{_qty['Volum [m3]'].sum():,.2f} m³".replace(",", " ")
            ss["rapport_bg_etasjer"]   = str(_def_params["antall_etasjer"])
            ss["rapport_bg_bredde"]    = f"{_def_params['fag_x_r1'] * _def_params['faglengde_x_mm']:,}".replace(",", " ")
            ss["rapport_bg_lengde"]    = f"{_def_params['fag_y_r1'] * _def_params['faglengde_y_mm']:,}".replace(",", " ")
            ss["rapport_bg_hoyde"]     = str(_def_params["etasjehoyde_mm"])
            ss["rapport_bg_qty_df"]    = _qty.copy()
        except Exception as _e:
            st.warning(f"Kunne ikke beregne standardverdier for Bygggenerator: {_e}")

    if "rapport_jordart" not in ss:
        try:
            _def_soil    = "Leire"
            _def_area    = 200.0
            _def_load    = 2000.0
            _def_rec     = recommend_foundation(_def_soil, _def_area, _def_load)
            _def_pile    = estimate_pile_length(8.0, _def_soil)
            _def_n_piles = max(4, int(_def_area / 16))
            ss["rapport_jordart"]              = SOIL_DATABASE[_def_soil]["label"]
            ss["rapport_bæreevne"]             = f"{_def_rec['bæreevne_kPa']:.0f} kPa"
            ss["rapport_fundament_anbefalt"]   = _def_rec["label"]
            ss["rapport_bygningsareal"]        = f"{_def_area:,.0f} m²".replace(",", " ")
            ss["rapport_total_last"]           = f"{_def_load:,.0f} kN".replace(",", " ")
            ss["rapport_peler_antall"]         = str(_def_n_piles)
            ss["rapport_peler_lengde"]         = f"{_def_pile['anbefalt_pellengde_m']:.1f} m"
            ss["rapport_peler_total_lm"]       = f"{_def_pile['anbefalt_pellengde_m'] * _def_n_piles:.0f} lm"
            ss["rapport_peler_kostnad_staal"]  = f"{_def_pile['kostnad_stålpel_kr'] * _def_n_piles:,.0f} kr".replace(",", " ")
            ss["rapport_peler_kostnad_betong"] = f"{_def_pile['kostnad_betongpel_kr'] * _def_n_piles:,.0f} kr".replace(",", " ")
        except Exception as _e:
            st.warning(f"Kunne ikke beregne standardverdier for Grunn: {_e}")

    # -----------------------------------------------------------------------
    # DEMO-DATA – vises når ingen ekte data finnes
    # -----------------------------------------------------------------------
    DEMO_BG = {
        "rapport_bg_elementer": "42",
        "rapport_bg_kostnad":   "4 280 000 kr",
        "rapport_bg_co2":       "185 400 kgCO₂e",
        "rapport_bg_vekt":      "312 000 kg",
        "rapport_bg_volum":     "128.50 m³",
        "rapport_bg_etasjer":   "4",
        "rapport_bg_bredde":    "12 000",
        "rapport_bg_lengde":    "18 000",
        "rapport_bg_hoyde":     "3 000",
        "rapport_bg_qty_df":    pd.DataFrame({
            "Type": ["Søyle", "Bjelke", "Dekke", "Søyle", "Bjelke"],
            "materiale": ["Stål", "Stål", "Betong", "Limtre", "Limtre"],
            "Antall": [12, 18, 4, 8, 10],
            "Lengde_m": [12.0, 54.0, 0.0, 9.6, 30.0],
            "Volum_m3": [0.42, 1.08, 86.4, 0.55, 0.90],
            "Vekt_kg": [3294, 8478, 207360, 253, 414],
            "Kostnad_kr": [154818, 398466, 3456000, 15400, 25200],
            "CO2_kgCO2e": [2405, 6189, 176220, 55, 99],
        }),
    }
    DEMO_GRUNN = {
        "rapport_tomteareal":        "1 250.0 m²",
        "rapport_prosjektkote":      "12.50 m",
        "rapport_utgraving":         "875.0 m³",
        "rapport_oppfylling":        "320.0 m³",
        "rapport_antall_punkt":      "156",
        "rapport_jordart":           "Leire",
        "rapport_bæreevne":          "40 kPa",
        "rapport_fundament_anbefalt":"Pelefundament – stål H-pile",
        "rapport_bygningsareal":     "200 m²",
        "rapport_total_last":        "2 000 kN",
        "rapport_peler_antall":      "16",
        "rapport_peler_lengde":      "10.0 m",
        "rapport_peler_total_lm":    "160 lm",
        "rapport_peler_kostnad_staal":"192 000 kr",
        "rapport_peler_kostnad_betong":"156 800 kr",
        "rapport_grunn_co2":         "24 650 kgCO₂e",
        "rapport_grunn_co2_m2":      "123.3 kgCO₂e/m²",
        "rapport_grunn_soil":        "Leire",
        "rapport_grunn_foundation":  "Pelefundament_staal",
        "rapport_grunn_co2_df": pd.DataFrame({
            "Post": ["Utgraving", "Bortkjøring", "Importert fyllmasse", "Forsterkningslag", "Peler (stål)"],
            "Enhet": ["m3","m3","m3","m3","lm"],
            "Mengde": [875, 875, 320, 60, 160],
            "CO₂-faktor": [2.5, 5.5, 8.0, 15.0, 28.0],
            "CO₂ [kgCO2e]": [2187.5, 4812.5, 2560.0, 900.0, 4480.0],
        }),
    }

    has_bg    = "rapport_bg_elementer" in ss
    has_grunn = "rapport_jordart" in ss

    # -----------------------------------------------------------------------
    # TOPPLINJE
    # -----------------------------------------------------------------------
    if st.button("← Hjem", key="hjem_rapport"): _nav_to("Hjem")
    st.header("📝 Prosjektrapport")
    head_col, demo_col = st.columns([3, 1])
    with head_col:
        st.caption("Rapporten viser gjeldende verdier fra Bygggenerator og Grunn og oppdateres automatisk ved endringer.")
    with demo_col:
        show_demo = st.toggle("Vis demorapport", value=False, key="rapport_demo_toggle")

    if show_demo:
        data_bg    = DEMO_BG
        data_grunn = DEMO_GRUNN
        st.info("📋 Demorapport – eksempeldata. Slå av for å se prosjektets egne verdier.")
    else:
        data_bg    = ss
        data_grunn = ss

    st.divider()

    # -----------------------------------------------------------------------
    # INNHOLDSVELGER
    # -----------------------------------------------------------------------
    with st.expander("⚙️ Velg innhold i rapporten", expanded=not (has_bg or has_grunn or show_demo)):
        st.markdown("**Slå av/på seksjoner:**")
        ic1, ic2, ic3, ic4 = st.columns(4)
        with ic1:
            show_bg_hoved   = st.checkbox("Bygggenerator – nøkkeltall",  value=True,  key="rpt_bg_hoved")
            show_bg_tabell  = st.checkbox("Bygggenerator – mengdetabell", value=True,  key="rpt_bg_tabell")
        with ic2:
            show_terreng    = st.checkbox("Grunn – terreng/stikningsdata", value=True, key="rpt_terreng")
            show_geoteknikk = st.checkbox("Grunn – geoteknikk/fundament", value=True,  key="rpt_geoteknikk")
        with ic3:
            show_peler      = st.checkbox("Grunn – peler",                value=True,  key="rpt_peler")
            show_co2_grunn  = st.checkbox("Grunn – CO₂-regnskap",         value=True,  key="rpt_co2_grunn")
        with ic4:
            show_co2_sum    = st.checkbox("Sammenstilt CO₂-diagram",      value=True,  key="rpt_co2_sum")
            show_eksport    = st.checkbox("Eksport og IFC",               value=True,  key="rpt_eksport")

    # -----------------------------------------------------------------------
    # BYGGGENERATOR – NØKKELTALL
    # -----------------------------------------------------------------------
    if show_bg_hoved:
        st.subheader("🏗️ Bygggenerator")
        b1, b2, b3, b4, b5 = st.columns(5)
        with b1: metric_card("Elementer",   data_bg.get("rapport_bg_elementer", "–"))
        with b2: metric_card("Kostnad",     data_bg.get("rapport_bg_kostnad",   "–"))
        with b3: metric_card("CO₂",         data_bg.get("rapport_bg_co2",       "–"))
        with b4: metric_card("Vekt",        data_bg.get("rapport_bg_vekt",      "–"))
        with b5: metric_card("Volum",       data_bg.get("rapport_bg_volum",     "–"))

        p1, p2, p3, p4 = st.columns(4)
        with p1: metric_card("Etasjer",         data_bg.get("rapport_bg_etasjer", "–"))
        with p2: metric_card("Bredde [mm]",     data_bg.get("rapport_bg_bredde",  "–"))
        with p3: metric_card("Lengde [mm]",     data_bg.get("rapport_bg_lengde",  "–"))
        with p4: metric_card("Etasjehøyde [mm]",data_bg.get("rapport_bg_hoyde",  "–"))

    if show_bg_tabell:
        with st.expander("📋 Mengdetabell – Bygggenerator", expanded=show_bg_hoved):
            qty_r = data_bg.get("rapport_bg_qty_df", pd.DataFrame())
            if not qty_r.empty:
                agg_cols = {c: "sum" for c in ["Antall","Lengde_m","Volum_m3","Vekt_kg","Kostnad_kr","CO2_kgCO2e"] if c in qty_r.columns}
                if agg_cols:
                    group_cols = [c for c in ["Type","materiale"] if c in qty_r.columns]
                    if group_cols:
                        sum_r = qty_r.groupby(group_cols, dropna=False).agg(agg_cols).reset_index()
                        st.dataframe(sum_r, use_container_width=True, hide_index=True)
                    else:
                        st.dataframe(qty_r, use_container_width=True, hide_index=True)
                else:
                    st.dataframe(qty_r, use_container_width=True, hide_index=True)
            else:
                st.info("Ingen mengdedata tilgjengelig.")

    if show_bg_hoved or show_bg_tabell:
        st.divider()

    # -----------------------------------------------------------------------
    # GRUNNFORHOLD
    # -----------------------------------------------------------------------
    any_grunn_shown = show_terreng or show_geoteknikk or show_peler or show_co2_grunn
    if any_grunn_shown:
        st.subheader("🌍 Grunnforhold")

    if show_terreng:
        st.markdown("**Terreng og stikningsdata**")
        t1, t2, t3, t4, t5 = st.columns(5)
        with t1: metric_card("Tomteareal",   data_grunn.get("rapport_tomteareal",   "–"))
        with t2: metric_card("Prosjektkote", data_grunn.get("rapport_prosjektkote", "–"))
        with t3: metric_card("Utgraving",    data_grunn.get("rapport_utgraving",    "–"))
        with t4: metric_card("Oppfylling",   data_grunn.get("rapport_oppfylling",   "–"))
        with t5: metric_card("Stikningspunkt", data_grunn.get("rapport_antall_punkt","–"))

    if show_geoteknikk:
        st.markdown("**Geoteknikk og fundamentering**")
        g1, g2, g3, g4, g5 = st.columns(5)
        with g1: metric_card("Jordart",            data_grunn.get("rapport_jordart",           "–"))
        with g2: metric_card("Bæreevne",           data_grunn.get("rapport_bæreevne",          "–"))
        with g3: metric_card("Anbefalt fundament", data_grunn.get("rapport_fundament_anbefalt","–"))
        with g4: metric_card("Bygningsareal",      data_grunn.get("rapport_bygningsareal",     "–"))
        with g5: metric_card("Total last",         data_grunn.get("rapport_total_last",        "–"))

    if show_peler and data_grunn.get("rapport_peler_antall", "0") not in ("0", "–", ""):
        st.markdown("**Peler**")
        pe1, pe2, pe3, pe4, pe5 = st.columns(5)
        with pe1: metric_card("Antall peler",     data_grunn.get("rapport_peler_antall",          "–"))
        with pe2: metric_card("Pellengde",         data_grunn.get("rapport_peler_lengde",          "–"))
        with pe3: metric_card("Totalt løpemeter",  data_grunn.get("rapport_peler_total_lm",        "–"))
        with pe4: metric_card("Kostnad stålpel",   data_grunn.get("rapport_peler_kostnad_staal",   "–"))
        with pe5: metric_card("Kostnad betongpel", data_grunn.get("rapport_peler_kostnad_betong",  "–"))

    if show_co2_grunn:
        st.markdown("**CO₂ – grunnarbeider**")
        c1, c2, c3, c4 = st.columns(4)
        with c1: metric_card("Total CO₂ grunn",  data_grunn.get("rapport_grunn_co2",      "–"))
        with c2: metric_card("CO₂ per m²",        data_grunn.get("rapport_grunn_co2_m2",   "–"))
        with c3: metric_card("Jordart (CO₂)",     data_grunn.get("rapport_grunn_soil",     "–"))
        with c4: metric_card("Fundamenttype",      data_grunn.get("rapport_grunn_foundation","–"))

        with st.expander("📋 Detaljert CO₂-regnskap grunn", expanded=False):
            co2_df_r = data_grunn.get("rapport_grunn_co2_df", pd.DataFrame())
            if not co2_df_r.empty:
                st.dataframe(co2_df_r, use_container_width=True, hide_index=True)

    if any_grunn_shown:
        st.divider()

    # -----------------------------------------------------------------------
    # SAMMENSTILT CO₂
    # -----------------------------------------------------------------------
    if show_co2_sum:
        st.subheader("🌿 Sammenstilt CO₂-oversikt")

        def _parse_co2(val_str):
            try:
                return float(str(val_str).replace(" ", "").replace("kgCO₂e", "").replace(",", ".").split("k")[0])
            except Exception:
                return 0.0

        co2_bg_val  = _parse_co2(data_bg.get("rapport_bg_co2",    "0"))
        co2_gr_val  = _parse_co2(data_grunn.get("rapport_grunn_co2", "0"))
        co2_total   = co2_bg_val + co2_gr_val

        cs1, cs2, cs3 = st.columns(3)
        with cs1: metric_card("CO₂ konstruksjon",  f"{co2_bg_val:,.0f} kgCO₂e".replace(",", " "))
        with cs2: metric_card("CO₂ grunnarbeider", f"{co2_gr_val:,.0f} kgCO₂e".replace(",", " "))
        with cs3: metric_card("Total CO₂ prosjekt",f"{co2_total:,.0f} kgCO₂e".replace(",", " "))

        if co2_total > 0:
            fig_s, ax_s = plt.subplots(figsize=(5, 3))
            labels = []; values = []; colors = []
            if co2_bg_val > 0:
                labels.append("Konstruksjon"); values.append(co2_bg_val); colors.append("#1f4e79")
            if co2_gr_val > 0:
                labels.append("Grunnarbeider"); values.append(co2_gr_val); colors.append("#2e7d32")
            bars = ax_s.bar(labels, values, color=colors, edgecolor="#333", linewidth=0.5)
            ax_s.set_ylabel("kgCO₂e")
            ax_s.set_title("CO₂-fordeling")
            for bar, val in zip(bars, values):
                ax_s.text(bar.get_x() + bar.get_width() / 2,
                          bar.get_height() + co2_total * 0.01,
                          f"{val:,.0f}".replace(",", " "),
                          ha="center", va="bottom", fontsize=9)
            plt.tight_layout()
            st.pyplot(fig_s)
            plt.close(fig_s)

        st.divider()

    # -----------------------------------------------------------------------
    # -----------------------------------------------------------------------
    # BYGNINGSPLASSERING PÅ TOMT
    # -----------------------------------------------------------------------
    show_plassering = st.session_state.get("rpt_plassering", True)
    with st.expander("⚙️ Velg innhold i rapporten", expanded=False):
        pass  # already rendered above – this block handled by checkbox

    st.divider()
    st.subheader("📍 Plassering av bygg på tomt")
    st.caption("Flytt bygget manuelt på tomten. Koordinatene brukes ved IFC-eksport.")

    # Hent byggets dimensjoner fra session state
    _bg_p = ss.get("bg_params_last", {})
    _bx = safe_num(_bg_p.get("fag_x_r1", 4)) * safe_num(_bg_p.get("faglengde_x_mm", 8000)) / 1000.0
    _by = safe_num(_bg_p.get("fag_y_r1", 2)) * safe_num(_bg_p.get("faglengde_y_mm", 12000)) / 1000.0
    if _bx <= 0: _bx = 32.0
    if _by <= 0: _by = 24.0

    # Hent tomteareal fra session state (brukes til å sette max-grenser)
    _tomt_str = ss.get("rapport_tomteareal", "1250")
    try:
        _tomt_m2 = float(str(_tomt_str).replace(" ", "").replace("m²","").replace(",",".").split("m")[0])
    except Exception:
        _tomt_m2 = 1250.0
    _tomt_side = max(_tomt_m2 ** 0.5, max(_bx, _by) + 10.0)

    pl1, pl2, pl3, pl4 = st.columns(4)
    with pl1:
        bygg_x = st.number_input(
            "Bygg offset X [m]",
            min_value=0.0,
            max_value=float(_tomt_side - _bx),
            value=float(round((_tomt_side - _bx) / 2.0, 1)),
            step=0.5,
            key="bygg_offset_x",
            help="Avstand fra venstre tomtekant til venstre hjørne av bygget"
        )
    with pl2:
        bygg_y = st.number_input(
            "Bygg offset Y [m]",
            min_value=0.0,
            max_value=float(_tomt_side - _by),
            value=float(round((_tomt_side - _by) / 2.0, 1)),
            step=0.5,
            key="bygg_offset_y",
            help="Avstand fra nedre tomtekant til nedre hjørne av bygget"
        )
    with pl3:
        bygg_rot = st.number_input(
            "Rotasjon [grader]",
            min_value=0.0,
            max_value=360.0,
            value=0.0,
            step=5.0,
            key="bygg_rot",
            help="Roter bygget rundt sitt eget sentrum"
        )
    with pl4:
        tomt_side_input = st.number_input(
            "Tomtestørrelse [m] (kvadrat)",
            min_value=max(_bx, _by) + 2.0,
            max_value=500.0,
            value=float(round(_tomt_side, 0)),
            step=1.0,
            key="tomt_side_input",
        )

    # Verdiene leses direkte fra widget-nøklene i session state (bygg_offset_x, bygg_offset_y osv.)
    # Streamlit lagrer disse automatisk – ingen manuell tilordning nødvendig

    # Vis planvisning med bygget på tomten
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    from matplotlib.patches import FancyArrowPatch
    import math

    fig_pl, ax_pl = plt.subplots(figsize=(7, 7))

    # Tomt
    tomt = plt.Polygon(
        [(0,0),(tomt_side_input,0),(tomt_side_input,tomt_side_input),(0,tomt_side_input)],
        closed=True, fill=True, facecolor="#e8d5a0", edgecolor="#8B7355", linewidth=2, label="Tomt"
    )
    ax_pl.add_patch(tomt)

    # Bygg (rotert rundt sentrum)
    cx = bygg_x + _bx / 2.0
    cy = bygg_y + _by / 2.0
    angle_rad = math.radians(bygg_rot)
    corners_local = [(-_bx/2, -_by/2), (_bx/2, -_by/2), (_bx/2, _by/2), (-_bx/2, _by/2)]
    def rotate(px, py, a):
        return (px*math.cos(a) - py*math.sin(a), px*math.sin(a) + py*math.cos(a))
    corners_world = [(cx + rotate(lx, ly, angle_rad)[0], cy + rotate(lx, ly, angle_rad)[1]) for lx, ly in corners_local]
    bygg_patch = plt.Polygon(corners_world, closed=True, fill=True,
                              facecolor="#1f4e79", edgecolor="#0d2b45",
                              linewidth=2, alpha=0.85, label=f"Bygg ({_bx:.0f}×{_by:.0f} m)")
    ax_pl.add_patch(bygg_patch)

    # Sentrum-markør
    ax_pl.plot(cx, cy, "w+", markersize=10, markeredgewidth=2)

    # Nord-pil
    ax_pl.annotate("N", xy=(tomt_side_input * 0.95, tomt_side_input * 0.92),
                    fontsize=12, fontweight="bold", color="#333", ha="center")
    ax_pl.annotate("", xy=(tomt_side_input * 0.95, tomt_side_input * 0.98),
                    xytext=(tomt_side_input * 0.95, tomt_side_input * 0.88),
                    arrowprops=dict(arrowstyle="->", color="#333", lw=1.5))

    # Kotekanter med mål
    ax_pl.annotate("", xy=(_bx + bygg_x if bygg_rot == 0 else cx + _bx/2, bygg_y - 1.5 if bygg_rot == 0 else cy - _by/2 - 1.5),
                    xytext=(bygg_x if bygg_rot == 0 else cx - _bx/2, bygg_y - 1.5 if bygg_rot == 0 else cy - _by/2 - 1.5),
                    arrowprops=dict(arrowstyle="<->", color="#555", lw=1))

    ax_pl.set_xlim(-1, tomt_side_input + 1)
    ax_pl.set_ylim(-1, tomt_side_input + 1)
    ax_pl.set_aspect("equal")
    ax_pl.set_xlabel("X [m]")
    ax_pl.set_ylabel("Y [m]")
    ax_pl.set_title(f"Bygningsplassering på tomt  |  Bygg: {_bx:.0f}x{_by:.0f} m  |  Offset: ({bygg_x:.1f}, {bygg_y:.1f}) m  |  Rot: {bygg_rot:.0f} grader")
    ax_pl.grid(True, alpha=0.3, linestyle="--")
    ax_pl.legend(loc="lower right", fontsize=9)

    # Dimensjonslinjer
    if bygg_rot == 0:
        ax_pl.annotate(f"{_bx:.0f} m", xy=(cx, bygg_y - 2.5), ha="center", fontsize=9, color="#333")
        ax_pl.annotate(f"{_by:.0f} m", xy=(bygg_x - 2.5, cy), ha="center", fontsize=9, color="#333", rotation=90)

    pl_col1, pl_col2 = st.columns([2, 1])
    with pl_col1:
        st.pyplot(fig_pl)
        plt.close(fig_pl)
    with pl_col2:
        st.markdown("**Plasseringsinfo**")
        metric_card("Bygg X-bredde", f"{_bx:.1f} m")
        metric_card("Bygg Y-dybde", f"{_by:.1f} m")
        metric_card("Senter X", f"{cx:.1f} m")
        metric_card("Senter Y", f"{cy:.1f} m")
        metric_card("Rotasjon", f"{bygg_rot:.0f}°")
        metric_card("Tomteside", f"{tomt_side_input:.0f} m")
        st.info("💡 Plasseringen brukes automatisk ved IFC-eksport under.")

    st.divider()

    # -----------------------------------------------------------------------
    # FORHÅNDSVISNING
    # -----------------------------------------------------------------------
    with st.expander("👁️ Forhåndsvis rapport", expanded=False):
        st.markdown("Dette er en forhåndsvisning av hva Word/PDF-rapporten vil inneholde.")
        _prev_rows = []
        if show_bg_hoved:
            _prev_rows.append(("🏗️ Bygggenerator – nøkkeltall", [
                f"Elementer: {data_bg.get('rapport_bg_elementer', '–')}",
                f"Kostnad: {data_bg.get('rapport_bg_kostnad', '–')}",
                f"CO₂: {data_bg.get('rapport_bg_co2', '–')}",
            ]))
        if show_terreng:
            _prev_rows.append(("🌍 Terreng", [
                f"Tomteareal: {data_grunn.get('rapport_tomteareal', '–')}",
                f"Prosjektkote: {data_grunn.get('rapport_prosjektkote', '–')}",
            ]))
        if show_geoteknikk:
            _prev_rows.append(("🪨 Geoteknikk", [
                f"Jordart: {data_grunn.get('rapport_jordart', '–')}",
                f"Anbefalt fundament: {data_grunn.get('rapport_fundament_anbefalt', '–')}",
            ]))
        if show_co2_grunn:
            _prev_rows.append(("🌿 CO₂ grunn", [
                f"Total CO₂: {data_grunn.get('rapport_grunn_co2', '–')}",
            ]))
        if _prev_rows:
            for _section, _lines in _prev_rows:
                st.markdown(f"**{_section}**")
                for _line in _lines:
                    st.markdown(f"- {_line}")
                st.markdown("")
        else:
            st.info("Ingen data tilgjengelig ennå. Fyll inn data i Bygggenerator og Grunn-sidene.")

    st.divider()

    # EKSPORT
    # -----------------------------------------------------------------------
    if show_eksport:
        st.subheader("📥 Eksport")
        ex1, ex2, ex3 = st.columns(3)

        with ex1:
            st.markdown("**Word / PDF**")
            if not data.empty:
                summary_dict = make_report_summary_dict(filename, data)
                mat_r = material_summary if not material_summary.empty else pd.DataFrame({"Info": ["Ingen materialdata"]})
                docx_bytes = build_docx_report(summary_dict, mat_r)
                pdf_bytes  = build_pdf_report(summary_dict, mat_r)
                if docx_bytes:
                    st.download_button("📄 Word-rapport", data=docx_bytes,
                                       file_name="byggtotal_rapport.docx",
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                if pdf_bytes:
                    st.download_button("📄 PDF-rapport", data=pdf_bytes,
                                       file_name="byggtotal_rapport.pdf", mime="application/pdf")
            else:
                st.info("Generer et bygg for Word/PDF-eksport.")

        with ex2:
            st.markdown("**CSV**")
            bg_qty_r = data_bg.get("rapport_bg_qty_df", pd.DataFrame())
            if not bg_qty_r.empty:
                st.download_button("📊 Mengder CSV", data=bg_qty_r.to_csv(index=False).encode("utf-8-sig"),
                                   file_name="byggtotal_mengder.csv", mime="text/csv")
            co2_gr_r = data_grunn.get("rapport_grunn_co2_df", pd.DataFrame())
            if not co2_gr_r.empty:
                st.download_button("📊 CO₂ grunn CSV", data=co2_gr_r.to_csv(index=False).encode("utf-8-sig"),
                                   file_name="byggtotal_co2_grunn.csv", mime="text/csv")

        with ex3:
            st.markdown("**IFC – komplett modell**")
            if ifcopenshell is None:
                st.warning("`ifcopenshell` mangler.")
            else:
                ifc_n = st.number_input("Jordlag i IFC", min_value=0, max_value=6, value=2, step=1, key="ifc_n_r2")
                ifc_layers_r = []
                for li in range(int(ifc_n)):
                    sc, st2 = st.columns(2)
                    with sc:
                        s_type = st.selectbox(f"Lag {li+1} type", list(SOIL_DATABASE.keys()), key=f"ifc_soil_r2_{li}")
                    with st2:
                        s_thk = st.number_input(f"Tykkelse [m]", min_value=0.1, value=3.0, step=0.5, key=f"ifc_thk_r2_{li}")
                    ifc_layers_r.append({"soil_type": s_type, "thickness_m": s_thk})

                ifc_gwl_r2      = st.number_input("GVN dybde [m]", min_value=0.0, value=2.5, step=0.5, key="ifc_gwl_r2")
                ifc_found_r2    = st.selectbox("Fundamenttype", list(FOUNDATION_DATABASE.keys()),
                                               format_func=lambda k: FOUNDATION_DATABASE[k]["label"], key="ifc_found_r2")
                ifc_area_r2     = st.number_input("Fundamentareal [m²]", min_value=0.0, value=200.0, step=10.0, key="ifc_area_r2")
                ifc_npiles_r2   = st.number_input("Antall peler", min_value=0, max_value=200, value=0, step=1, key="ifc_np_r2")
                ifc_plen_r2     = st.number_input("Pellengde [m]", min_value=0.0, value=8.0, step=0.5, key="ifc_pl_r2")

                if st.button("🏗️ Generer komplett IFC", type="primary", key="btn_ifc_r2"):
                    try:
                        def _safe_area(s):
                            try:
                                return float(str(s).replace(" ","").replace("m²","").replace(",","."))
                            except Exception:
                                return 500.0
                        # Hent bygningsplassering direkte fra widget-nøklene
                        _off_x   = ss.get("bygg_offset_x", 0.0)
                        _off_y   = ss.get("bygg_offset_y", 0.0)
                        _rot_deg = ss.get("bygg_rot", 0.0)

                        with st.spinner("Genererer IFC med bygningsplassering ..."):
                            ifc_out = generate_complete_ifc_bytes(
                                frame_df           = ss.get("bg_frame_df_last", pd.DataFrame()),
                                slab_df            = ss.get("bg_slab_df_last",  pd.DataFrame()),
                                params             = ss.get("bg_params_last",   {}),
                                ground_layers      = ifc_layers_r or None,
                                gwl_depth_m        = ifc_gwl_r2,
                                foundation_key     = ifc_found_r2,
                                foundation_area_m2 = ifc_area_r2,
                                n_piles            = int(ifc_npiles_r2),
                                pile_length_m      = ifc_plen_r2,
                                pile_lm            = ifc_npiles_r2 * ifc_plen_r2,
                                site_area_m2       = ss.get("tomt_side_input", 35.0) ** 2,
                                building_offset_x  = _off_x,
                                building_offset_y  = _off_y,
                                building_rotation_deg = _rot_deg,
                            )
                        st.success(f"Ferdig! ({len(ifc_out)//1024} kB)  |  Offset: ({_off_x:.1f}, {_off_y:.1f}) m, {_rot_deg:.0f} grader")
                        st.download_button("⬇️ Last ned IFC", data=ifc_out,
                                           file_name="byggtotal_komplett.ifc",
                                           mime="application/octet-stream", key="dl_ifc_r2")
                    except Exception as e:
                        st.error(f"Feil ved IFC-generering: {e}")

elif valg == "Innstillinger":
    if st.button("← Hjem", key="hjem_innstillinger"): _nav_to("Hjem")
    st.header("⚙️ Innstillinger")
    st.caption("Endringer lagres automatisk og gjelder for hele økten.")

    st.subheader("Produktvalg fra Norsk Prisbok")
    _is1, _is2, _is3 = st.columns(3)
    with _is1:
        _dv = st.selectbox("Dekkeløsning", ["Hulldekke", "Hulldekke_lavCO2"],
            index=["Hulldekke","Hulldekke_lavCO2"].index(st.session_state["deck_variant_key"]),
            format_func=lambda x: MATERIAL_DATABASE[x]["label"])
        st.session_state["deck_variant_key"] = _dv
    with _is2:
        _cv = st.selectbox("Plasstøpt betong", ["Plasstøpt_betong","Plasstøpt_betong_lavCO2"],
            index=["Plasstøpt_betong","Plasstøpt_betong_lavCO2"].index(st.session_state["concrete_variant_key"]),
            format_func=lambda x: MATERIAL_DATABASE[x]["label"])
        st.session_state["concrete_variant_key"] = _cv
    with _is3:
        _wv = st.selectbox("Betongvegg", ["Betong_vegg","Betong_vegg_lavCO2"],
            index=["Betong_vegg","Betong_vegg_lavCO2"].index(st.session_state["wall_variant_key"]),
            format_func=lambda x: MATERIAL_DATABASE[x]["label"])
        st.session_state["wall_variant_key"] = _wv

    st.divider()
    st.subheader("Materialegenskaper")
    _ma1, _ma2 = st.columns(2)
    with _ma1:
        st.session_state["_glulam_density"] = st.number_input("Tetthet limtre (kg/m³)", min_value=100.0, max_value=900.0, value=float(st.session_state["_glulam_density"]), step=10.0)
    with _ma2:
        st.session_state["_clt_density"] = st.number_input("Tetthet massivtre / CLT (kg/m³)", min_value=100.0, max_value=900.0, value=float(st.session_state["_clt_density"]), step=10.0)

    st.divider()
    st.subheader("CO₂-kilde og visning")
    _co1, _co2 = st.columns(2)
    with _co1:
        st.session_state["_use_epd"] = st.toggle("Bruk EPD-/prosjektfaktorer som primær CO₂-kilde", value=st.session_state["_use_epd"])
    with _co2:
        st.session_state["_show_raw"] = st.toggle("Vis rådata i Mengder", value=st.session_state["_show_raw"])

    st.divider()
    st.subheader("Ytelse (IFC og 3D)")
    _yt1, _yt2 = st.columns(2)
    with _yt1:
        st.session_state["_fast_mode"] = st.toggle("Rask modus for IFC", value=st.session_state["_fast_mode"], help="Hopper over tunge geometriestimater ved IFC-innlasting.")
        st.session_state["_use_geom_fallback"] = st.toggle("Bruk geometriestimat ved manglende IFC-mengder", value=st.session_state["_use_geom_fallback"])
    with _yt2:
        st.session_state["_lazy_3d"] = st.toggle("Last 3D først når jeg klikker", value=st.session_state["_lazy_3d"])
        st.session_state["_profile_limit"] = float(st.number_input("Maks profiler i filter", min_value=20, max_value=500, value=int(st.session_state["_profile_limit"]), step=10))

    st.divider()
    if st.button("← Tilbake til Hjem", use_container_width=True):
        _nav_to("Hjem")

st.markdown("---")
st.markdown("**byggTotal**")