
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


st.set_page_config(
    page_title="byggTotal – Mengder, kalkyle, IFC og prosjektering",
    page_icon="🏗️",
    layout="wide",
)

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


def build_dataset_from_ifc(ifc_bytes: bytes):
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

                if all(pd.isna(v) or safe_num(v) == 0 for v in [length_m, area_m2, volume_m3]):
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

def load_workbook_values(file_bytes: bytes):
    if openpyxl is None:
        raise ImportError("openpyxl er ikke installert.")
    return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)


def read_cell(ws, cell, default=None):
    try:
        value = ws[cell].value
        return default if value is None else value
    except Exception:
        return default


def safe_bool_ja_nei(value) -> str:
    text = str(value or "").strip().upper()
    return "JA" if text == "JA" else "NEI"


def load_project_parameters_from_excel(file_bytes: bytes) -> dict:
    wb = load_workbook_values(file_bytes)
    if "GEOMETRI" not in wb.sheetnames or "FORSIDE" not in wb.sheetnames:
        raise ValueError("Excel-filen mangler FORSIDE og/eller GEOMETRI.")
    wsf = wb["FORSIDE"]
    wsg = wb["GEOMETRI"]

    params = {
        "bjelkemateriale": read_cell(wsf, "B3", "Stål"),
        "søylemateriale": read_cell(wsf, "B4", "Limtre"),
        "antall_fag_x_r1": safe_num(read_cell(wsf, "B5", read_cell(wsg, "B4", 1))),
        "antall_fag_y_r1": safe_num(read_cell(wsf, "B8", read_cell(wsg, "B5", 1))),
        "antall_etasjer": safe_num(read_cell(wsf, "F8", read_cell(wsg, "B36", 1))),
        "rektangel2_aktiv": safe_bool_ja_nei(read_cell(wsg, "B2", "NEI")),
        "fag_x_r1": safe_num(read_cell(wsg, "B4", 1)),
        "fag_y_r1": safe_num(read_cell(wsg, "B5", 1)),
        "faglengde_x_mm": safe_num(read_cell(wsg, "B6", 12000)),
        "faglengde_y_mm": safe_num(read_cell(wsg, "B7", 16000)),
        "fag_x_r2": safe_num(read_cell(wsg, "B10", 0)),
        "fag_y_r2": safe_num(read_cell(wsg, "B11", 0)),
        "r2_offset_x_fag": safe_num(read_cell(wsg, "B16", 0)),
        "r2_offset_y_fag": safe_num(read_cell(wsg, "B17", 0)),
        "dekker_aktiv": safe_bool_ja_nei(read_cell(wsg, "B26", "JA")),
        "skalltype": read_cell(wsg, "B28", "Platt skall"),
        "dekke_tykkelse_mm": safe_num(read_cell(wsg, "B29", 300)),
        "dekker_i_modell": safe_num(read_cell(wsg, "B36", 1)),
        "dekke_materiale": read_cell(wsg, "B37", "B35, Betong"),
        "opening_width_fag": safe_num(read_cell(wsg, "B41", 0)),
        "opening_height_fag": safe_num(read_cell(wsg, "B42", 0)),
        "opening_offset_x_fag": safe_num(read_cell(wsg, "B43", 0)),
        "opening_offset_y_fag": safe_num(read_cell(wsg, "B44", 0)),
        "aktivt_dekkeareal_excel_m2": safe_num(read_cell(wsg, "B52", 0)),
        "eksportprinsipp": read_cell(wsg, "B57", ""),
    }
    return params


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


def generate_frame_export(params: dict) -> pd.DataFrame:
    geom = generate_plan_geometry(params)
    dx = geom["dx"]
    dy = geom["dy"]
    fx = max(int(round(safe_num(params.get("fag_x_r1", 1)))), 1)
    fy = max(int(round(safe_num(params.get("fag_y_r1", 1)))), 1)
    etasjeh = 4.0
    n_levels = max(int(round(safe_num(params.get("antall_etasjer", params.get("dekker_i_modell", 1))))), 1)

    rows = []
    col_id = 1
    beam_id = 1

    for level in range(1, n_levels + 1):
        z0 = (level - 1) * etasjeh
        z1 = level * etasjeh
        for ix in range(fx + 1):
            for iy in range(fy + 1):
                x = ix * dx
                y = iy * dy
                rows.append({"Type": "Søyle", "ID": f"C{col_id}", "Nivå": level, "X1 [m]": x, "Y1 [m]": y, "Z1 [m]": z0, "X2 [m]": x, "Y2 [m]": y, "Z2 [m]": z1})
                col_id += 1

        # X-retning bjelker
        for iy in range(fy + 1):
            y = iy * dy
            for ix in range(fx):
                x1 = ix * dx
                x2 = (ix + 1) * dx
                rows.append({"Type": "Bjelke", "ID": f"B{beam_id}", "Nivå": level, "X1 [m]": x1, "Y1 [m]": y, "Z1 [m]": z1, "X2 [m]": x2, "Y2 [m]": y, "Z2 [m]": z1})
                beam_id += 1

        # Y-retning bjelker
        for ix in range(fx + 1):
            x = ix * dx
            for iy in range(fy):
                y1 = iy * dy
                y2 = (iy + 1) * dy
                rows.append({"Type": "Bjelke", "ID": f"B{beam_id}", "Nivå": level, "X1 [m]": x, "Y1 [m]": y1, "Z1 [m]": z1, "X2 [m]": x, "Y2 [m]": y2, "Z2 [m]": z1})
                beam_id += 1

    return pd.DataFrame(rows)


def format_point(x, y):
    return f"{int(round(x * 1000))}, {int(round(y * 1000))}"


def generate_slab_export(params: dict) -> pd.DataFrame:
    geom = generate_plan_geometry(params)
    n_decks = max(int(round(safe_num(params.get("dekker_i_modell", 1)))), 1)
    rows = []
    rectangles = geom["rectangles"]
    opening = geom["opening"]

    for i in range(n_decks):
        z_mm = int(round((i + 1) * 4000))
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


def read_export_sheet_from_excel(file_bytes: bytes, sheet_name: str, header_row_idx: int):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row_idx, engine="openpyxl")
    return clean_dataframe(df)


def build_docx_report(summary_dict, material_summary, extra_sections=None):
    if Document is None:
        return None
    doc = Document()
    doc.add_heading("byggTotal – Prosjektrapport", 0)
    p = doc.add_paragraph()
    p.add_run("Generert: ").bold = True
    p.add_run(datetime.now().strftime("%d.%m.%Y %H:%M"))

    doc.add_heading("Prosjektoversikt", level=1)
    for k, v in summary_dict.items():
        doc.add_paragraph(f"{k}: {v}")

    doc.add_heading("Materialoversikt", level=1)
    table = doc.add_table(rows=1, cols=len(material_summary.columns))
    table.style = "Table Grid"
    for i, col in enumerate(material_summary.columns):
        table.rows[0].cells[i].text = str(col)
    for _, row in material_summary.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = f"{val}" if not isinstance(val, float) else f"{val:,.2f}".replace(",", " ")

    if extra_sections:
        for title, df in extra_sections:
            if df is None or df.empty:
                continue
            doc.add_heading(title, level=1)
            t = doc.add_table(rows=1, cols=len(df.columns))
            t.style = "Table Grid"
            for i, col in enumerate(df.columns):
                t.rows[0].cells[i].text = str(col)
            for _, row in df.iterrows():
                cells = t.add_row().cells
                for i, val in enumerate(row):
                    cells[i].text = f"{val}" if not isinstance(val, float) else f"{val:,.2f}".replace(",", " ")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


def build_pdf_report(summary_dict, material_summary, extra_sections=None):
    if SimpleDocTemplate is None:
        return None
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = [Paragraph("byggTotal – Prosjektrapport", styles["Title"]), Spacer(1, 12), Paragraph(f"Generert: {datetime.now().strftime('%d.%m.%Y %H:%M')}", styles["Normal"]), Spacer(1, 12)]

    summary_table_data = [["Parameter", "Verdi"]] + [[str(k), str(v)] for k, v in summary_dict.items()]
    t1 = Table(summary_table_data, hAlign="LEFT")
    t1.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey)]))
    elements += [Paragraph("Prosjektoversikt", styles["Heading2"]), t1, Spacer(1, 16)]

    t2 = Table([list(material_summary.columns)] + material_summary.astype(str).values.tolist(), hAlign="LEFT")
    t2.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey)]))
    elements += [Paragraph("Materialoversikt", styles["Heading2"]), t2, Spacer(1, 16)]

    if extra_sections:
        for title, df in extra_sections:
            if df is None or df.empty:
                continue
            t = Table([list(df.columns)] + df.astype(str).values.tolist(), hAlign="LEFT")
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey)]))
            elements += [Paragraph(title, styles["Heading2"]), t, Spacer(1, 16)]

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


def make_report_summary_dict(filename, filtered_df):
    return {
        "Fil": filename,
        "Antall elementer": int(len(filtered_df)),
        "Total lengde [m]": float(pd.to_numeric(filtered_df["Lengde [m]"], errors="coerce").fillna(0).sum()),
        "Total areal [m2]": float(pd.to_numeric(filtered_df["Areal [m2]"], errors="coerce").fillna(0).sum()),
        "Total volum [m3]": float(pd.to_numeric(filtered_df["Volum [m3]"], errors="coerce").fillna(0).sum()),
        "Total vekt [kg]": float(pd.to_numeric(filtered_df["Vekt [kg]"], errors="coerce").fillna(0).sum()),
        "Total kostnad [kr]": float(pd.to_numeric(filtered_df["Kostnad [kr]"], errors="coerce").fillna(0).sum()),
        "Total CO2 [kgCO2e]": float(pd.to_numeric(filtered_df["CO2 [kgCO2e]"], errors="coerce").fillna(0).sum()),
    }


st.markdown("""
<div class="custom-card">
    <div style="font-size:42px; font-weight:800; color:#1f2937;">byggTotal</div>
    <div style="font-size:20px; font-weight:600; color:#374151; margin-top:6px;">
        Mengde-, kalkyle-, IFC- og prosjekteringsverktøy
    </div>
    <div style="margin-top:10px; color:#6b7280; font-size:15px;">
        Leser Excel og IFC, gir mengder, kostnader, CO₂, materialbytte, 3D-visning og en ny modul for prosjektering.
    </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.title("byggTotal")
valg = st.sidebar.radio(
    "Velg side",
    ["Mengder", "Pristilbud", "Analyse", "Materialbytte", "CO₂-regnskap", "3D-modell", "Prosjektering", "Rapport"],
)

with st.sidebar:
    st.header("Fil og innstillinger")
    uploaded_excel = st.file_uploader("Last opp Excel-fil (.xlsx)", type=["xlsx"])
    uploaded_ifc = st.file_uploader("Last opp IFC-fil (.ifc)", type=["ifc"])

    st.subheader("Produktvalg fra prisbok")
    deck_variant = st.selectbox("Dekkeløsning", ["Hulldekke", "Hulldekke_lavCO2"], format_func=lambda x: MATERIAL_DATABASE[x]["label"])
    concrete_variant = st.selectbox("Plasstøpt betong", ["Plasstøpt_betong", "Plasstøpt_betong_lavCO2"], format_func=lambda x: MATERIAL_DATABASE[x]["label"])
    wall_variant = st.selectbox("Betongvegg", ["Betong_vegg", "Betong_vegg_lavCO2"], format_func=lambda x: MATERIAL_DATABASE[x]["label"])

    st.subheader("Materialegenskaper")
    glulam_density = st.number_input("Tetthet limtre (kg/m³)", min_value=100.0, value=460.0, step=10.0)
    clt_density = st.number_input("Tetthet massivtre / CLT (kg/m³)", min_value=100.0, value=500.0, step=10.0)

    st.subheader("CO₂-kilde")
    use_epd = st.toggle("Bruk EPD-/prosjektfaktorer som primær CO₂-kilde", value=True)

    st.subheader("Visning")
    show_raw = st.toggle("Vis rådata", value=False)

GLULAM_DENSITY = glulam_density
CLT_DENSITY = clt_density
MATERIAL_DATABASE["Limtre"]["density"] = glulam_density
MATERIAL_DATABASE["Massivtre"]["density"] = clt_density
MATERIAL_DATABASE["Massivtre_vegg"]["density"] = clt_density

data = None
nodes = pd.DataFrame()
forside = pd.DataFrame()
filename = None
excel_supports_prosjektering = False
project_params = {}
prosjekt_frame_df = pd.DataFrame()
prosjekt_slab_df = pd.DataFrame()
prosjekt_qa_df = pd.DataFrame()
excel_export_frame_df = pd.DataFrame()
excel_export_slab_df = pd.DataFrame()
excel_qa_df = pd.DataFrame()

try:
    if uploaded_ifc is not None:
        filename = uploaded_ifc.name
        data, nodes, forside = build_dataset_from_ifc(uploaded_ifc.getvalue())
    elif uploaded_excel is not None:
        filename = uploaded_excel.name
        try:
            data, nodes, forside = build_dataset_from_excel(uploaded_excel.getvalue())
        except Exception:
            # Tillat Prosjektering selv om MENGDER/Segmenter/Knutepunkter mangler
            data = pd.DataFrame(columns=["Segment", "Type", "Knutepunkter", "Material / Tverrsnitt", "Lengde [m]", "Areal [m2]", "Volum [m3]", "Vekt [kg]", "materiale", "Endret IFC", "Mengdegrunnlag"])
            nodes = pd.DataFrame()
            forside = pd.DataFrame()
        try:
            project_params = load_project_parameters_from_excel(uploaded_excel.getvalue())
            prosjekt_frame_df = generate_frame_export(project_params)
            prosjekt_slab_df = generate_slab_export(project_params)
            prosjekt_qa_df = run_project_qa(project_params, prosjekt_frame_df, prosjekt_slab_df)
            excel_supports_prosjektering = True
            try:
                excel_export_frame_df = read_export_sheet_from_excel(uploaded_excel.getvalue(), "EXPORT_RAMME", 2)
            except Exception:
                excel_export_frame_df = pd.DataFrame()
            try:
                excel_export_slab_df = read_export_sheet_from_excel(uploaded_excel.getvalue(), "EXPORT_FOCUS", 9)
            except Exception:
                excel_export_slab_df = pd.DataFrame()
            try:
                excel_qa_df = read_export_sheet_from_excel(uploaded_excel.getvalue(), "QA_IFC_kontroll", 2)
            except Exception:
                excel_qa_df = pd.DataFrame()
        except Exception:
            excel_supports_prosjektering = False
    else:
        st.info("Last opp en Excel-fil eller IFC-fil i sidepanelet for å starte analysen.")
        st.stop()
except Exception as e:
    st.error(f"Kunne ikke lese filen: {e}")
    st.stop()

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

with st.container():
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        type_options = sorted([x for x in data["Type"].dropna().astype(str).unique().tolist()]) if not data.empty else []
        selected_types = st.multiselect("Type", type_options, default=type_options)
    with c2:
        mat_options = sorted([x for x in data["materiale"].dropna().astype(str).unique().tolist()]) if not data.empty else []
        selected_materials = st.multiselect("Materiale", mat_options, default=mat_options)
    with c3:
        profile_options = sorted([x for x in data["Material / Tverrsnitt"].dropna().astype(str).unique().tolist()]) if not data.empty else []
        selected_profiles = st.multiselect("Profil / tverrsnitt", profile_options, default=profile_options[:8] if len(profile_options) > 8 else profile_options)
    with c4:
        max_length = float(pd.to_numeric(data.get("Lengde [m]", pd.Series(dtype=float)), errors="coerce").fillna(0).max() or 0)
        length_range = st.slider("Lengdeintervall [m]", 0.0, max(1.0, max_length), (0.0, max(1.0, max_length)))

filtered = data.copy()
if not filtered.empty:
    if selected_types:
        filtered = filtered[filtered["Type"].isin(selected_types)]
    if selected_materials:
        filtered = filtered[filtered["materiale"].isin(selected_materials)]
    if selected_profiles:
        filtered = filtered[filtered["Material / Tverrsnitt"].isin(selected_profiles)]
    filtered = filtered[(pd.to_numeric(filtered["Lengde [m]"], errors="coerce").fillna(0) >= length_range[0]) & (pd.to_numeric(filtered["Lengde [m]"], errors="coerce").fillna(0) <= length_range[1])]
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
    st.header("📊 Mengder")
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
        st.subheader("Kostnadsfordeling")
        pie_data = summary[summary["kostnad_kr"] > 0].copy() if not summary.empty else pd.DataFrame()
        if not pie_data.empty:
            pie_data["navn"] = pie_data["Type"].fillna("Ukjent") + " – " + pie_data["materiale"].fillna("Ukjent")
            fig1, ax1 = plt.subplots(figsize=(6, 5))
            ax1.pie(pie_data["kostnad_kr"], labels=pie_data["navn"], autopct="%1.1f%%", startangle=90)
            ax1.axis("equal")
            st.pyplot(fig1)
        else:
            st.info("Ingen kostnadsdata er tilgjengelige for valgt utvalg.")

        st.subheader("CO₂ per produkt")
        co2_data = filtered.groupby("Produktnavn", dropna=False)["CO2 [kgCO2e]"].sum().reset_index() if not filtered.empty else pd.DataFrame()
        co2_data = co2_data[co2_data["CO2 [kgCO2e]"] > 0] if not co2_data.empty else co2_data
        if not co2_data.empty:
            fig2, ax2 = plt.subplots(figsize=(6, 4))
            ax2.bar(co2_data["Produktnavn"].fillna("Ukjent"), co2_data["CO2 [kgCO2e]"])
            plt.xticks(rotation=25, ha="right")
            st.pyplot(fig2)
        else:
            st.info("Ingen CO₂-data er tilgjengelige for valgt utvalg.")

    show_cols = [c for c in ["Segment", "Type", "Knutepunkter", "Material / Tverrsnitt", "materiale", "Produktnøkkel", "Produktnavn", "NS3420-kode", "Mengdegrunnlag", "Endret IFC", "Lengde [m]", "Areal [m2]", "Volum [m3]", "Vekt [kg]", "Kostnad [kr]", "CO2 [kgCO2e]", "IFC Type", "IFC GlobalId"] if c in filtered.columns]
    st.subheader("Filtrerte elementer")
    st.dataframe(filtered[show_cols], use_container_width=True, hide_index=True)
    st.download_button("Last ned filtrerte data som CSV", filtered[show_cols].to_csv(index=False).encode("utf-8-sig"), file_name="filtrerte_mengder.csv", mime="text/csv")
    if show_raw:
        with st.expander("Rådata"):
            st.dataframe(data, use_container_width=True)

elif valg == "Pristilbud":
    st.header("💰 Pristilbud")
    total_staal_kg = filtered.loc[filtered["materiale"] == "Stål", "Vekt [kg]"].sum() if not filtered.empty else 0
    total_limtre_m3 = filtered.loc[filtered["materiale"] == "Limtre", "Volum [m3]"].sum() if not filtered.empty else 0
    total_massivtre_m3 = filtered.loc[filtered["materiale"] == "Massivtre", "Volum [m3]"].sum() if not filtered.empty else 0
    total_betong_m3 = filtered.loc[filtered["materiale"] == "Betong", "Volum [m3]"].sum() if not filtered.empty else 0
    total_pris = filtered["Kostnad [kr]"].sum() if not filtered.empty else 0
    total_co2 = filtered["CO2 [kgCO2e]"].sum() if not filtered.empty else 0
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: metric_card("Stålvekt", f"{total_staal_kg:,.0f} kg".replace(",", " "))
    with c2: metric_card("Limtrevolum", f"{total_limtre_m3:,.2f} m³".replace(",", " "))
    with c3: metric_card("Massivtrevolum", f"{total_massivtre_m3:,.2f} m³".replace(",", " "))
    with c4: metric_card("Betongvolum", f"{total_betong_m3:,.2f} m³".replace(",", " "))
    with c5: metric_card("Total estimert pris", f"{total_pris:,.0f} kr".replace(",", " "))
    with c6: metric_card("Total CO₂", f"{total_co2:,.0f} kgCO₂e".replace(",", " "))

    tilbud = (
        filtered.groupby(["materiale", "Produktnavn", "Material / Tverrsnitt", "NS3420-kode"], dropna=False)
        .agg(antall=("Segment", "count"), areal_m2=("Areal [m2]", "sum"), lengde_m=("Lengde [m]", "sum"), volum_m3=("Volum [m3]", "sum"), vekt_kg=("Vekt [kg]", "sum"), kostnad_kr=("Kostnad [kr]", "sum"), co2_kg=("CO2 [kgCO2e]", "sum"))
        .reset_index().sort_values("kostnad_kr", ascending=False)
    ) if not filtered.empty else pd.DataFrame()

    st.subheader("Tilbudsgrunnlag")
    st.dataframe(tilbud, use_container_width=True, hide_index=True)
    st.download_button("Last ned tilbud som CSV", tilbud.to_csv(index=False).encode("utf-8-sig"), file_name="pristilbud.csv", mime="text/csv")

elif valg == "Analyse":
    st.header("📈 Analyse")
    st.subheader("Materialfordeling")
    st.bar_chart(filtered["materiale"].value_counts(dropna=False) if not filtered.empty else pd.Series(dtype=float))
    st.subheader("Kostnad per type")
    st.bar_chart(filtered.groupby("Type", dropna=False)["Kostnad [kr]"].sum() if not filtered.empty else pd.Series(dtype=float))
    st.subheader("Areal per type")
    st.bar_chart(filtered.groupby("Type", dropna=False)["Areal [m2]"].sum() if not filtered.empty else pd.Series(dtype=float))
    st.subheader("CO₂ per produkt")
    st.bar_chart(filtered.groupby("Produktnavn", dropna=False)["CO2 [kgCO2e]"].sum() if not filtered.empty else pd.Series(dtype=float))
    st.subheader("Mengdegrunnlag")
    st.bar_chart(filtered["Mengdegrunnlag"].value_counts(dropna=False) if not filtered.empty else pd.Series(dtype=float))

elif valg == "Materialbytte":
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
            with m4: metric_card("Kostnadsendring", f"{swap_df['Kostnadsendring [kr]'].sum():,.0f} kr".replace(",", " "))
            st.dataframe(swap_df, use_container_width=True, hide_index=True)

            if uploaded_ifc is not None and "IFC GlobalId" in swap_df.columns:
                preview_ids = tuple(sorted(set(swap_df["IFC GlobalId"].dropna().astype(str).tolist())))
                try:
                    preview_meshes = extract_ifc_meshes_filtered(uploaded_ifc.getvalue(), visible_ids_tuple=None, max_elements=1500)
                    fig_preview = build_ifc_3d_figure(preview_meshes, preview_ids=preview_ids, show_only_preview=False, preview_material=defaults["label"])
                    st.plotly_chart(fig_preview, use_container_width=True)
                except Exception as e:
                    st.warning(f"Kunne ikke generere 3D-forhåndsvisning: {e}")

                if st.button("Generer ny IFC-fil"):
                    try:
                        new_ifc_bytes, ifc_change_log = export_ifc_material_swap(uploaded_ifc.getvalue(), data, selected_swap_type, from_material, target_key, new_profile_text)
                        if new_ifc_bytes is None or ifc_change_log.empty:
                            st.warning("Ingen elementer ble oppdatert i IFC-filen.")
                        else:
                            out_name = Path(uploaded_ifc.name).stem + f"_materialbytte_{re.sub(r'[^A-Za-z0-9_-]+', '_', str(target_key))}.ifc"
                            st.download_button("Last ned ny IFC-fil", data=new_ifc_bytes, file_name=out_name, mime="application/octet-stream")
                            st.dataframe(ifc_change_log, use_container_width=True, hide_index=True)
                    except Exception as e:
                        st.error(f"Kunne ikke generere IFC-fil: {e}")

elif valg == "CO₂-regnskap":
    st.header("🌍 CO₂-regnskap")
    total_co2 = filtered["CO2 [kgCO2e]"].sum() if not filtered.empty else 0
    total_cost = filtered["Kostnad [kr]"].sum() if not filtered.empty else 0
    c1, c2, c3 = st.columns(3)
    with c1: metric_card("Totalt CO₂-avtrykk", f"{total_co2:,.0f} kgCO₂e".replace(",", " "))
    with c2: metric_card("Estimert kostnad", f"{total_cost:,.0f} kr".replace(",", " "))
    with c3: metric_card("CO₂ per element", f"{(total_co2 / len(filtered)) if len(filtered) > 0 else 0:,.1f}".replace(",", " "))

    co2_material = (
        filtered.groupby(["materiale", "Produktnavn", "NS3420-kode"], dropna=False)
        .agg(antall=("Segment", "count"), areal_m2=("Areal [m2]", "sum"), volum_m3=("Volum [m3]", "sum"), vekt_kg=("Vekt [kg]", "sum"), co2_kg=("CO2 [kgCO2e]", "sum"), kostnad_kr=("Kostnad [kr]", "sum"))
        .reset_index().sort_values("co2_kg", ascending=False)
    ) if not filtered.empty else pd.DataFrame()
    st.dataframe(co2_material, use_container_width=True, hide_index=True)

    left, right = st.columns(2)
    with left:
        if not co2_material.empty:
            fig4, ax4 = plt.subplots(figsize=(6, 4))
            ax4.bar(co2_material["Produktnavn"].fillna("Ukjent"), co2_material["co2_kg"])
            plt.xticks(rotation=25, ha="right")
            st.pyplot(fig4)
    with right:
        if not co2_material.empty:
            fig5, ax5 = plt.subplots(figsize=(6, 4))
            ax5.bar(co2_material["Produktnavn"].fillna("Ukjent"), co2_material["kostnad_kr"])
            plt.xticks(rotation=25, ha="right")
            st.pyplot(fig5)

elif valg == "3D-modell":
    st.header("🧊 3D-modellvisning")
    if uploaded_ifc is None:
        st.info("3D-modellvisning er tilgjengelig når en IFC-fil er lastet opp.")
    else:
        visning = st.radio("Visning", ["Kun filtrerte elementer", "Alle elementer"], horizontal=True)
        max_elements_3d = st.slider("Maks antall elementer i 3D-visning", 100, 5000, 1500, 100)
        visible_ids = tuple(sorted(set(filtered["IFC GlobalId"].dropna().astype(str).tolist()))) if visning == "Kun filtrerte elementer" and "IFC GlobalId" in filtered.columns else None
        try:
            meshes = extract_ifc_meshes_filtered(uploaded_ifc.getvalue(), visible_ids_tuple=visible_ids, max_elements=max_elements_3d)
            if meshes:
                fig3d = build_ifc_3d_figure(meshes)
                st.plotly_chart(fig3d, use_container_width=True)
            else:
                st.warning("Ingen 3D-geometri ble funnet for valgt utvalg.")
        except Exception as e:
            st.error(f"Kunne ikke generere 3D-visning: {e}")

elif valg == "Prosjektering":
    st.header("🧩 Prosjektering")
    if uploaded_excel is None:
        st.info("Last opp Excel-filen med FORSIDE/GEOMETRI for å bruke prosjekteringsmodulen.")
    elif not excel_supports_prosjektering:
        st.warning("Excel-filen ble lest, men FORSIDE/GEOMETRI kunne ikke tolkes sikkert for prosjektering.")
    else:
        st.caption("Denne modulen er lagt inn som første fungerende Python-versjon av prosjekteringsmotoren. Den leser FORSIDE/GEOMETRI, gir redigerbar input, genererer planform, ramme- og dekkeeksport, samt QA-kontroll.")

        st.subheader("1. Input")
        i1, i2, i3, i4 = st.columns(4)
        with i1:
            project_params["fag_x_r1"] = st.number_input("Fag X R1", min_value=1, value=int(round(project_params["fag_x_r1"])), step=1)
            project_params["faglengde_x_mm"] = st.number_input("Faglengde X [mm]", min_value=1000, value=int(round(project_params["faglengde_x_mm"])), step=100)
            project_params["dekke_tykkelse_mm"] = st.number_input("Dekketøykkelse [mm]", min_value=1, value=int(round(project_params["dekke_tykkelse_mm"])), step=5)
        with i2:
            project_params["fag_y_r1"] = st.number_input("Fag Y R1", min_value=1, value=int(round(project_params["fag_y_r1"])), step=1)
            project_params["faglengde_y_mm"] = st.number_input("Faglengde Y [mm]", min_value=1000, value=int(round(project_params["faglengde_y_mm"])), step=100)
            project_params["dekker_i_modell"] = st.number_input("Dekker i modell", min_value=1, value=int(round(project_params["dekker_i_modell"])), step=1)
        with i3:
            project_params["rektangel2_aktiv"] = "JA" if st.toggle("Aktiver rektangel 2 / tilbygg", value=project_params["rektangel2_aktiv"] == "JA") else "NEI"
            project_params["fag_x_r2"] = st.number_input("Fag X R2", min_value=0, value=int(round(project_params["fag_x_r2"])), step=1)
            project_params["fag_y_r2"] = st.number_input("Fag Y R2", min_value=0, value=int(round(project_params["fag_y_r2"])), step=1)
        with i4:
            project_params["r2_offset_x_fag"] = st.number_input("R2 offset X [fag]", min_value=0, value=int(round(project_params["r2_offset_x_fag"])), step=1)
            project_params["r2_offset_y_fag"] = st.number_input("R2 offset Y [fag]", min_value=0, value=int(round(project_params["r2_offset_y_fag"])), step=1)
            project_params["dekker_aktiv"] = "JA" if st.toggle("Dekker aktive", value=project_params["dekker_aktiv"] == "JA") else "NEI"

        j1, j2, j3, j4 = st.columns(4)
        with j1:
            project_params["opening_width_fag"] = st.number_input("Åpning bredde [fag]", min_value=0, value=int(round(project_params["opening_width_fag"])), step=1)
        with j2:
            project_params["opening_height_fag"] = st.number_input("Åpning høyde [fag]", min_value=0, value=int(round(project_params["opening_height_fag"])), step=1)
        with j3:
            project_params["opening_offset_x_fag"] = st.number_input("Åpning offset X [fag]", min_value=0, value=int(round(project_params["opening_offset_x_fag"])), step=1)
        with j4:
            project_params["opening_offset_y_fag"] = st.number_input("Åpning offset Y [fag]", min_value=0, value=int(round(project_params["opening_offset_y_fag"])), step=1)

        project_params["skalltype"] = st.text_input("Skalltype", value=str(project_params["skalltype"]))
        project_params["dekke_materiale"] = st.text_input("Dekkemateriale", value=str(project_params["dekke_materiale"]))
        project_params["eksportprinsipp"] = st.text_input("Eksportprinsipp", value=str(project_params["eksportprinsipp"]))

        geom = generate_plan_geometry(project_params)
        prosjekt_frame_df = generate_frame_export(project_params)
        prosjekt_slab_df = generate_slab_export(project_params)
        prosjekt_qa_df = run_project_qa(project_params, prosjekt_frame_df, prosjekt_slab_df)

        k1, k2, k3, k4 = st.columns(4)
        with k1: metric_card("Planformkode", geom["planformkode"])
        with k2: metric_card("Bruttoareal", f"{geom['gross_area_m2']:,.1f} m²".replace(",", " "))
        with k3: metric_card("Åpningsareal", f"{geom['opening_area_m2']:,.1f} m²".replace(",", " "))
        with k4: metric_card("Aktivt areal", f"{geom['active_area_m2']:,.1f} m²".replace(",", " "))

        left, right = st.columns([1.05, 1])
        with left:
            st.subheader("2. Geometrilogikk / 2D-preview")
            fig_plan = plot_plan_geometry(geom)
            st.pyplot(fig_plan)
        with right:
            st.subheader("3. QA / kontroll")
            st.dataframe(prosjekt_qa_df, use_container_width=True, hide_index=True)

        a, b = st.columns(2)
        with a:
            st.subheader("4A. Generert rammeeksport (Python)")
            st.dataframe(prosjekt_frame_df, use_container_width=True, hide_index=True, height=420)
            st.download_button("Last ned rammeeksport CSV", prosjekt_frame_df.to_csv(index=False).encode("utf-8-sig"), file_name="export_ramme_python.csv", mime="text/csv")
        with b:
            st.subheader("4B. Generert dekke-/skalleksport (Python)")
            st.dataframe(prosjekt_slab_df, use_container_width=True, hide_index=True, height=420)
            st.download_button("Last ned skalleksport CSV", prosjekt_slab_df.to_csv(index=False).encode("utf-8-sig"), file_name="export_skall_python.csv", mime="text/csv")

        with st.expander("Sammenligning mot Excel-arkets eksportfaner"):
            s1, s2 = st.columns(2)
            with s1:
                st.markdown("**Excel – EXPORT_RAMME**")
                if not excel_export_frame_df.empty:
                    st.dataframe(excel_export_frame_df, use_container_width=True, hide_index=True)
                else:
                    st.info("Fant ikke lesbar EXPORT_RAMME i filen.")
            with s2:
                st.markdown("**Excel – EXPORT_FOCUS**")
                if not excel_export_slab_df.empty:
                    st.dataframe(excel_export_slab_df, use_container_width=True, hide_index=True)
                else:
                    st.info("Fant ikke lesbar EXPORT_FOCUS i filen.")

        with st.expander("Innlastede prosjektparametere"):
            params_df = pd.DataFrame({"Parameter": list(project_params.keys()), "Verdi": list(project_params.values())})
            st.dataframe(params_df, use_container_width=True, hide_index=True)

        if not excel_qa_df.empty:
            with st.expander("Excel – QA_IFC_kontroll"):
                st.dataframe(excel_qa_df, use_container_width=True, hide_index=True)

elif valg == "Rapport":
    st.header("📝 Rapport og eksport")
    summary_dict = make_report_summary_dict(filename, filtered)
    c1, c2, c3 = st.columns(3)
    with c1: metric_card("Elementer", f"{summary_dict['Antall elementer']:,}".replace(",", " "))
    with c2: metric_card("Total kostnad", f"{summary_dict['Total kostnad [kr]']:,.0f} kr".replace(",", " "))
    with c3: metric_card("Total CO₂", f"{summary_dict['Total CO2 [kgCO2e]']:,.0f} kgCO₂e".replace(",", " "))

    report_df = pd.DataFrame({"Parameter": list(summary_dict.keys()), "Verdi": list(summary_dict.values())})
    st.dataframe(report_df, use_container_width=True, hide_index=True)
    st.subheader("Materialoversikt")
    st.dataframe(material_summary, use_container_width=True, hide_index=True)

    extra_sections = []
    if excel_supports_prosjektering:
        include_pros = st.checkbox("Ta med Prosjektering i rapporten", value=True)
        if include_pros:
            extra_sections += [("Prosjektering – QA", prosjekt_qa_df), ("Prosjektering – Rammeeksport", prosjekt_frame_df.head(100)), ("Prosjektering – Skalleksport", prosjekt_slab_df.head(100))]

    docx_bytes = build_docx_report(summary_dict, material_summary if not material_summary.empty else pd.DataFrame({"Info": ["Ingen materialdata"]}), extra_sections=extra_sections)
    pdf_bytes = build_pdf_report(summary_dict, material_summary if not material_summary.empty else pd.DataFrame({"Info": ["Ingen materialdata"]}), extra_sections=extra_sections)

    ca, cb = st.columns(2)
    with ca:
        if docx_bytes is not None:
            st.download_button("Last ned rapport som Word", data=docx_bytes, file_name="byggtotal_rapport.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.info("Word-eksport er ikke tilgjengelig i dette miljøet.")
    with cb:
        if pdf_bytes is not None:
            st.download_button("Last ned rapport som PDF", data=pdf_bytes, file_name="byggtotal_rapport.pdf", mime="application/pdf")
        else:
            st.info("PDF-eksport er ikke tilgjengelig i dette miljøet.")

st.markdown("---")
st.markdown("**byggTotal**")
