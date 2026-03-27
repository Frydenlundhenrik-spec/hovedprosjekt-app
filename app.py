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
    import openpyxl  # noqa: F401
except Exception:
    openpyxl = None  # noqa: F841

try:
    import ifcopenshell
    import ifcopenshell.geom
    from ifcopenshell.util import element as ifc_element_util
except ImportError:
    ifcopenshell = None
    ifc_element_util = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
    )
except ImportError:
    SimpleDocTemplate = None


st.set_page_config(
    page_title="byggTotal – Mengder, kalkyle og IFC-analyse",
    page_icon="🏗️",
    layout="wide",
)

st.markdown("""
<style>
    .main {
        background-color: #f6f7fb;
    }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
    h1, h2, h3 {
        color: #1f2937;
    }
    .stMetric {
        background: white;
        padding: 18px;
        border-radius: 16px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
        border: 1px solid #e5e7eb;
    }
    div[data-testid="stDataFrame"] {
        background: white;
        padding: 10px;
        border-radius: 16px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }
    div[data-testid="stPlotlyChart"],
    div[data-testid="stPyplot"] {
        background: white;
        padding: 16px;
        border-radius: 16px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }
    section[data-testid="stSidebar"] {
        background-color: #eef2f7;
    }
    .stButton > button, .stDownloadButton > button {
        border-radius: 12px;
        border: none;
        background-color: #1f4e79;
        color: white;
        font-weight: 600;
        padding: 0.6rem 1rem;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #163a5a;
        color: white;
    }
    .custom-card {
        background: white;
        padding: 18px 20px;
        border-radius: 18px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 14px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    .custom-title {
        font-size: 15px;
        font-weight: 600;
        color: #6b7280;
        margin-bottom: 6px;
    }
    .custom-value {
        font-size: 30px;
        font-weight: 700;
        color: #111827;
    }
    .small-muted {
        color: #6b7280;
        font-size: 13px;
    }
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
    "IfcBeam",
    "IfcColumn",
    "IfcSlab",
    "IfcWall",
    "IfcWallStandardCase",
    "IfcRoof",
    "IfcMember",
    "IfcFooting",
]

MATERIAL_DATABASE = {
    "Stål": {
        "unit": "kg",
        "price": 47.0,
        "co2": 0.73,
        "density": 7850.0,
        "label": "Stål",
    },
    "Limtre": {
        "unit": "m3",
        "price": 28000.0,
        "co2": 100.0,
        "density": 460.0,
        "label": "Limtre",
    },
    "Massivtre": {
        "unit": "m3",
        "price": 30000.0,
        "co2": 110.0,
        "density": 500.0,
        "label": "Massivtre / CLT",
    },
    "Tre": {
        "unit": "m3",
        "price": 5000.0,
        "co2": 120.0,
        "density": 450.0,
        "label": "Tre",
    },
    "Betong_volum": {
        "unit": "m3",
        "price": 1800.0,
        "co2": 350.0,
        "density": 2400.0,
        "label": "Betong volum",
    },
    "Hulldekke": {
        "unit": "m2",
        "price": 1635.0,
        "co2": 84.56,
        "density": 2400.0,
        "label": "Hulldekke",
    },
    "Hulldekke_lavCO2": {
        "unit": "m2",
        "price": 1821.0,
        "co2": 64.86,
        "density": 2400.0,
        "label": "Hulldekke lavCO₂",
    },
    "Plasstøpt_betong": {
        "unit": "m2",
        "price": 2422.0,
        "co2": 69.59,
        "density": 2400.0,
        "label": "Plasstøpt betong",
    },
    "Plasstøpt_betong_lavCO2": {
        "unit": "m2",
        "price": 3015.0,
        "co2": 54.64,
        "density": 2400.0,
        "label": "Plasstøpt betong lavCO₂",
    },
    "Massivtre_vegg": {
        "unit": "m2",
        "price": 1337.0,
        "co2": 8.93,
        "density": 500.0,
        "label": "Massivtre vegg",
    },
    "Betong_vegg": {
        "unit": "m2",
        "price": 2910.0,
        "co2": 52.84,
        "density": 2400.0,
        "label": "Betong vegg",
    },
    "Betong_vegg_lavCO2": {
        "unit": "m2",
        "price": 3370.0,
        "co2": 43.54,
        "density": 2400.0,
        "label": "Betong vegg lavCO₂",
    },
    "Ukjent": {
        "unit": "m3",
        "price": 1000.0,
        "co2": 200.0,
        "density": 1000.0,
        "label": "Ukjent",
    },
}

NORSK_PRISBOK_DATABASE = {
    # ----------------------------
    # VEGGER – MASSIVTRE YTTERVEGG
    # ----------------------------
    "Massivtre_vegg_100": {
        "category": "Vegg",
        "unit": "m2",
        "price": 1575.0,
        "co2": 11.16,
        "ak": 96.06,
        "label": "Massive treelementer, yttervegg, t = 100 mm",
        "npb_code": "02.3.1.5.0110",
        "source": "Norsk Prisbok",
        "thickness_mm": 100,
    },
    "Massivtre_vegg_120": {
        "category": "Vegg",
        "unit": "m2",
        "price": 1879.0,
        "co2": 13.40,
        "ak": 114.58,
        "label": "Massive treelementer, yttervegg, t = 120 mm",
        "npb_code": "02.3.1.5.0120",
        "source": "Norsk Prisbok",
        "thickness_mm": 120,
    },
    "Massivtre_vegg_140": {
        "category": "Vegg",
        "unit": "m2",
        "price": 2177.0,
        "co2": 15.63,
        "ak": 132.70,
        "label": "Massive treelementer, yttervegg, t = 140 mm",
        "npb_code": "02.3.1.5.0130",
        "source": "Norsk Prisbok",
        "thickness_mm": 140,
    },
    "Massivtre_vegg_160": {
        "category": "Vegg",
        "unit": "m2",
        "price": 2466.0,
        "co2": 17.86,
        "ak": 150.32,
        "label": "Massive treelementer, yttervegg, t = 160 mm",
        "npb_code": "02.3.1.5.0140",
        "source": "Norsk Prisbok",
        "thickness_mm": 160,
    },
    "Massivtre_vegg_200": {
        "category": "Vegg",
        "unit": "m2",
        "price": 2897.0,
        "co2": 22.33,
        "ak": 160.08,
        "label": "Massive treelementer, yttervegg, t = 200 mm",
        "npb_code": "02.3.1.5.0160",
        "source": "Norsk Prisbok",
        "thickness_mm": 200,
    },
    "Massivtre_vegg_240": {
        "category": "Vegg",
        "unit": "m2",
        "price": 3225.0,
        "co2": 26.80,
        "ak": 178.20,
        "label": "Massive treelementer, yttervegg, t = 240 mm",
        "npb_code": "02.3.1.5.0170",
        "source": "Norsk Prisbok",
        "thickness_mm": 240,
    },
    "Massivtre_vegg_synlig_gran": {
        "category": "VeggTillegg",
        "unit": "m2",
        "price": 210.0,
        "co2": 0.0,
        "ak": 11.63,
        "label": "Merkostnad synlig overflate massivtreelement, innside yttervegg, gran",
        "npb_code": "02.3.1.5.0190",
        "source": "Norsk Prisbok",
        "thickness_mm": None,
    },

    # ----------------------------
    # VEGGER – PREFAB BETONG YTTERVEGG
    # ----------------------------
    "Betong_vegg_150": {
        "category": "Vegg",
        "unit": "m2",
        "price": 2566.0,
        "co2": 56.10,
        "ak": 141.78,
        "label": "Prefab betongyttervegg over mark, t = 150 mm",
        "npb_code": "02.3.B.001",
        "source": "Norsk Prisbok",
        "thickness_mm": 150,
    },
    "Betong_vegg_180": {
        "category": "Vegg",
        "unit": "m2",
        "price": 2885.0,
        "co2": 67.32,
        "ak": 159.39,
        "label": "Prefab betongyttervegg over mark, t = 180 mm",
        "npb_code": "02.3.B.002",
        "source": "Norsk Prisbok",
        "thickness_mm": 180,
    },
    "Betong_vegg_200": {
        "category": "Vegg",
        "unit": "m2",
        "price": 3100.0,
        "co2": 74.80,
        "ak": 171.28,
        "label": "Prefab betongyttervegg over mark, t = 200 mm",
        "npb_code": "02.3.B.003",
        "source": "Norsk Prisbok",
        "thickness_mm": 200,
    },

    # ----------------------------
    # DEKKER – PLASSTØPT BETONG
    # ----------------------------
    "Plasstopt_dekke_180": {
        "category": "Dekke",
        "unit": "m2",
        "price": 2285.0,
        "co2": 62.87,
        "ak": 126.26,
        "label": "Betongdekke, t = 180 mm. 130 kg armering pr m3 betong, B30",
        "npb_code": "02.5.B.001",
        "source": "Norsk Prisbok",
        "thickness_mm": 180,
    },
    "Plasstopt_dekke_200": {
        "category": "Dekke",
        "unit": "m2",
        "price": 2422.0,
        "co2": 69.59,
        "ak": 133.81,
        "label": "Betongdekke, t = 200 mm. 130 kg armering pr m3 betong, B30",
        "npb_code": "02.5.B.002",
        "source": "Norsk Prisbok",
        "thickness_mm": 200,
    },
    "Plasstopt_dekke_220": {
        "category": "Dekke",
        "unit": "m2",
        "price": 2559.0,
        "co2": 76.31,
        "ak": 141.37,
        "label": "Betongdekke, t = 220 mm. 130 kg armering pr m3 betong, B30",
        "npb_code": "02.5.B.003",
        "source": "Norsk Prisbok",
        "thickness_mm": 220,
    },
    "Plasstopt_dekke_250": {
        "category": "Dekke",
        "unit": "m2",
        "price": 2764.0,
        "co2": 86.40,
        "ak": 152.71,
        "label": "Betongdekke, t = 250 mm. 130 kg armering pr m3 betong, B30",
        "npb_code": "02.5.B.004",
        "source": "Norsk Prisbok",
        "thickness_mm": 250,
    },
    "Plasstopt_dekke_300": {
        "category": "Dekke",
        "unit": "m2",
        "price": 3157.0,
        "co2": 103.25,
        "ak": 174.41,
        "label": "Betongdekke, t = 300 mm. 130 kg armering pr m3 betong, B30",
        "npb_code": "02.5.B.005",
        "source": "Norsk Prisbok",
        "thickness_mm": 300,
    },
    "Plasstopt_dekke_350": {
        "category": "Dekke",
        "unit": "m2",
        "price": 3499.0,
        "co2": 120.06,
        "ak": 193.31,
        "label": "Betongdekke, t = 350 mm. 130 kg armering pr m3 betong, B30",
        "npb_code": "02.5.B.006",
        "source": "Norsk Prisbok",
        "thickness_mm": 350,
    },
    "Plasstopt_dekke_lavCO2": {
        "category": "Dekke",
        "unit": "m2",
        "price": 3015.0,
        "co2": 54.64,
        "ak": 166.58,
        "label": "Betongdekke med redusert klimagassutslipp, fasthetsklasse B30",
        "npb_code": "02.5.B.007",
        "source": "Norsk Prisbok",
        "thickness_mm": None,
    },

    # ----------------------------
    # DEKKER – HULLDEKKE
    # ----------------------------
    "Hulldekke_200": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1490.0,
        "co2": 65.06,
        "ak": 82.34,
        "label": "HD-element, t = 200 mm, med gysing og fuging, REI60",
        "npb_code": "02.5.C.001",
        "source": "Norsk Prisbok",
        "thickness_mm": 200,
    },
    "Hulldekke_220": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1577.0,
        "co2": 72.14,
        "ak": 87.13,
        "label": "HD-element, t = 220 mm, med gysing og fuging, REI120",
        "npb_code": "02.5.C.002",
        "source": "Norsk Prisbok",
        "thickness_mm": 220,
    },
    "Hulldekke_265": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1635.0,
        "co2": 84.56,
        "ak": 90.32,
        "label": "HD-element, t = 265 mm, med gysing og fuging, REI60",
        "npb_code": "02.5.C.003",
        "source": "Norsk Prisbok",
        "thickness_mm": 265,
    },
    "Hulldekke_265_lavCO2": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1821.0,
        "co2": 64.86,
        "ak": 100.64,
        "label": "HD-element, t = 265 mm, redusert klimagassutslipp, med gysing og fuging, REI60",
        "npb_code": "02.5.C.004",
        "source": "Norsk Prisbok",
        "thickness_mm": 265,
    },
    "Hulldekke_290": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1721.0,
        "co2": 86.94,
        "ak": 95.11,
        "label": "HD-element, t = 290 mm, med gysing og fuging, REI120",
        "npb_code": "02.5.C.005",
        "source": "Norsk Prisbok",
        "thickness_mm": 290,
    },
    "Hulldekke_320": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1779.0,
        "co2": 93.80,
        "ak": 98.31,
        "label": "HD-element, t = 320 mm, med gysing og fuging, REI90",
        "npb_code": "02.5.C.006",
        "source": "Norsk Prisbok",
        "thickness_mm": 320,
    },
    "Hulldekke_320_lavCO2": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1970.0,
        "co2": 71.62,
        "ak": 108.82,
        "label": "HD-element, t = 320 mm, redusert klimagassutslipp, med gysing og fuging, REI90",
        "npb_code": "02.5.C.007",
        "source": "Norsk Prisbok",
        "thickness_mm": 320,
    },
    "Hulldekke_340": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1808.0,
        "co2": 95.40,
        "ak": 99.90,
        "label": "HD-element, t = 340 mm, med gysing og fuging, REI90",
        "npb_code": "02.5.C.008",
        "source": "Norsk Prisbok",
        "thickness_mm": 340,
    },
    "Hulldekke_400": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1837.0,
        "co2": 100.48,
        "ak": 101.50,
        "label": "HD-element, t = 400 mm, med gysing og fuging, REI90",
        "npb_code": "02.5.C.009",
        "source": "Norsk Prisbok",
        "thickness_mm": 400,
    },
    "Hulldekke_420": {
        "category": "Dekke",
        "unit": "m2",
        "price": 1924.0,
        "co2": 107.76,
        "ak": 106.29,
        "label": "HD-element, t = 420 mm, med gysing og fuging, REI120",
        "npb_code": "02.5.C.010",
        "source": "Norsk Prisbok",
        "thickness_mm": 420,
    },
    "Hulldekke_500": {
        "category": "Dekke",
        "unit": "m2",
        "price": 2110.0,
        "co2": 127.88,
        "ak": 116.60,
        "label": "HD-element, t = 500 mm, med gysing og fuging, REI120",
        "npb_code": "02.5.C.011",
        "source": "Norsk Prisbok",
        "thickness_mm": 500,
    },

    # ----------------------------
    # DEKKER – MASSIVTRE BÆRENDE DEKKE
    # ----------------------------
    "Massivtre_dekke_160": {
        "category": "Dekke",
        "unit": "m2",
        "price": 2570.0,
        "co2": 17.86,
        "ak": 142.01,
        "label": "Massive treelementer i dekker, bærende, t = 160 mm",
        "npb_code": "02.5.C.031",
        "source": "Norsk Prisbok",
        "thickness_mm": 160,
    },
    "Massivtre_dekke_180": {
        "category": "Dekke",
        "unit": "m2",
        "price": 2798.0,
        "co2": 20.10,
        "ak": 154.61,
        "label": "Massive treelementer i dekker, bærende, t = 180 mm",
        "npb_code": "02.5.C.032",
        "source": "Norsk Prisbok",
        "thickness_mm": 180,
    },
    "Massivtre_dekke_200": {
        "category": "Dekke",
        "unit": "m2",
        "price": 3018.0,
        "co2": 22.33,
        "ak": 166.77,
        "label": "Massive treelementer i dekker, bærende, t = 200 mm",
        "npb_code": "02.5.C.033",
        "source": "Norsk Prisbok",
        "thickness_mm": 200,
    },
    "Massivtre_dekke_220": {
        "category": "Dekke",
        "unit": "m2",
        "price": 3161.0,
        "co2": 24.56,
        "ak": 174.64,
        "label": "Massive treelementer i dekker, bærende, t = 220 mm",
        "npb_code": "02.5.C.034",
        "source": "Norsk Prisbok",
        "thickness_mm": 220,
    },
    "Massivtre_dekke_240": {
        "category": "Dekke",
        "unit": "m2",
        "price": 3419.0,
        "co2": 26.80,
        "ak": 188.88,
        "label": "Massive treelementer i dekker, bærende, t = 240 mm",
        "npb_code": "02.5.C.035",
        "source": "Norsk Prisbok",
        "thickness_mm": 240,
    },
    "Massivtre_dekke_260": {
        "category": "Dekke",
        "unit": "m2",
        "price": 3700.0,
        "co2": 29.03,
        "ak": 204.42,
        "label": "Massive treelementer i dekker, bærende, t = 260 mm",
        "npb_code": "02.5.C.036",
        "source": "Norsk Prisbok",
        "thickness_mm": 260,
    },
    "Massivtre_dekke_280": {
        "category": "Dekke",
        "unit": "m2",
        "price": 3972.0,
        "co2": 31.26,
        "ak": 219.46,
        "label": "Massive treelementer i dekker, bærende, t = 280 mm",
        "npb_code": "02.5.C.037",
        "source": "Norsk Prisbok",
        "thickness_mm": 280,
    },
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
    "Limtre": [
        "90x315",
        "90x405",
        "115x315",
        "115x360",
        "115x405",
        "140x315",
        "140x360",
        "140x405",
        "140x450",
        "165x315",
        "165x360",
        "165x405",
        "190x405",
        "190x450",
        "215x405",
        "215x450",
    ],
    "Massivtre": [
        "100x300",
        "120x300",
        "120x400",
        "140x400",
        "160x400",
        "200x400",
    ],
    "Stål": [
        "KFHUP 120x120x8",
        "KFHUP 140x140x10",
        "KFHUP 160x160x10",
        "KFHUP 180x180x12.5",
        "KFHUP 200x200x12.5",
        "KFHUP 220x220x12.5",
    ],
    "Betong": [
        "200x200",
        "250x250",
        "300x300",
        "350x350",
        "400x400",
    ],
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
    st.markdown(
        f"""
        <div class="custom-card">
            <div class="custom-title">{title}</div>
            <div class="custom-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def file_hash(file_bytes: bytes) -> str:
    return hashlib.md5(file_bytes).hexdigest()


@st.cache_data(show_spinner=False)
def load_sheet_df(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl")


def parse_profile(profile: str):
    text = str(profile or "")
    material = "Ukjent"

    if "stål" in text.lower() or "steel" in text.lower():
        material = "Stål"
    elif "limtre" in text.lower() or "glulam" in text.lower():
        material = "Limtre"
    elif "massivtre" in text.lower() or "clt" in text.lower():
        material = "Massivtre"

    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+[\.,]?\d*", text)]
    area_m2 = None
    width_mm = height_mm = thickness_mm = None

    if material == "Stål" and len(nums) >= 3:
        width_mm, height_mm, thickness_mm = nums[-3], nums[-2], nums[-1]
        inner_w = max(width_mm - 2 * thickness_mm, 0)
        inner_h = max(height_mm - 2 * thickness_mm, 0)
        area_mm2 = (width_mm * height_mm) - (inner_w * inner_h)
        area_m2 = area_mm2 / 1_000_000
    elif material in ["Limtre", "Massivtre"] and len(nums) >= 2:
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


def parse_rect_profile_mm(profile_text: str):
    text = str(profile_text or "")
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+[\.,]?\d*", text)]
    if len(nums) >= 2:
        return nums[-2], nums[-1]
    return None, None


def parse_profile_area_from_text(profile_text: str, material_hint: str = "") -> float:
    text = str(profile_text or "")
    material_hint = str(material_hint or "")
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+[\.,]?\d*", text)]

    if len(nums) < 2:
        return math.nan

    lower_text = text.lower()
    material_guess = classify_material(material_hint if material_hint else text)

    if material_guess == "Stål" or any(x in lower_text for x in ["kfh", "rhs", "shs", "hup"]):
        if len(nums) >= 3:
            width_mm, height_mm, thickness_mm = nums[-3], nums[-2], nums[-1]
            inner_w = max(width_mm - 2 * thickness_mm, 0)
            inner_h = max(height_mm - 2 * thickness_mm, 0)
            area_mm2 = (width_mm * height_mm) - (inner_w * inner_h)
            return area_mm2 / 1_000_000
        return math.nan

    width_mm, height_mm = nums[-2], nums[-1]
    area_mm2 = width_mm * height_mm
    return area_mm2 / 1_000_000


def get_swap_profile_options(df: pd.DataFrame, to_material: str):
    model_profiles = []

    if "Material / Tverrsnitt" in df.columns and "materiale" in df.columns:
        model_profiles = sorted(
            df.loc[df["materiale"] == to_material, "Material / Tverrsnitt"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

    library_profiles = PROFILE_LIBRARY.get(to_material, [])

    combined = []
    for p in library_profiles + model_profiles:
        if p and p not in combined:
            combined.append(p)

    return combined


def get_swap_target_options(selected_type: str, from_material: str):
    selected_type = str(selected_type or "")

    options = []

    if selected_type == "Søyle":
        options = ["Stål", "Limtre", "Betong"]
    elif selected_type == "Bjelke":
        options = ["Stål", "Limtre", "Betong"]
    elif selected_type == "Vegg":
        options = [
            "Betong_vegg_150",
            "Betong_vegg_180",
            "Betong_vegg_200",
            "Massivtre_vegg_100",
            "Massivtre_vegg_120",
            "Massivtre_vegg_140",
            "Massivtre_vegg_160",
            "Massivtre_vegg_200",
            "Massivtre_vegg_240",
        ]
    elif selected_type == "Dekke":
        options = [
            "Plasstopt_dekke_180",
            "Plasstopt_dekke_200",
            "Plasstopt_dekke_220",
            "Plasstopt_dekke_250",
            "Plasstopt_dekke_300",
            "Plasstopt_dekke_350",
            "Plasstopt_dekke_lavCO2",
            "Hulldekke_200",
            "Hulldekke_220",
            "Hulldekke_265",
            "Hulldekke_265_lavCO2",
            "Hulldekke_290",
            "Hulldekke_320",
            "Hulldekke_320_lavCO2",
            "Hulldekke_340",
            "Hulldekke_400",
            "Hulldekke_420",
            "Hulldekke_500",
            "Massivtre_dekke_160",
            "Massivtre_dekke_180",
            "Massivtre_dekke_200",
            "Massivtre_dekke_220",
            "Massivtre_dekke_240",
            "Massivtre_dekke_260",
            "Massivtre_dekke_280",
        ]
    else:
        options = ["Stål", "Limtre", "Betong"]

    return options


def format_swap_target_option(option_key: str) -> str:
    if option_key in NORSK_PRISBOK_DATABASE:
        item = NORSK_PRISBOK_DATABASE[option_key]
        return f"{item['label']} ({item['npb_code']})"
    if option_key in MATERIAL_DATABASE:
        return MATERIAL_DATABASE[option_key]["label"]
    return option_key


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
        "Stål": "#4F81BD",
        "Betong": "#A6A6A6",
        "Limtre": "#C58C4B",
        "Massivtre": "#8CBF3F",
        "Tre": "#B97A57",
        "Ukjent": "#D9D9D9",
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
    product_key = detect_product_key(row, deck_variant, concrete_variant, wall_variant)
    product = MATERIAL_DATABASE.get(product_key, MATERIAL_DATABASE["Ukjent"])
    qty = get_quantity_for_product(row, product_key)
    return qty * product["price"]


def co2_for_row(row, deck_variant, concrete_variant, wall_variant, use_epd=True):
    product_key = detect_product_key(row, deck_variant, concrete_variant, wall_variant)
    qty = get_quantity_for_product(row, product_key)

    if use_epd and product_key in EPD_DATABASE:
        return qty * EPD_DATABASE[product_key]["co2"]

    product = MATERIAL_DATABASE.get(product_key, MATERIAL_DATABASE["Ukjent"])
    return qty * product["co2"]


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


def get_swap_material_defaults(to_material: str):
    if to_material == "Stål":
        return {
            "density": MATERIAL_DATABASE["Stål"]["density"],
            "price": MATERIAL_DATABASE["Stål"]["price"],
            "price_unit": "kg",
            "co2": EPD_DATABASE["Stål"]["co2"],
            "label": "Stål",
        }
    if to_material == "Limtre":
        return {
            "density": MATERIAL_DATABASE["Limtre"]["density"],
            "price": MATERIAL_DATABASE["Limtre"]["price"],
            "price_unit": "m3",
            "co2": EPD_DATABASE["Limtre"]["co2"],
            "label": "Limtre",
        }
    if to_material == "Massivtre":
        return {
            "density": MATERIAL_DATABASE["Massivtre"]["density"],
            "price": MATERIAL_DATABASE["Massivtre"]["price"],
            "price_unit": "m3",
            "co2": EPD_DATABASE["Massivtre"]["co2"],
            "label": "Massivtre",
        }
    if to_material == "Betong":
        return {
            "density": MATERIAL_DATABASE["Betong_volum"]["density"],
            "price": MATERIAL_DATABASE["Betong_volum"]["price"],
            "price_unit": "m3",
            "co2": EPD_DATABASE["Betong_volum"]["co2"],
            "label": "Betong",
        }
    return {
        "density": 0.0,
        "price": 0.0,
        "price_unit": "",
        "co2": 0.0,
        "label": "Ukjent",
    }


def get_swap_target_defaults(target_key: str):
    if target_key in NORSK_PRISBOK_DATABASE:
        db = NORSK_PRISBOK_DATABASE[target_key]
        return {
            "density": 0.0,
            "price": db["price"],
            "price_unit": db["unit"],
            "co2": db["co2"],
            "label": db["label"],
            "target_key": target_key,
            "source": db["source"],
            "npb_code": db["npb_code"],
            "ak": db.get("ak", 0.0),
            "thickness_mm": db.get("thickness_mm"),
        }

    if target_key in MATERIAL_DATABASE:
        db = MATERIAL_DATABASE[target_key]
        co2_val = EPD_DATABASE.get(target_key, {}).get("co2", db.get("co2", 0.0))
        return {
            "density": db.get("density", 0.0),
            "price": db.get("price", 0.0),
            "price_unit": db.get("unit", ""),
            "co2": co2_val,
            "label": db.get("label", target_key),
            "target_key": target_key,
            "source": "Materialdatabase",
            "npb_code": "",
            "ak": 0.0,
            "thickness_mm": None,
        }

    if target_key in ["Stål", "Limtre", "Massivtre", "Betong"]:
        material_defaults = get_swap_material_defaults(target_key)
        return {
            "density": material_defaults.get("density", 0.0),
            "price": material_defaults.get("price", 0.0),
            "price_unit": material_defaults.get("price_unit", ""),
            "co2": material_defaults.get("co2", 0.0),
            "label": material_defaults.get("label", target_key),
            "target_key": target_key,
            "source": "Materialdatabase",
            "npb_code": "",
            "ak": 0.0,
            "thickness_mm": None,
        }

    return {
        "density": 0.0,
        "price": 0.0,
        "price_unit": "",
        "co2": 0.0,
        "label": target_key,
        "target_key": target_key,
        "source": "Ukjent",
        "npb_code": "",
        "ak": 0.0,
        "thickness_mm": None,
    }


def is_area_based_swap_target(target_key: str) -> bool:
    if target_key in NORSK_PRISBOK_DATABASE:
        return NORSK_PRISBOK_DATABASE[target_key]["unit"] == "m2"

    return target_key in [
        "Betong_vegg",
        "Betong_vegg_lavCO2",
        "Hulldekke",
        "Hulldekke_lavCO2",
        "Plasstøpt_betong",
        "Plasstøpt_betong_lavCO2",
        "Massivtre_vegg",
    ]


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
    )
    merged = merged.drop(columns=[c for c in ["Navn"] if c in merged.columns])

    profile_df = merged["Material / Tverrsnitt"].apply(parse_profile).apply(pd.Series)
    merged = pd.concat([merged, profile_df], axis=1)

    if "Lengde [m]" in merged.columns:
        merged["Lengde [m]"] = pd.to_numeric(merged["Lengde [m]"], errors="coerce")
    else:
        merged["Lengde [m]"] = math.nan

    if "Areal [m2]" in merged.columns:
        merged["Areal [m2]"] = pd.to_numeric(merged["Areal [m2]"], errors="coerce")
    else:
        merged["Areal [m2]"] = math.nan

    if "Volum [m3]" in merged.columns:
        merged["Volum [m3]"] = pd.to_numeric(merged["Volum [m3]"], errors="coerce")
    else:
        merged["Volum [m3]"] = merged["Lengde [m]"] * merged["areal_m2"]

    def calc_weight(row):
        if row["materiale"] == "Stål":
            return safe_num(row["Volum [m3]"]) * STEEL_DENSITY
        if row["materiale"] == "Limtre":
            return safe_num(row["Volum [m3]"]) * GLULAM_DENSITY
        if row["materiale"] == "Massivtre":
            return safe_num(row["Volum [m3]"]) * CLT_DENSITY
        if row["materiale"] == "Betong":
            return safe_num(row["Volum [m3]"]) * CONCRETE_DENSITY
        if row["materiale"] == "Tre":
            return safe_num(row["Volum [m3]"]) * TIMBER_DENSITY
        return math.nan

    merged["Vekt [kg]"] = merged.apply(calc_weight, axis=1)
    merged["Mengdegrunnlag"] = merged.apply(
        lambda row: "Excel"
        if any(pd.notna(row.get(c)) and safe_num(row.get(c)) > 0 for c in ["Lengde [m]", "Areal [m2]", "Volum [m3]"])
        else "Manglende mengder",
        axis=1,
    )
    merged["Endret IFC"] = False

    knutepunkter.columns = [str(c).strip() for c in knutepunkter.columns]
    return merged, knutepunkter, forside


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
                if getattr(prop, "Name", "") != prop_name:
                    continue
                nominal = getattr(prop, "NominalValue", None)
                if nominal is None:
                    return None
                if hasattr(nominal, "wrappedValue"):
                    return nominal.wrappedValue
                return nominal
    except Exception:
        pass
    return None


def is_ifc_element_changed(element) -> bool:
    value = get_property_from_pset(element, BYGGTOTAL_PSET_NAME, BYGGTOTAL_CHANGED_PROP)
    return bool(value) if value is not None else False


def estimate_dimensions_from_mesh(verts):
    if not verts:
        return None

    x = verts[0::3]
    y = verts[1::3]
    z = verts[2::3]

    if not x or not y or not z:
        return None

    dx = max(x) - min(x)
    dy = max(y) - min(y)
    dz = max(z) - min(z)

    dims = sorted([abs(dx), abs(dy), abs(dz)], reverse=True)
    return dims


def estimate_quantities_from_geometry(element, settings):
    try:
        shape = ifcopenshell.geom.create_shape(settings, element)
        geom = shape.geometry
        verts = geom.verts
        dims = estimate_dimensions_from_mesh(verts)
        if not dims:
            return {"length": None, "area": None, "volume": None, "weight": None, "method": None}

        d1, d2, d3 = dims
        volume = d1 * d2 * d3
        side_area = d1 * d2
        length = d1

        return {
            "length": length if length > 0 else None,
            "area": side_area if side_area > 0 else None,
            "volume": volume if volume > 0 else None,
            "weight": None,
            "method": "Geometriestimat",
        }
    except Exception:
        return {"length": None, "area": None, "volume": None, "weight": None, "method": None}


def build_dataset_from_ifc(ifc_bytes: bytes):
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert. Kjør: py -m pip install ifcopenshell")

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

                if (
                    (pd.isna(length_m) or safe_num(length_m) == 0)
                    and (pd.isna(area_m2) or safe_num(area_m2) == 0)
                    and (pd.isna(volume_m3) or safe_num(volume_m3) == 0)
                ):
                    geo_q = estimate_quantities_from_geometry(el, settings)
                    length_m = pd.to_numeric(geo_q["length"], errors="coerce")
                    area_m2 = pd.to_numeric(geo_q["area"], errors="coerce")
                    volume_m3 = pd.to_numeric(geo_q["volume"], errors="coerce")
                    quantity_method = geo_q["method"] or "Ikke funnet"

                if (pd.isna(area_m2) or safe_num(area_m2) == 0) and pd.notna(volume_m3) and pd.notna(length_m) and safe_num(length_m) > 0:
                    area_m2 = safe_num(volume_m3) / safe_num(length_m)

                if pd.notna(weight_kg) and safe_num(weight_kg) > 0:
                    vekt_kg = weight_kg
                elif materiale == "Stål" and pd.notna(volume_m3):
                    vekt_kg = safe_num(volume_m3) * STEEL_DENSITY
                elif materiale == "Limtre" and pd.notna(volume_m3):
                    vekt_kg = safe_num(volume_m3) * GLULAM_DENSITY
                elif materiale == "Massivtre" and pd.notna(volume_m3):
                    vekt_kg = safe_num(volume_m3) * CLT_DENSITY
                elif materiale == "Betong" and pd.notna(volume_m3):
                    vekt_kg = safe_num(volume_m3) * CONCRETE_DENSITY
                elif materiale == "Tre" and pd.notna(volume_m3):
                    vekt_kg = safe_num(volume_m3) * TIMBER_DENSITY
                else:
                    vekt_kg = math.nan

                rows.append(
                    {
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
                    }
                )

        data = pd.DataFrame(rows)
        nodes = pd.DataFrame()
        forside = pd.DataFrame(
            [
                ["Kilde", "IFC"],
                ["Antall elementer", len(data)],
                ["Filtype", "IFC"],
            ],
            columns=["Parameter", "Verdi"],
        )

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

                        # noqa: E701
                    x = verts[0::3]
                    y = verts[1::3]
                    z = verts[2::3]

                    i_idx = faces[0::3]
                    j_idx = faces[1::3]
                    k_idx = faces[2::3]

                    material_raw = get_ifc_material_name(el)
                    materiale = classify_material(material_raw)
                    changed_flag = is_ifc_element_changed(el)

                    meshes.append(
                        {
                            "global_id": gid,
                            "name": getattr(el, "Name", "") or gid or "Ukjent",
                            "type": map_ifc_type(type_name),
                            "ifc_type": type_name,
                            "materiale": materiale,
                            "changed": changed_flag,
                            "x": x,
                            "y": y,
                            "z": z,
                            "i": i_idx,
                            "j": j_idx,
                            "k": k_idx,
                        }
                    )

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


def build_ifc_3d_figure(meshes):
    fig = go.Figure()

    for mesh in meshes:
        color = material_color(mesh["materiale"], mesh.get("changed", False))
        label_material = f"{mesh['materiale']} (endret)" if mesh.get("changed", False) else mesh["materiale"]

        fig.add_trace(
            go.Mesh3d(
                x=mesh["x"],
                y=mesh["y"],
                z=mesh["z"],
                i=mesh["i"],
                j=mesh["j"],
                k=mesh["k"],
                color=color,
                opacity=0.95,
                flatshading=True,
                name=f"{mesh['type']} – {label_material}",
                hovertext=(
                    f"Navn: {mesh['name']}<br>"
                    f"Type: {mesh['type']}<br>"
                    f"IFC-type: {mesh['ifc_type']}<br>"
                    f"Materiale: {mesh['materiale']}<br>"
                    f"Endret i IFC: {'Ja' if mesh.get('changed', False) else 'Nei'}<br>"
                    f"GlobalId: {mesh['global_id']}"
                ),
                hoverinfo="text",
                showscale=False,
            )
        )

    fig.update_layout(
        margin=dict(l=0, r=0, t=20, b=0),
        scene=dict(
            xaxis_title="X",
            yaxis_title="Y",
            zaxis_title="Z",
            aspectmode="data",
            bgcolor="rgba(0,0,0,0)",
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
        ),
        height=760,
    )
    return fig


def build_ifc_3d_preview_figure(meshes, preview_ids=None, show_only_preview=False, preview_material=None):
    preview_ids = set(preview_ids or [])
    fig = go.Figure()

    for mesh in meshes:
        is_preview = mesh["global_id"] in preview_ids

        if show_only_preview and not is_preview:
            continue

        if is_preview:
            color = "#ff66cc"
            opacity = 1.0
            preview_text = "Ja"
            display_material = preview_material if preview_material else f"{mesh['materiale']} → ny"
        else:
            color = material_color(mesh["materiale"], mesh.get("changed", False))
            opacity = 0.12
            preview_text = "Nei"
            display_material = mesh["materiale"]

        fig.add_trace(
            go.Mesh3d(
                x=mesh["x"],
                y=mesh["y"],
                z=mesh["z"],
                i=mesh["i"],
                j=mesh["j"],
                k=mesh["k"],
                color=color,
                opacity=opacity,
                flatshading=True,
                name=f"{mesh['type']} – {display_material}",
                hovertext=(
                    f"Navn: {mesh['name']}<br>"
                    f"Type: {mesh['type']}<br>"
                    f"IFC-type: {mesh['ifc_type']}<br>"
                    f"Eksisterende materiale: {mesh['materiale']}<br>"
                    f"Forhåndsvises som byttet: {preview_text}<br>"
                    f"GlobalId: {mesh['global_id']}"
                ),
                hoverinfo="text",
                showscale=False,
            )
        )

    fig.update_layout(
        margin=dict(l=0, r=0, t=20, b=0),
        scene=dict(
            xaxis_title="X",
            yaxis_title="Y",
            zaxis_title="Z",
            aspectmode="data",
            bgcolor="rgba(0,0,0,0)",
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
        ),
        height=760,
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


def set_or_create_pset_text_property(model, element, pset_name: str, prop_name: str, value: str):
    owner_history = get_owner_history(model)

    existing_pset = None
    for rel in getattr(element, "IsDefinedBy", []) or []:
        pdef = getattr(rel, "RelatingPropertyDefinition", None)
        if pdef and pdef.is_a("IfcPropertySet") and getattr(pdef, "Name", "") == pset_name:
            existing_pset = pdef
            break

    if existing_pset is None:
        prop = model.create_entity(
            "IfcPropertySingleValue",
            Name=prop_name,
            Description=None,
            NominalValue=_make_ifc_text(model, value),
            Unit=None,
        )
        pset = model.create_entity(
            "IfcPropertySet",
            GlobalId=ifcopenshell.guid.new(),
            OwnerHistory=owner_history,
            Name=pset_name,
            Description=None,
            HasProperties=[prop],
        )
        model.create_entity(
            "IfcRelDefinesByProperties",
            GlobalId=ifcopenshell.guid.new(),
            OwnerHistory=owner_history,
            Name=f"{pset_name} relation",
            Description=None,
            RelatedObjects=[element],
            RelatingPropertyDefinition=pset,
        )
        return

    props = list(getattr(existing_pset, "HasProperties", []) or [])
    for prop in props:
        if getattr(prop, "Name", "") == prop_name:
            prop.NominalValue = _make_ifc_text(model, value)
            return

    props.append(
        model.create_entity(
            "IfcPropertySingleValue",
            Name=prop_name,
            Description=None,
            NominalValue=_make_ifc_text(model, value),
            Unit=None,
        )
    )
    existing_pset.HasProperties = props


def set_or_create_pset_bool_property(model, element, pset_name: str, prop_name: str, value: bool):
    owner_history = get_owner_history(model)

    existing_pset = None
    for rel in getattr(element, "IsDefinedBy", []) or []:
        pdef = getattr(rel, "RelatingPropertyDefinition", None)
        if pdef and pdef.is_a("IfcPropertySet") and getattr(pdef, "Name", "") == pset_name:
            existing_pset = pdef
            break

    if existing_pset is None:
        prop = model.create_entity(
            "IfcPropertySingleValue",
            Name=prop_name,
            Description=None,
            NominalValue=_make_ifc_boolean(model, value),
            Unit=None,
        )
        pset = model.create_entity(
            "IfcPropertySet",
            GlobalId=ifcopenshell.guid.new(),
            OwnerHistory=owner_history,
            Name=pset_name,
            Description=None,
            HasProperties=[prop],
        )
        model.create_entity(
            "IfcRelDefinesByProperties",
            GlobalId=ifcopenshell.guid.new(),
            OwnerHistory=owner_history,
            Name=f"{pset_name} relation",
            Description=None,
            RelatedObjects=[element],
            RelatingPropertyDefinition=pset,
        )
        return

    props = list(getattr(existing_pset, "HasProperties", []) or [])
    for prop in props:
        if getattr(prop, "Name", "") == prop_name:
            prop.NominalValue = _make_ifc_boolean(model, value)
            return

    props.append(
        model.create_entity(
            "IfcPropertySingleValue",
            Name=prop_name,
            Description=None,
            NominalValue=_make_ifc_boolean(model, value),
            Unit=None,
        )
    )
    existing_pset.HasProperties = props


def export_ifc_material_swap(
    ifc_bytes: bytes,
    source_df: pd.DataFrame,
    selected_type: str,
    from_material: str,
    target_key: str,
    new_profile_text: str = "",
):
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert.")

    temp_in = None
    temp_out = None

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_in_file:
            tmp_in_file.write(ifc_bytes)
            temp_in = tmp_in_file.name

        model = ifcopenshell.open(temp_in)

        matched = source_df[
            (source_df["Type"] == selected_type) &
            (source_df["materiale"] == from_material)
        ].copy()

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
                        el.ObjectType = f"{target_key}"
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

                set_or_create_pset_bool_property(
                    model, el, BYGGTOTAL_PSET_NAME, BYGGTOTAL_CHANGED_PROP, True
                )
                set_or_create_pset_text_property(
                    model, el, BYGGTOTAL_PSET_NAME, BYGGTOTAL_OLD_MATERIAL_PROP, str(old_material)
                )
                set_or_create_pset_text_property(
                    model, el, BYGGTOTAL_PSET_NAME, BYGGTOTAL_NEW_MATERIAL_PROP, str(target_label)
                )
                set_or_create_pset_text_property(
                    model, el, BYGGTOTAL_PSET_NAME, BYGGTOTAL_PROFILE_PROP, str(new_profile_text or "")
                )

                changed_rows.append(
                    {
                        "IFC GlobalId": gid,
                        "Navn": old_name,
                        "Type": map_ifc_type(type_name),
                        "Gammelt materiale": old_material,
                        "Nytt materiale": target_label,
                        "Norsk Prisbok-kode": defaults.get("npb_code", ""),
                        "Nytt tverrsnitt": new_profile_text,
                        "Gammel ObjectType": old_object_type,
                        "Ny ObjectType": getattr(el, "ObjectType", "") or "",
                    }
                )

        if not changed_rows:
            return None, pd.DataFrame()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp_out_file:
            temp_out = tmp_out_file.name

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


def calculate_material_swap(
    source_df: pd.DataFrame,
    selected_type: str,
    from_material: str,
    target_key: str,
    new_profile_text: str,
):
    matched = source_df[
        (source_df["Type"] == selected_type) &
        (source_df["materiale"] == from_material)
    ].copy()

    if matched.empty:
        return matched

    defaults = get_swap_target_defaults(target_key)
    target_density = defaults["density"]
    target_price = defaults["price"]
    target_price_unit = defaults["price_unit"]
    target_co2 = defaults["co2"]

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
        matched["Ny kostnad [kr]"] = matched["Areal [m2]"].fillna(0) * target_price
        matched["Ny CO2 [kgCO2e]"] = matched["Areal [m2]"].fillna(0) * target_co2
        matched["Byttemetode"] = "Areal × Norsk Prisbok-post"
        matched["Nytt tverrsnittsareal [m2]"] = math.nan
    else:
        new_area_m2 = parse_profile_area_from_text(new_profile_text, target_key)

        matched["Nytt volum [m3]"] = matched.apply(
            lambda row: safe_num(row["Byttelengde [m]"]) * new_area_m2
            if pd.notna(row["Byttelengde [m]"]) and pd.notna(new_area_m2) and new_area_m2 > 0
            else safe_num(row["Gammelt volum [m3]"]),
            axis=1,
        )

        matched["Ny vekt [kg]"] = matched["Nytt volum [m3]"] * target_density

        if target_price_unit == "kg":
            matched["Ny kostnad [kr]"] = matched["Ny vekt [kg]"] * target_price
            matched["Ny CO2 [kgCO2e]"] = matched["Ny vekt [kg]"] * target_co2
        elif target_price_unit == "m3":
            matched["Ny kostnad [kr]"] = matched["Nytt volum [m3]"] * target_price
            matched["Ny CO2 [kgCO2e]"] = matched["Nytt volum [m3]"] * target_co2
        elif target_price_unit == "m2":
            matched["Ny kostnad [kr]"] = matched["Areal [m2]"].fillna(0) * target_price
            matched["Ny CO2 [kgCO2e]"] = matched["Areal [m2]"].fillna(0) * target_co2
        else:
            matched["Ny kostnad [kr]"] = math.nan
            matched["Ny CO2 [kgCO2e]"] = math.nan

        matched["Byttemetode"] = matched.apply(
            lambda row: "Utledet lengde × nytt tverrsnitt"
            if pd.notna(row["Byttelengde [m]"]) and pd.notna(new_area_m2) and new_area_m2 > 0
            else "Fallback til eksisterende volum",
            axis=1,
        )
        matched["Nytt tverrsnittsareal [m2]"] = new_area_m2

    matched["Kostnadsendring [kr]"] = matched["Ny kostnad [kr]"] - matched["Gammel kostnad [kr]"]
    matched["Vektendring [kg]"] = matched["Ny vekt [kg]"] - matched["Gammel vekt [kg]"]
    matched["CO2-endring [kgCO2e]"] = matched["Ny CO2 [kgCO2e]"] - matched["Gammel CO2 [kgCO2e]"]

    if defaults.get("source") == "Norsk Prisbok":
        matched["Prisgrunnlag"] = (
            f"{defaults['label']} ({target_price_unit}) fra Norsk Prisbok, post {defaults.get('npb_code')}"
        )
    else:
        matched["Prisgrunnlag"] = f"{defaults['label']} ({target_price_unit}) fra materialdatabase"

    matched["Tetthet brukt [kg/m3]"] = target_density
    matched["CO2-faktor brukt"] = target_co2
    matched["Norsk Prisbok-kode"] = defaults.get("npb_code", "")
    matched["ÅK/enh"] = defaults.get("ak", 0.0)

    return matched


def build_docx_report(summary_dict, material_summary, swap_summary=None):
    if Document is None:
        return None

    doc = Document()
    doc.add_heading("byggTotal – Prosjektrapport", 0)

    p = doc.add_paragraph()
    p.add_run("Generert: ").bold = True
    p.add_run(datetime.now().strftime("%d.%m.%Y %H:%M"))

    doc.add_heading("Prosjektoversikt", level=1)
    for key, value in summary_dict.items():
        doc.add_paragraph(f"{key}: {value}")

    doc.add_heading("Materialoversikt", level=1)
    table = doc.add_table(rows=1, cols=len(material_summary.columns))
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    for i, col in enumerate(material_summary.columns):
        hdr[i].text = str(col)

    for _, row in material_summary.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row):
            if isinstance(val, float):
                cells[i].text = f"{val:,.2f}".replace(",", " ")
            else:
                cells[i].text = str(val)

    if swap_summary is not None and not swap_summary.empty:
        doc.add_heading("Materialbytte", level=1)
        table2 = doc.add_table(rows=1, cols=len(swap_summary.columns))
        table2.style = "Table Grid"
        hdr2 = table2.rows[0].cells
        for i, col in enumerate(swap_summary.columns):
            hdr2[i].text = str(col)

        for _, row in swap_summary.iterrows():
            cells = table2.add_row().cells
            for i, val in enumerate(row):
                if isinstance(val, float):
                    cells[i].text = f"{val:,.2f}".replace(",", " ")
                else:
                    cells[i].text = str(val)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()


def build_pdf_report(summary_dict, material_summary, swap_summary=None):
    if SimpleDocTemplate is None:
        return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("byggTotal – Prosjektrapport", styles["Title"]))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Generert: {datetime.now().strftime('%d.%m.%Y %H:%M')}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Prosjektoversikt", styles["Heading2"]))
    summary_table_data = [["Parameter", "Verdi"]]
    for k, v in summary_dict.items():
        summary_table_data.append([str(k), str(v)])

    t1 = Table(summary_table_data, hAlign="LEFT")
    t1.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
            ]
        )
    )
    elements.append(t1)
    elements.append(Spacer(1, 16))

    elements.append(Paragraph("Materialoversikt", styles["Heading2"]))
    material_data = [list(material_summary.columns)] + material_summary.astype(str).values.tolist()
    t2 = Table(material_data, hAlign="LEFT")
    t2.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
            ]
        )
    )
    elements.append(t2)
    elements.append(Spacer(1, 16))

    if swap_summary is not None and not swap_summary.empty:
        elements.append(Paragraph("Materialbytte", styles["Heading2"]))
        swap_data = [list(swap_summary.columns)] + swap_summary.astype(str).values.tolist()
        t3 = Table(swap_data, hAlign="LEFT")
        t3.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
                ]
            )
        )
        elements.append(t3)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


st.markdown("""
<div class="custom-card">
    <div style="font-size:42px; font-weight:800; color:#1f2937;">
        byggTotal
    </div>
    <div style="font-size:20px; font-weight:600; color:#374151; margin-top:6px;">
        Mengde-, kalkyle- og CO₂-verktøy for Excel og IFC
    </div>
    <div style="margin-top:10px; color:#6b7280; font-size:15px;">
        Applikasjonen leser Excel- eller IFC-modeller og presenterer mengder, kostnader, CO₂-beregninger, materialbytte, NS3420-kobling og modellvisning i ett samlet grensesnitt.
    </div>
</div>
""", unsafe_allow_html=True)

st.sidebar.title("byggTotal")
valg = st.sidebar.radio(
    "Velg side",
    ["Mengder", "Pristilbud", "Analyse", "Materialbytte", "CO₂-regnskap", "3D-modell", "Rapport"],
)

with st.sidebar:
    st.header("Fil og innstillinger")
    uploaded_excel = st.file_uploader("Last opp Excel-fil (.xlsx)", type=["xlsx"])
    uploaded_ifc = st.file_uploader("Last opp IFC-fil (.ifc)", type=["ifc"])

    st.subheader("Produktvalg fra prisbok")
    deck_variant = st.selectbox(
        "Dekkeløsning",
        ["Hulldekke", "Hulldekke_lavCO2"],
        format_func=lambda x: MATERIAL_DATABASE[x]["label"],
        index=0,
    )

    concrete_variant = st.selectbox(
        "Plasstøpt betong",
        ["Plasstøpt_betong", "Plasstøpt_betong_lavCO2"],
        format_func=lambda x: MATERIAL_DATABASE[x]["label"],
        index=0,
    )

    wall_variant = st.selectbox(
        "Betongvegg",
        ["Betong_vegg", "Betong_vegg_lavCO2"],
        format_func=lambda x: MATERIAL_DATABASE[x]["label"],
        index=0,
    )

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

try:
    if uploaded_ifc is not None:
        filename = uploaded_ifc.name
        data, nodes, forside = build_dataset_from_ifc(uploaded_ifc.getvalue())
    elif uploaded_excel is not None:
        filename = uploaded_excel.name
        data, nodes, forside = build_dataset_from_excel(uploaded_excel.getvalue())
    else:
        st.info("Last opp en Excel-fil eller IFC-fil i sidepanelet for å starte analysen.")
        st.stop()
except Exception as e:
    st.error(f"Kunne ikke lese filen: {e}")
    st.stop()

a, b, c = st.columns([1.6, 1.2, 1.2])
with a:
    st.success(f"Aktiv fil: **{filename}**")
with b:
    st.write("")
with c:
    st.write("")

for col in [
    "Segment",
    "Type",
    "Knutepunkter",
    "Material / Tverrsnitt",
    "Lengde [m]",
    "Areal [m2]",
    "Volum [m3]",
    "Vekt [kg]",
    "materiale",
    "Endret IFC",
    "Mengdegrunnlag",
]:
    if col not in data.columns:
        data[col] = pd.NA

data["Produktnøkkel"] = data.apply(
    lambda row: detect_product_key(row, deck_variant, concrete_variant, wall_variant),
    axis=1,
)

data["Produktnavn"] = data["Produktnøkkel"].apply(
    lambda key: MATERIAL_DATABASE.get(key, MATERIAL_DATABASE["Ukjent"])["label"]
)

data["NS3420-kode"] = data.apply(map_ns3420_code, axis=1)

data["Kostnad [kr]"] = data.apply(
    lambda row: cost_for_row(row, deck_variant, concrete_variant, wall_variant),
    axis=1,
)

data["CO2 [kgCO2e]"] = data.apply(
    lambda row: co2_for_row(row, deck_variant, concrete_variant, wall_variant, use_epd=use_epd),
    axis=1,
)

param = {}
if not forside.empty:
    for _, row in forside.iterrows():
        if len(row) >= 2 and pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
            param[str(row.iloc[0]).strip()] = row.iloc[1]

with st.container():
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        type_options = sorted([x for x in data["Type"].dropna().unique().tolist()])
        selected_types = st.multiselect("Type", type_options, default=type_options)

    with c2:
        mat_options = sorted([x for x in data["materiale"].dropna().unique().tolist()])
        selected_materials = st.multiselect("Materiale", mat_options, default=mat_options)

    with c3:
        profile_options = sorted([x for x in data["Material / Tverrsnitt"].dropna().unique().tolist()])
        selected_profiles = st.multiselect(
            "Profil / tverrsnitt",
            profile_options,
            default=profile_options[:8] if len(profile_options) > 8 else profile_options,
        )

    with c4:
        max_length = float(pd.to_numeric(data["Lengde [m]"], errors="coerce").fillna(0).max() or 0)
        length_range = st.slider(
            "Lengdeintervall [m]",
            0.0,
            max(1.0, max_length),
            (0.0, max(1.0, max_length)),
        )

filtered = data.copy()

if selected_types:
    filtered = filtered[filtered["Type"].isin(selected_types)]
if selected_materials:
    filtered = filtered[filtered["materiale"].isin(selected_materials)]
if selected_profiles:
    filtered = filtered[filtered["Material / Tverrsnitt"].isin(selected_profiles)]

filtered = filtered[
    (pd.to_numeric(filtered["Lengde [m]"], errors="coerce").fillna(0) >= length_range[0])
    & (pd.to_numeric(filtered["Lengde [m]"], errors="coerce").fillna(0) <= length_range[1])
]

summary = (
    filtered.groupby(["Type", "materiale"], dropna=False)
    .agg(
        antall=("Segment", "count"),
        areal_m2=("Areal [m2]", "sum"),
        lengde_m=("Lengde [m]", "sum"),
        volum_m3=("Volum [m3]", "sum"),
        vekt_kg=("Vekt [kg]", "sum"),
        kostnad_kr=("Kostnad [kr]", "sum"),
        co2_kg=("CO2 [kgCO2e]", "sum"),
    )
    .reset_index()
    .sort_values(["Type", "materiale"])
)

material_summary = (
    filtered.groupby(["materiale", "Produktnavn", "NS3420-kode"], dropna=False)
    .agg(
        antall=("Segment", "count"),
        areal_m2=("Areal [m2]", "sum"),
        lengde_m=("Lengde [m]", "sum"),
        volum_m3=("Volum [m3]", "sum"),
        vekt_kg=("Vekt [kg]", "sum"),
        kostnad_kr=("Kostnad [kr]", "sum"),
        co2_kg=("CO2 [kgCO2e]", "sum"),
    )
    .reset_index()
    .sort_values("kostnad_kr", ascending=False)
)

swap_df = pd.DataFrame()
ifc_change_log = pd.DataFrame()

if valg == "Mengder":
    st.header("📊 Mengder")

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1:
        metric_card("Elementer", f"{len(filtered):,}".replace(",", " "))
    with k2:
        metric_card("Total lengde", f"{filtered['Lengde [m]'].sum():,.1f} m".replace(",", " "))
    with k3:
        metric_card("Total areal", f"{filtered['Areal [m2]'].sum():,.1f} m²".replace(",", " "))
    with k4:
        metric_card("Stålvekt", f"{filtered.loc[filtered['materiale']=='Stål', 'Vekt [kg]'].sum():,.0f} kg".replace(",", " "))
    with k5:
        metric_card("Estimert kostnad", f"{filtered['Kostnad [kr]'].sum():,.0f} kr".replace(",", " "))
    with k6:
        metric_card("CO₂-avtrykk", f"{filtered['CO2 [kgCO2e]'].sum():,.0f} kgCO₂e".replace(",", " "))

    left, right = st.columns([1.2, 1])

    with left:
        st.subheader("Oppsummering per type og materiale")
        st.dataframe(
            summary.style.format(
                {
                    "areal_m2": "{:.1f}",
                    "lengde_m": "{:.1f}",
                    "volum_m3": "{:.3f}",
                    "vekt_kg": "{:.0f}",
                    "kostnad_kr": "{:,.0f}",
                    "co2_kg": "{:,.0f}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Oppsummering per profil / tverrsnitt")
        profiles = (
            filtered.groupby(["Material / Tverrsnitt", "Produktnavn", "Mengdegrunnlag"], dropna=False)
            .agg(
                antall=("Segment", "count"),
                areal_m2=("Areal [m2]", "sum"),
                lengde_m=("Lengde [m]", "sum"),
                kostnad_kr=("Kostnad [kr]", "sum"),
                co2_kg=("CO2 [kgCO2e]", "sum"),
            )
            .reset_index()
            .sort_values("kostnad_kr", ascending=False)
        )

        st.dataframe(
            profiles.style.format(
                {
                    "areal_m2": "{:.1f}",
                    "lengde_m": "{:.1f}",
                    "kostnad_kr": "{:,.0f}",
                    "co2_kg": "{:,.0f}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    with right:
        st.subheader("Kostnadsfordeling")
        pie_data = summary[summary["kostnad_kr"] > 0].copy()

        if not pie_data.empty:
            pie_data["navn"] = pie_data["Type"].fillna("Ukjent") + " – " + pie_data["materiale"].fillna("Ukjent")
            fig1, ax1 = plt.subplots(figsize=(6, 5))
            ax1.pie(pie_data["kostnad_kr"], labels=pie_data["navn"], autopct="%1.1f%%", startangle=90)
            ax1.axis("equal")
            st.pyplot(fig1)
        else:
            st.info("Ingen kostnadsdata er tilgjengelige for valgt utvalg.")

        st.subheader("CO₂ per produkt")
        co2_data = filtered.groupby("Produktnavn", dropna=False)["CO2 [kgCO2e]"].sum().reset_index()
        co2_data = co2_data[co2_data["CO2 [kgCO2e]"] > 0]
        if not co2_data.empty:
            fig2, ax2 = plt.subplots(figsize=(6, 4))
            ax2.bar(co2_data["Produktnavn"].fillna("Ukjent"), co2_data["CO2 [kgCO2e]"])
            ax2.set_ylabel("kg CO₂e")
            ax2.set_xlabel("Produkt")
            plt.xticks(rotation=25, ha="right")
            st.pyplot(fig2)
        else:
            st.info("Ingen CO₂-data er tilgjengelige for valgt utvalg.")

    st.subheader("Filtrerte elementer")
    show_cols = [
        c
        for c in [
            "Segment",
            "Type",
            "Knutepunkter",
            "Material / Tverrsnitt",
            "materiale",
            "Produktnøkkel",
            "Produktnavn",
            "NS3420-kode",
            "Mengdegrunnlag",
            "Endret IFC",
            "Lengde [m]",
            "Areal [m2]",
            "Volum [m3]",
            "Vekt [kg]",
            "Kostnad [kr]",
            "CO2 [kgCO2e]",
            "IFC Type",
            "IFC GlobalId",
        ]
        if c in filtered.columns
    ]

    st.dataframe(
        filtered[show_cols].style.format(
            {
                "Lengde [m]": "{:.2f}",
                "Areal [m2]": "{:.2f}",
                "Volum [m3]": "{:.3f}",
                "Vekt [kg]": "{:.0f}",
                "Kostnad [kr]": "{:,.0f}",
                "CO2 [kgCO2e]": "{:,.0f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    csv = filtered[show_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Last ned filtrerte data som CSV",
        csv,
        file_name="filtrerte_mengder.csv",
        mime="text/csv",
    )

    with st.expander("Prosjektparametere"):
        if param:
            param_df = pd.DataFrame({"Parameter": list(param.keys()), "Verdi": list(param.values())})
            st.dataframe(param_df, use_container_width=True, hide_index=True)
        else:
            st.info("Ingen prosjektparametere er registrert i filen.")

    if show_raw:
        with st.expander("Rådata"):
            st.dataframe(data, use_container_width=True)

elif valg == "Pristilbud":
    st.header("💰 Pristilbud")

    total_staal_kg = filtered.loc[filtered["materiale"] == "Stål", "Vekt [kg]"].sum()
    total_limtre_m3 = filtered.loc[filtered["materiale"] == "Limtre", "Volum [m3]"].sum()
    total_massivtre_m3 = filtered.loc[filtered["materiale"] == "Massivtre", "Volum [m3]"].sum()
    total_betong_m3 = filtered.loc[filtered["materiale"] == "Betong", "Volum [m3]"].sum()
    total_pris = filtered["Kostnad [kr]"].sum()
    total_co2 = filtered["CO2 [kgCO2e]"].sum()

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        metric_card("Stålvekt", f"{total_staal_kg:,.0f} kg".replace(",", " "))
    with c2:
        metric_card("Limtrevolum", f"{total_limtre_m3:,.2f} m³".replace(",", " "))
    with c3:
        metric_card("Massivtrevolum", f"{total_massivtre_m3:,.2f} m³".replace(",", " "))
    with c4:
        metric_card("Betongvolum", f"{total_betong_m3:,.2f} m³".replace(",", " "))
    with c5:
        metric_card("Total estimert pris", f"{total_pris:,.0f} kr".replace(",", " "))
    with c6:
        metric_card("Total CO₂", f"{total_co2:,.0f} kgCO₂e".replace(",", " "))

    tilbud = (
        filtered.groupby(["materiale", "Produktnavn", "Material / Tverrsnitt", "NS3420-kode"], dropna=False)
        .agg(
            antall=("Segment", "count"),
            areal_m2=("Areal [m2]", "sum"),
            lengde_m=("Lengde [m]", "sum"),
            volum_m3=("Volum [m3]", "sum"),
            vekt_kg=("Vekt [kg]", "sum"),
            kostnad_kr=("Kostnad [kr]", "sum"),
            co2_kg=("CO2 [kgCO2e]", "sum"),
        )
        .reset_index()
        .sort_values("kostnad_kr", ascending=False)
    )

    st.subheader("Tilbudsgrunnlag")
    st.dataframe(
        tilbud.style.format(
            {
                "areal_m2": "{:.1f}",
                "lengde_m": "{:.1f}",
                "volum_m3": "{:.3f}",
                "vekt_kg": "{:.0f}",
                "kostnad_kr": "{:,.0f}",
                "co2_kg": "{:,.0f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    tilbud_csv = tilbud.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Last ned tilbud som CSV",
        tilbud_csv,
        file_name="pristilbud.csv",
        mime="text/csv",
    )

elif valg == "Analyse":
    st.header("📈 Analyse")

    st.subheader("Materialfordeling")
    mat_count = filtered["materiale"].value_counts(dropna=False)
    if not mat_count.empty:
        st.bar_chart(mat_count)
    else:
        st.info("Ingen materialdata er tilgjengelige.")

    st.subheader("Kostnad per type")
    kost_type = filtered.groupby("Type", dropna=False)["Kostnad [kr]"].sum()
    if not kost_type.empty:
        st.bar_chart(kost_type)
    else:
        st.info("Ingen kostnadsdata er tilgjengelige.")

    st.subheader("Areal per type")
    areal_type = filtered.groupby("Type", dropna=False)["Areal [m2]"].sum()
    if not areal_type.empty:
        st.bar_chart(areal_type)
    else:
        st.info("Ingen arealdata er tilgjengelige.")

    st.subheader("CO₂ per produkt")
    co2_type = filtered.groupby("Produktnavn", dropna=False)["CO2 [kgCO2e]"].sum()
    if not co2_type.empty:
        st.bar_chart(co2_type)
    else:
        st.info("Ingen CO₂-data er tilgjengelige.")

    st.subheader("Mengdegrunnlag")
    q_source = filtered["Mengdegrunnlag"].value_counts(dropna=False)
    st.bar_chart(q_source)

elif valg == "Materialbytte":
    st.header("🔁 Materialbytte")

    st.info(
        "Materialbytte er nå delt i to prinsipper: "
        "Søyler og bjelker bruker material-/tverrsnittbytte. "
        "Vegger og dekker bruker systemvalg fra Norsk Prisbok."
    )

    if data.empty:
        st.warning("Ingen data er tilgjengelige.")
        st.stop()

    col1, col2, col3 = st.columns(3)

    with col1:
        available_types = sorted([x for x in data["Type"].dropna().unique().tolist()])
        if not available_types:
            st.warning("Ingen elementtyper er tilgjengelige i datasettet.")
            st.stop()

        default_type = "Søyle" if "Søyle" in available_types else available_types[0]
        selected_swap_type = st.selectbox(
            "Elementtype som skal byttes",
            available_types,
            index=available_types.index(default_type),
        )

    with col2:
        available_materials_for_type = sorted(
            data.loc[data["Type"] == selected_swap_type, "materiale"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if not available_materials_for_type:
            st.warning("Ingen materialer er tilgjengelige for valgt elementtype.")
            st.stop()

        default_from = "Stål" if "Stål" in available_materials_for_type else available_materials_for_type[0]
        from_material = st.selectbox(
            "Nåværende materiale",
            available_materials_for_type,
            index=available_materials_for_type.index(default_from),
        )

    with col3:
        target_options = get_swap_target_options(selected_swap_type, from_material)
        target_key = st.selectbox(
            "Nytt system / materiale",
            target_options,
            format_func=format_swap_target_option,
            index=0,
        )

    swap_defaults = get_swap_target_defaults(target_key)

    st.caption(
        f"Kilde: {swap_defaults.get('source', '-')}"
        + (f" | Post: {swap_defaults.get('npb_code')}" if swap_defaults.get("npb_code") else "")
        + (f" | ÅK/enh: {swap_defaults.get('ak')}" if swap_defaults.get("ak", 0) else "")
    )

    area_based_target = is_area_based_swap_target(target_key)

    new_profile_text = ""
    if not area_based_target:
        st.markdown("#### Nytt tverrsnitt / ObjectType")

        profile_material = target_key if target_key in PROFILE_LIBRARY else classify_material(target_key)
        profile_options = get_swap_profile_options(data, profile_material)

        default_profile = "115x360"
        if default_profile in profile_options:
            default_index = profile_options.index(default_profile)
        elif profile_options:
            default_index = 0
        else:
            default_index = None

        use_manual_profile = st.toggle("Manuell inntasting av tverrsnitt", value=False)

        if use_manual_profile:
            new_profile_text = st.text_input(
                "Oppgi nytt tverrsnitt",
                value=default_profile,
                label_visibility="collapsed",
            )
        else:
            if profile_options:
                new_profile_text = st.selectbox(
                    "Velg tverrsnitt",
                    profile_options,
                    index=default_index if default_index is not None else 0,
                    label_visibility="collapsed",
                )
            else:
                new_profile_text = st.text_input(
                    "Oppgi nytt tverrsnitt",
                    value=default_profile,
                    label_visibility="collapsed",
                )
    else:
        st.markdown("#### Systemvalg")
        st.caption("For vegger og dekker brukes arealbasert Norsk Prisbok-post. Tverrsnitt er derfor ikke nødvendig her.")

    s1, s2, s3 = st.columns(3)
    with s1:
        if swap_defaults["price_unit"] == "kg":
            density_text = f"{swap_defaults['density']:,.0f} kg/m³".replace(",", " ")
        else:
            density_text = f"{swap_defaults['density']:,.0f} kg/m³".replace(",", " ") if swap_defaults["density"] else "-"
        metric_card("Tetthet", density_text)

    with s2:
        if swap_defaults["price_unit"] == "kg":
            unit_label = "kr/kg"
        elif swap_defaults["price_unit"] == "m3":
            unit_label = "kr/m³"
        elif swap_defaults["price_unit"] == "m2":
            unit_label = "kr/m²"
        else:
            unit_label = "-"
        metric_card("Prisgrunnlag", f"{swap_defaults['price']:,.0f} {unit_label}".replace(",", " "))

    with s3:
        if swap_defaults["price_unit"] == "kg":
            co2_unit = "kgCO₂e/kg"
        elif swap_defaults["price_unit"] == "m3":
            co2_unit = "kgCO₂e/m³"
        elif swap_defaults["price_unit"] == "m2":
            co2_unit = "kgCO₂e/m²"
        else:
            co2_unit = "-"
        metric_card("CO₂-faktor", f"{swap_defaults['co2']:,.2f} {co2_unit}".replace(",", " "))

    swap_df = calculate_material_swap(
        source_df=data,
        selected_type=selected_swap_type,
        from_material=from_material,
        target_key=target_key,
        new_profile_text=new_profile_text,
    )

    if not area_based_target and not swap_df.empty and swap_df["Ny vekt [kg]"].sum() > 1_000_000:
        st.warning("Ny vekt etter materialbytte er høy. Kontroller modellgrunnlag, måleenheter og valgt tverrsnitt.")

    if swap_df.empty:
        st.warning("Ingen elementer samsvarer med valgt elementtype og materiale.")
    else:
        st.subheader("Forhåndsvisning av materialbytte")
        st.caption("Dette er en simulering før IFC-filen eksporteres. Ingen endringer er skrevet til modellen ennå.")

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            metric_card("Antall elementer", f"{len(swap_df):,}".replace(",", " "))
        with m2:
            metric_card("Gammel kostnad", f"{swap_df['Gammel kostnad [kr]'].sum():,.0f} kr".replace(",", " "))
        with m3:
            metric_card("Ny kostnad", f"{swap_df['Ny kostnad [kr]'].sum():,.0f} kr".replace(",", " "))
        with m4:
            metric_card("Kostnadsendring", f"{swap_df['Kostnadsendring [kr]'].sum():,.0f} kr".replace(",", " "))

        m5, m6, m7, m8 = st.columns(4)
        with m5:
            metric_card("Gammel vekt", f"{swap_df['Gammel vekt [kg]'].sum():,.0f} kg".replace(",", " "))
        with m6:
            metric_card("Ny vekt", f"{swap_df['Ny vekt [kg]'].sum():,.0f} kg".replace(",", " "))
        with m7:
            metric_card("Vektendring", f"{swap_df['Vektendring [kg]'].sum():,.0f} kg".replace(",", " "))
        with m8:
            metric_card("CO₂-endring", f"{swap_df['CO2-endring [kgCO2e]'].sum():,.0f} kgCO₂e".replace(",", " "))

        preview_before_after = pd.DataFrame(
            {
                "Parameter": ["Kostnad [kr]", "Vekt [kg]", "CO2 [kgCO2e]"],
                "Før": [
                    swap_df["Gammel kostnad [kr]"].sum(),
                    swap_df["Gammel vekt [kg]"].sum(),
                    swap_df["Gammel CO2 [kgCO2e]"].sum(),
                ],
                "Etter": [
                    swap_df["Ny kostnad [kr]"].sum(),
                    swap_df["Ny vekt [kg]"].sum(),
                    swap_df["Ny CO2 [kgCO2e]"].sum(),
                ],
            }
        )
        preview_before_after["Endring"] = preview_before_after["Etter"] - preview_before_after["Før"]

        left_preview, right_preview = st.columns([1.1, 1])

        with left_preview:
            st.markdown("#### Før / etter-oppsummering")
            st.dataframe(
                preview_before_after.style.format(
                    {
                        "Før": "{:,.0f}",
                        "Etter": "{:,.0f}",
                        "Endring": "{:,.0f}",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        with right_preview:
            st.markdown("#### Før / etter-diagram")
            fig_preview_bar, ax_preview_bar = plt.subplots(figsize=(6, 4))
            x = range(len(preview_before_after))
            width = 0.35
            ax_preview_bar.bar(
                [i - width / 2 for i in x],
                preview_before_after["Før"],
                width=width,
                label="Før",
            )
            ax_preview_bar.bar(
                [i + width / 2 for i in x],
                preview_before_after["Etter"],
                width=width,
                label="Etter",
            )
            ax_preview_bar.set_xticks(list(x))
            ax_preview_bar.set_xticklabels(preview_before_after["Parameter"], rotation=15, ha="right")
            ax_preview_bar.legend()
            ax_preview_bar.set_ylabel("Verdi")
            st.pyplot(fig_preview_bar)

        st.subheader("Før / etter materialbytte")
        show_swap_cols = [
            c
            for c in [
                "Segment",
                "Type",
                "IFC Type",
                "IFC GlobalId",
                "Gammelt materiale",
                "Nytt materiale",
                "Nytt systemvalg",
                "Norsk Prisbok-kode",
                "ÅK/enh",
                "Material / Tverrsnitt",
                "Nytt tverrsnitt",
                "Byttelengde [m]",
                "Nytt tverrsnittsareal [m2]",
                "Lengde [m]",
                "Areal [m2]",
                "Gammelt volum [m3]",
                "Nytt volum [m3]",
                "Gammel vekt [kg]",
                "Ny vekt [kg]",
                "Gammel kostnad [kr]",
                "Ny kostnad [kr]",
                "Gammel CO2 [kgCO2e]",
                "Ny CO2 [kgCO2e]",
                "Kostnadsendring [kr]",
                "Vektendring [kg]",
                "CO2-endring [kgCO2e]",
                "Byttemetode",
                "Prisgrunnlag",
                "Tetthet brukt [kg/m3]",
                "CO2-faktor brukt",
            ]
            if c in swap_df.columns
        ]

        st.dataframe(
            swap_df[show_swap_cols].style.format(
                {
                    "Byttelengde [m]": "{:.2f}",
                    "Nytt tverrsnittsareal [m2]": "{:.4f}",
                    "Lengde [m]": "{:.2f}",
                    "Areal [m2]": "{:.2f}",
                    "Gammelt volum [m3]": "{:.3f}",
                    "Nytt volum [m3]": "{:.3f}",
                    "Gammel vekt [kg]": "{:.0f}",
                    "Ny vekt [kg]": "{:.0f}",
                    "Gammel kostnad [kr]": "{:,.0f}",
                    "Ny kostnad [kr]": "{:,.0f}",
                    "Gammel CO2 [kgCO2e]": "{:,.0f}",
                    "Ny CO2 [kgCO2e]": "{:,.0f}",
                    "Kostnadsendring [kr]": "{:,.0f}",
                    "Vektendring [kg]": "{:,.0f}",
                    "CO2-endring [kgCO2e]": "{:,.0f}",
                    "Tetthet brukt [kg/m3]": "{:.0f}",
                    "CO2-faktor brukt": "{:.2f}",
                    "ÅK/enh": "{:.2f}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        swap_csv = swap_df[show_swap_cols].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Last ned endringsliste som CSV",
            swap_csv,
            file_name="materialbytte.csv",
            mime="text/csv",
        )

        if uploaded_ifc is not None and "IFC GlobalId" in swap_df.columns:
            st.markdown("---")
            st.subheader("3D-forhåndsvisning av materialbytte")
            st.caption("Rosa elementer er de som vil bli byttet dersom du eksporterer ny IFC-fil.")

            p1, p2 = st.columns(2)
            with p1:
                preview_mode = st.radio(
                    "Forhåndsvisningsmodus",
                    ["Kun elementer som skal byttes", "Hele modellen med markerte endringer"],
                    horizontal=True,
                )
            with p2:
                preview_max_elements = st.slider(
                    "Maks antall elementer i forhåndsvisning",
                    min_value=100,
                    max_value=5000,
                    value=1500,
                    step=100,
                )

            preview_ids = tuple(sorted(set(swap_df["IFC GlobalId"].dropna().astype(str).tolist())))
            show_only_preview = preview_mode == "Kun elementer som skal byttes"
            visible_ids_for_preview = preview_ids if show_only_preview else None

            try:
                preview_meshes = extract_ifc_meshes_filtered(
                    uploaded_ifc.getvalue(),
                    visible_ids_tuple=visible_ids_for_preview,
                    max_elements=preview_max_elements,
                )

                if not preview_meshes:
                    st.warning("Ingen geometri ble funnet for forhåndsvisningen.")
                else:
                    fig_preview = build_ifc_3d_preview_figure(
                        preview_meshes,
                        preview_ids=preview_ids,
                        show_only_preview=show_only_preview,
                        preview_material=swap_defaults["label"],
                    )
                    st.plotly_chart(fig_preview, use_container_width=True)

                    preview_legend_df = pd.DataFrame(
                        {
                            "Visning": ["Rosa", "Opprinnelig farge / transparent"],
                            "Betydning": ["Element som vil bli byttet", "Eksisterende element i bakgrunn"],
                        }
                    )
                    st.dataframe(preview_legend_df, use_container_width=True, hide_index=True)

            except Exception as e:
                st.warning(f"Kunne ikke generere 3D-forhåndsvisning: {e}")

    st.markdown("---")
    st.subheader("Eksport av oppdatert IFC-fil")

    if uploaded_ifc is None:
        st.info("IFC-eksport er tilgjengelig når en IFC-fil er lastet opp.")
    else:
        st.markdown(
            '<div class="small-muted">Endrede elementer merkes med ByggTotal-egenskaper i IFC-filen. Når filen åpnes i byggTotal, vises disse elementene i rosa i 3D-modellen.</div>',
            unsafe_allow_html=True,
        )

        if st.button("Generer ny IFC-fil"):
            try:
                new_ifc_bytes, ifc_change_log = export_ifc_material_swap(
                    ifc_bytes=uploaded_ifc.getvalue(),
                    source_df=data,
                    selected_type=selected_swap_type,
                    from_material=from_material,
                    target_key=target_key,
                    new_profile_text=new_profile_text,
                )

                if new_ifc_bytes is None or ifc_change_log.empty:
                    st.warning("Ingen elementer ble oppdatert i IFC-filen.")
                else:
                    st.success(f"Ny IFC-fil er generert. Antall oppdaterte elementer: {len(ifc_change_log)}.")

                    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(target_key))
                    out_name = Path(uploaded_ifc.name).stem + f"_materialbytte_{safe_name}.ifc"
                    st.download_button(
                        "Last ned ny IFC-fil",
                        data=new_ifc_bytes,
                        file_name=out_name,
                        mime="application/octet-stream",
                    )

                    st.subheader("Endringslogg for IFC")
                    st.dataframe(ifc_change_log, use_container_width=True, hide_index=True)

                    ifc_change_csv = ifc_change_log.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        "Last ned IFC-endringslogg som CSV",
                        data=ifc_change_csv,
                        file_name="ifc_endringslogg.csv",
                        mime="text/csv",
                    )
            except Exception as e:
                st.error(f"Kunne ikke generere IFC-fil: {e}")

elif valg == "CO₂-regnskap":
    st.header("🌍 CO₂-regnskap")

    total_co2 = filtered["CO2 [kgCO2e]"].sum()
    total_cost = filtered["Kostnad [kr]"].sum()

    c1, c2, c3 = st.columns(3)
    with c1:
        metric_card("Totalt CO₂-avtrykk", f"{total_co2:,.0f} kgCO₂e".replace(",", " "))
    with c2:
        metric_card("Estimert kostnad", f"{total_cost:,.0f} kr".replace(",", " "))
    with c3:
        metric_card("CO₂ per element", f"{(total_co2 / len(filtered)) if len(filtered) > 0 else 0:,.1f}".replace(",", " "))

    st.subheader("CO₂ per materiale")
    co2_material = (
        filtered.groupby(["materiale", "Produktnavn", "NS3420-kode"], dropna=False)
        .agg(
            antall=("Segment", "count"),
            areal_m2=("Areal [m2]", "sum"),
            volum_m3=("Volum [m3]", "sum"),
            vekt_kg=("Vekt [kg]", "sum"),
            co2_kg=("CO2 [kgCO2e]", "sum"),
            kostnad_kr=("Kostnad [kr]", "sum"),
        )
        .reset_index()
        .sort_values("co2_kg", ascending=False)
    )

    st.dataframe(
        co2_material.style.format(
            {
                "areal_m2": "{:.1f}",
                "volum_m3": "{:.3f}",
                "vekt_kg": "{:.0f}",
                "co2_kg": "{:,.0f}",
                "kostnad_kr": "{:,.0f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    left, right = st.columns(2)

    with left:
        chart_data = co2_material[co2_material["co2_kg"] > 0].copy()
        if not chart_data.empty:
            fig4, ax4 = plt.subplots(figsize=(6, 4))
            ax4.bar(chart_data["Produktnavn"].fillna("Ukjent"), chart_data["co2_kg"])
            ax4.set_ylabel("kg CO₂e")
            ax4.set_xlabel("Produkt")
            ax4.set_title("CO₂ per produkt")
            plt.xticks(rotation=25, ha="right")
            st.pyplot(fig4)
        else:
            st.info("Ingen CO₂-data er tilgjengelige for valgt utvalg.")

    with right:
        chart_data = co2_material[co2_material["kostnad_kr"] > 0].copy()
        if not chart_data.empty:
            fig5, ax5 = plt.subplots(figsize=(6, 4))
            ax5.bar(chart_data["Produktnavn"].fillna("Ukjent"), chart_data["kostnad_kr"])
            ax5.set_ylabel("kr")
            ax5.set_xlabel("Produkt")
            ax5.set_title("Kostnad per produkt")
            plt.xticks(rotation=25, ha="right")
            st.pyplot(fig5)
        else:
            st.info("Ingen kostnadsdata er tilgjengelige for valgt utvalg.")

    epd_info_df = pd.DataFrame(
        [
            {
                "Produktnøkkel": key,
                "Enhet": value["unit"],
                "CO₂-faktor": value["co2"],
                "Kilde": value["source"],
            }
            for key, value in EPD_DATABASE.items()
        ]
    )
    with st.expander("Aktive CO₂-faktorer"):
        st.dataframe(epd_info_df, use_container_width=True, hide_index=True)

elif valg == "3D-modell":
    st.header("🧊 3D-modellvisning")

    if uploaded_ifc is None:
        st.info("3D-modellvisning er tilgjengelig når en IFC-fil er lastet opp.")
        st.stop()

    st.markdown(
        "3D-visningen viser geometri fra IFC-modellen. Elementer som er endret i eksportert IFC via byggTotal markeres i rosa."
    )

    left, right = st.columns(2)

    with left:
        visning = st.radio(
            "Visning",
            ["Kun filtrerte elementer", "Alle elementer"],
            horizontal=True,
        )

    with right:
        max_elements_3d = st.slider(
            "Maks antall elementer i 3D-visning",
            min_value=100,
            max_value=5000,
            value=1500,
            step=100,
        )

    if visning == "Kun filtrerte elementer":
        visible_ids = tuple(sorted(set(filtered["IFC GlobalId"].dropna().astype(str).tolist())))
    else:
        visible_ids = None

    try:
        meshes = extract_ifc_meshes_filtered(
            uploaded_ifc.getvalue(),
            visible_ids_tuple=visible_ids,
            max_elements=max_elements_3d,
        )
    except Exception as e:
        st.error(f"Kunne ikke generere 3D-visning: {e}")
        st.stop()

    if not meshes:
        st.warning("Ingen 3D-geometri ble funnet for valgt utvalg.")
        st.stop()

    fig3d = build_ifc_3d_figure(meshes)
    st.plotly_chart(fig3d, use_container_width=True)

    st.subheader("Fargeforklaring")
    legend_df = pd.DataFrame(
        {
            "Materiale": ["Stål", "Betong", "Limtre", "Massivtre", "Tre", "Ukjent", "Endret IFC"],
            "Farge": ["Blå", "Grå", "Brun", "Grønn", "Trebrun", "Lys grå", "Rosa"],
        }
    )
    st.dataframe(legend_df, use_container_width=True, hide_index=True)

    changed_count = sum(1 for m in meshes if m.get("changed", False))
    if visning == "Alle elementer":
        if changed_count > 0:
            st.success(
                f"Viser {len(meshes)} elementer i 3D fra hele IFC-modellen. {changed_count} endrede IFC-elementer er markert i rosa."
            )
        else:
            st.success(f"Viser {len(meshes)} elementer i 3D fra hele IFC-modellen.")
    else:
        if changed_count > 0:
            st.success(
                f"Viser {len(meshes)} filtrerte elementer i 3D. {changed_count} endrede IFC-elementer er markert i rosa."
            )
        else:
            st.success(f"Viser {len(meshes)} filtrerte elementer i 3D.")

elif valg == "Rapport":
    st.header("📝 Rapport og eksport")

    summary_dict = make_report_summary_dict(filename, filtered)

    c1, c2, c3 = st.columns(3)
    with c1:
        metric_card("Elementer", f"{summary_dict['Antall elementer']:,}".replace(",", " "))
    with c2:
        metric_card("Total kostnad", f"{summary_dict['Total kostnad [kr]']:,.0f} kr".replace(",", " "))
    with c3:
        metric_card("Total CO₂", f"{summary_dict['Total CO2 [kgCO2e]']:,.0f} kgCO₂e".replace(",", " "))

    st.subheader("Rapportsammendrag")
    report_df = pd.DataFrame({"Parameter": list(summary_dict.keys()), "Verdi": list(summary_dict.values())})
    st.dataframe(report_df, use_container_width=True, hide_index=True)

    st.subheader("Materialoversikt")
    export_material_summary = material_summary.copy()
    st.dataframe(
        export_material_summary.style.format(
            {
                "areal_m2": "{:.1f}",
                "lengde_m": "{:.1f}",
                "volum_m3": "{:.3f}",
                "vekt_kg": "{:.0f}",
                "kostnad_kr": "{:,.0f}",
                "co2_kg": "{:,.0f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    include_swap = st.checkbox("Ta med materialbytte i rapporten", value=False)

    export_swap_summary = None
    if include_swap and not swap_df.empty:
        export_swap_summary = swap_df[
            [
                c
                for c in [
                    "Segment",
                    "Gammelt materiale",
                    "Nytt materiale",
                    "Norsk Prisbok-kode",
                    "ÅK/enh",
                    "Gammel kostnad [kr]",
                    "Ny kostnad [kr]",
                    "Gammel CO2 [kgCO2e]",
                    "Ny CO2 [kgCO2e]",
                    "Kostnadsendring [kr]",
                    "CO2-endring [kgCO2e]",
                    "Prisgrunnlag",
                ]
                if c in swap_df.columns
            ]
        ]

    docx_bytes = build_docx_report(summary_dict, export_material_summary, export_swap_summary)
    pdf_bytes = build_pdf_report(summary_dict, export_material_summary, export_swap_summary)

    col_a, col_b = st.columns(2)

    with col_a:
        if docx_bytes is not None:
            st.download_button(
                "Last ned rapport som Word",
                data=docx_bytes,
                file_name="byggtotal_rapport.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.info("Word-eksport er ikke tilgjengelig i dette miljøet.")

    with col_b:
        if pdf_bytes is not None:
            st.download_button(
                "Last ned rapport som PDF",
                data=pdf_bytes,
                file_name="byggtotal_rapport.pdf",
                mime="application/pdf",
            )
        else:
            st.info("PDF-eksport er ikke tilgjengelig i dette miljøet.")

st.markdown("---")
st.markdown("**byggTotal**")
