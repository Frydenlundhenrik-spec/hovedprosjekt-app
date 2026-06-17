# ground_module.py
# Utvidet grunnmodul for byggTotal.
# Håndterer geoteknikk, georeferering, fundamentering, peling og CO₂ fra grunnarbeider.

import io
import json
import math
import re
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# DATABASER
# ---------------------------------------------------------------------------

# Geotekniske jordarter med egenskaper
SOIL_DATABASE = {
    "Fjell": {
        "label": "Fjell",
        "bearing_capacity_kPa": 5000.0,
        "unit_weight_kNm3": 26.0,
        "friction_angle_deg": 45.0,
        "cohesion_kPa": 0.0,
        "compression_modulus_MPa": 1000.0,
        "co2_excavation_kgm3": 12.0,
        "color": "#808080",
    },
    "Morene": {
        "label": "Morene / grus",
        "bearing_capacity_kPa": 400.0,
        "unit_weight_kNm3": 21.0,
        "friction_angle_deg": 38.0,
        "cohesion_kPa": 0.0,
        "compression_modulus_MPa": 50.0,
        "co2_excavation_kgm3": 4.0,
        "color": "#C8A882",
    },
    "Sand": {
        "label": "Sand",
        "bearing_capacity_kPa": 200.0,
        "unit_weight_kNm3": 19.0,
        "friction_angle_deg": 32.0,
        "cohesion_kPa": 0.0,
        "compression_modulus_MPa": 30.0,
        "co2_excavation_kgm3": 3.5,
        "color": "#E8D5A0",
    },
    "Silt": {
        "label": "Silt",
        "bearing_capacity_kPa": 80.0,
        "unit_weight_kNm3": 18.0,
        "friction_angle_deg": 28.0,
        "cohesion_kPa": 5.0,
        "compression_modulus_MPa": 8.0,
        "co2_excavation_kgm3": 3.0,
        "color": "#D4C5A0",
    },
    "Leire": {
        "label": "Leire",
        "bearing_capacity_kPa": 40.0,
        "unit_weight_kNm3": 17.5,
        "friction_angle_deg": 0.0,
        "cohesion_kPa": 20.0,
        "compression_modulus_MPa": 2.5,
        "co2_excavation_kgm3": 2.5,
        "color": "#B5A088",
    },
    "Kvikkleire": {
        "label": "Kvikkleire",
        "bearing_capacity_kPa": 15.0,
        "unit_weight_kNm3": 17.0,
        "friction_angle_deg": 0.0,
        "cohesion_kPa": 8.0,
        "compression_modulus_MPa": 1.0,
        "co2_excavation_kgm3": 2.5,
        "color": "#C4A0A0",
    },
    "Torv": {
        "label": "Torv / myr",
        "bearing_capacity_kPa": 10.0,
        "unit_weight_kNm3": 12.0,
        "friction_angle_deg": 0.0,
        "cohesion_kPa": 3.0,
        "compression_modulus_MPa": 0.5,
        "co2_excavation_kgm3": 2.0,
        "color": "#6B4C2A",
    },
    "Ukjent": {
        "label": "Ukjent",
        "bearing_capacity_kPa": 50.0,
        "unit_weight_kNm3": 18.0,
        "friction_angle_deg": 25.0,
        "cohesion_kPa": 10.0,
        "compression_modulus_MPa": 5.0,
        "co2_excavation_kgm3": 3.0,
        "color": "#AAAAAA",
    },
}

# Fundamenteringstyper med enhetspriser og CO₂
FOUNDATION_DATABASE = {
    "Platefundament_betong": {
        "label": "Platefundament – betong",
        "unit": "m2",
        "price_kr": 1850.0,
        "co2_kgm2": 105.0,
        "min_bearing_kPa": 100.0,
        "description": "Platefundament egner seg for god bæreevne. Krever komprimert forsterkningslag.",
    },
    "Stripefundament_betong": {
        "label": "Stripefundament – betong",
        "unit": "lm",
        "price_kr": 2800.0,
        "co2_kgm2": 145.0,
        "min_bearing_kPa": 80.0,
        "description": "Stripefundament langs vegger og bærende konstruksjoner.",
    },
    "Punktfundament_betong": {
        "label": "Punktfundament – betong",
        "unit": "stk",
        "price_kr": 18000.0,
        "co2_kgm2": 210.0,
        "min_bearing_kPa": 150.0,
        "description": "Punktfundament for søylebaserte konstruksjoner.",
    },
    "Pelefundament_staal": {
        "label": "Pelefundament – stål H-pile",
        "unit": "lm",
        "price_kr": 1200.0,
        "co2_kgm2": 28.0,
        "min_bearing_kPa": 0.0,
        "description": "Stålpeler til fjell eller fast lag. CO₂ per løpemeter pel.",
    },
    "Pelefundament_betong": {
        "label": "Pelefundament – betongpel",
        "unit": "lm",
        "price_kr": 980.0,
        "co2_kgm2": 35.0,
        "min_bearing_kPa": 0.0,
        "description": "Betongpeler, spuntet eller rammet til bæredyktig lag.",
    },
    "Pelefundament_tre": {
        "label": "Pelefundament – trepel (historisk/rehab)",
        "unit": "lm",
        "price_kr": 450.0,
        "co2_kgm2": 8.0,
        "min_bearing_kPa": 0.0,
        "description": "Trepeler brukes ved rehabilitering av eldre bygg under grunnvannsnivå.",
    },
    "Masseutskifting": {
        "label": "Masseutskifting",
        "unit": "m3",
        "price_kr": 850.0,
        "co2_kgm2": 45.0,
        "min_bearing_kPa": 0.0,
        "description": "Svake masser graves ut og erstattes med komprimert grus/pukk.",
    },
    "Kalksementstabilisering": {
        "label": "Kalk-/sementstabilisering",
        "unit": "m3",
        "price_kr": 650.0,
        "co2_kgm2": 120.0,
        "min_bearing_kPa": 0.0,
        "description": "Blanding av kalk og/eller sement i svake masser for å øke bæreevnen.",
    },
}

# CO₂-faktorer for grunnarbeider (kgCO₂e per enhet)
GROUND_CO2_DATABASE = {
    "Utgraving_løsmasse": {"unit": "m3", "co2": 3.2, "label": "Utgraving løsmasse"},
    "Utgraving_fjell": {"unit": "m3", "co2": 12.0, "label": "Utgraving / sprengning fjell"},
    "Bortkjøring": {"unit": "m3", "co2": 5.5, "label": "Transport og deponi av masser"},
    "Fyllmasse_import": {"unit": "m3", "co2": 8.0, "label": "Importerte fyllmasser"},
    "Pukk_forsterkningslag": {"unit": "m3", "co2": 15.0, "label": "Pukk / forsterkningslag"},
    "Betong_fundamentplate": {"unit": "m3", "co2": 350.0, "label": "Betong i fundamenter"},
    "Armering_staal": {"unit": "kg", "co2": 1.75, "label": "Armeringsstål"},
    "Staalpal": {"unit": "lm", "co2": 28.0, "label": "Stålpel"},
    "Betongpal": {"unit": "lm", "co2": 35.0, "label": "Betongpel"},
}

# ---------------------------------------------------------------------------
# HJELPEFUNKSJONER
# ---------------------------------------------------------------------------

def safe_num(value) -> float:
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def recommend_foundation(soil_type: str, building_area_m2: float, load_kN: float) -> dict:
    """Anbefaler fundamenteringstype basert på jordart og last."""
    soil = SOIL_DATABASE.get(soil_type, SOIL_DATABASE["Ukjent"])
    bearing = soil["bearing_capacity_kPa"]
    required_area = load_kN / bearing if bearing > 0 else float("inf")

    if soil_type in ["Fjell", "Morene"] and bearing >= 200:
        rec = "Platefundament_betong"
        reason = f"God bæreevne ({bearing:.0f} kPa). Platefundament egnet."
    elif soil_type in ["Sand", "Silt"] and bearing >= 80:
        rec = "Stripefundament_betong"
        reason = f"Middels bæreevne ({bearing:.0f} kPa). Stripefundament anbefales."
    elif soil_type in ["Leire", "Kvikkleire", "Torv"]:
        rec = "Pelefundament_staal"
        reason = f"Lav bæreevne ({bearing:.0f} kPa). Peling til bæredyktig lag anbefales."
    else:
        rec = "Platefundament_betong"
        reason = "Ukjent grunnforhold – konservativt valg."

    return {
        "anbefalt_fundament": rec,
        "label": FOUNDATION_DATABASE[rec]["label"],
        "begrunnelse": reason,
        "bæreevne_kPa": bearing,
        "nødvendig_areal_m2": round(required_area, 1) if required_area != float("inf") else "N/A",
    }


def estimate_pile_length(depth_to_rock_m: float, soil_type: str) -> dict:
    """Estimerer pellengde og kostnad."""
    friction_add = 2.0 if soil_type in ["Sand", "Morene"] else 4.0
    total_length = depth_to_rock_m + friction_add
    cost_steel = total_length * FOUNDATION_DATABASE["Pelefundament_staal"]["price_kr"]
    cost_concrete = total_length * FOUNDATION_DATABASE["Pelefundament_betong"]["price_kr"]
    co2_steel = total_length * GROUND_CO2_DATABASE["Staalpal"]["co2"]
    co2_concrete = total_length * GROUND_CO2_DATABASE["Betongpal"]["co2"]
    return {
        "dybde_til_fjell_m": depth_to_rock_m,
        "anbefalt_pellengde_m": round(total_length, 1),
        "kostnad_stålpel_kr": round(cost_steel, 0),
        "kostnad_betongpel_kr": round(cost_concrete, 0),
        "co2_stålpel_kgCO2e": round(co2_steel, 1),
        "co2_betongpel_kgCO2e": round(co2_concrete, 1),
    }


def calculate_foundation_cost(foundation_key: str, quantity: float) -> dict:
    """Beregner kostnad og CO₂ for valgt fundamenteringstype."""
    db = FOUNDATION_DATABASE.get(foundation_key, FOUNDATION_DATABASE["Platefundament_betong"])
    cost = quantity * db["price_kr"]
    co2 = quantity * db["co2_kgm2"]
    return {
        "type": db["label"],
        "mengde": quantity,
        "enhet": db["unit"],
        "enhetspris_kr": db["price_kr"],
        "total_kostnad_kr": round(cost, 0),
        "co2_faktor": db["co2_kgm2"],
        "total_co2_kgCO2e": round(co2, 1),
    }


def calculate_ground_co2(cut_m3: float, fill_m3: float, foundation_area_m2: float,
                          soil_type: str = "Leire", pile_lm: float = 0.0,
                          foundation_type: str = "Platefundament_betong") -> pd.DataFrame:
    """Beregner CO₂-regnskap for grunnarbeider."""
    soil = SOIL_DATABASE.get(soil_type, SOIL_DATABASE["Ukjent"])
    co2_ex = soil["co2_excavation_kgm3"]
    rows = [
        {"Post": "Utgraving", "Enhet": "m3", "Mengde": cut_m3,
         "CO₂-faktor": co2_ex, "CO₂ [kgCO2e]": cut_m3 * co2_ex},
        {"Post": "Bortkjøring av masser", "Enhet": "m3", "Mengde": cut_m3,
         "CO₂-faktor": GROUND_CO2_DATABASE["Bortkjøring"]["co2"],
         "CO₂ [kgCO2e]": cut_m3 * GROUND_CO2_DATABASE["Bortkjøring"]["co2"]},
        {"Post": "Importerte fyllmasser", "Enhet": "m3", "Mengde": fill_m3,
         "CO₂-faktor": GROUND_CO2_DATABASE["Fyllmasse_import"]["co2"],
         "CO₂ [kgCO2e]": fill_m3 * GROUND_CO2_DATABASE["Fyllmasse_import"]["co2"]},
        {"Post": "Forsterkningslag pukk", "Enhet": "m3", "Mengde": foundation_area_m2 * 0.3,
         "CO₂-faktor": GROUND_CO2_DATABASE["Pukk_forsterkningslag"]["co2"],
         "CO₂ [kgCO2e]": foundation_area_m2 * 0.3 * GROUND_CO2_DATABASE["Pukk_forsterkningslag"]["co2"]},
    ]
    if foundation_type in ["Platefundament_betong", "Stripefundament_betong", "Punktfundament_betong"]:
        vol = foundation_area_m2 * 0.25
        rows.append({"Post": "Betong i fundamenter", "Enhet": "m3", "Mengde": vol,
                     "CO₂-faktor": GROUND_CO2_DATABASE["Betong_fundamentplate"]["co2"],
                     "CO₂ [kgCO2e]": vol * GROUND_CO2_DATABASE["Betong_fundamentplate"]["co2"]})
        arme_kg = vol * 80
        rows.append({"Post": "Armeringsstål", "Enhet": "kg", "Mengde": arme_kg,
                     "CO₂-faktor": GROUND_CO2_DATABASE["Armering_staal"]["co2"],
                     "CO₂ [kgCO2e]": arme_kg * GROUND_CO2_DATABASE["Armering_staal"]["co2"]})
    if pile_lm > 0:
        rows.append({"Post": "Peler (stål)", "Enhet": "lm", "Mengde": pile_lm,
                     "CO₂-faktor": GROUND_CO2_DATABASE["Staalpal"]["co2"],
                     "CO₂ [kgCO2e]": pile_lm * GROUND_CO2_DATABASE["Staalpal"]["co2"]})
    df = pd.DataFrame(rows)
    df["CO₂ [kgCO2e]"] = df["CO₂ [kgCO2e]"].round(1)
    return df


# ---------------------------------------------------------------------------
# GEOREFERERING – FILINNLESING
# ---------------------------------------------------------------------------

def load_geojson(file) -> tuple[pd.DataFrame, dict]:
    """Leser GeoJSON og returnerer punkter og metadata."""
    content = json.loads(file.read())
    points = []
    for feature in content.get("features", []):
        geom = feature.get("geometry", {})
        props = feature.get("properties", {}) or {}
        gt = geom.get("type", "")
        if gt == "Point":
            coords = geom.get("coordinates", [])
            if len(coords) >= 2:
                points.append({
                    "X": coords[0], "Y": coords[1],
                    "Z": coords[2] if len(coords) > 2 else 0.0,
                    "Kode": str(props.get("code", props.get("kode", props.get("type", "Ukjent")))),
                    "Punkt": str(props.get("id", props.get("name", len(points) + 1))),
                })
        elif gt == "MultiPoint":
            for c in geom.get("coordinates", []):
                if len(c) >= 2:
                    points.append({
                        "X": c[0], "Y": c[1],
                        "Z": c[2] if len(c) > 2 else 0.0,
                        "Kode": str(props.get("code", "Ukjent")),
                        "Punkt": str(len(points) + 1),
                    })
    if not points:
        raise ValueError("Fant ingen punktgeometrier i GeoJSON-filen.")
    df = pd.DataFrame(points)
    meta = {
        "type": content.get("type", ""),
        "antall_features": len(content.get("features", [])),
        "antall_punkt": len(df),
        "crs": content.get("crs", {}).get("properties", {}).get("name", "Ukjent CRS"),
    }
    return df, meta


def load_dxf_points(file) -> pd.DataFrame:
    """Leser DXF-fil og henter ut punkter (POINT-entities)."""
    try:
        import ezdxf
    except ImportError:
        raise ImportError("ezdxf er ikke installert. Legg til 'ezdxf' i requirements.txt.")
    content = file.read()
    doc = ezdxf.read(io.StringIO(content.decode("utf-8", errors="replace")))
    msp = doc.modelspace()
    points = []
    for entity in msp:
        if entity.dxftype() == "POINT":
            loc = entity.dxf.location
            points.append({
                "X": loc.x, "Y": loc.y, "Z": loc.z,
                "Kode": str(entity.dxf.get("layer", "0")),
                "Punkt": str(len(points) + 1),
            })
        elif entity.dxftype() == "INSERT":
            loc = entity.dxf.insert
            points.append({
                "X": loc.x, "Y": loc.y, "Z": loc.z,
                "Kode": str(entity.dxf.get("layer", "0")),
                "Punkt": str(entity.dxf.get("name", len(points) + 1)),
            })
    if not points:
        raise ValueError("Fant ingen punkter i DXF-filen. Sjekk at filen inneholder POINT-entiteter.")
    return pd.DataFrame(points)


def parse_wms_bbox_from_url(url: str) -> dict:
    """Trekker ut BBOX og lag-info fra en WMS/WFS-URL."""
    info = {}
    for param in url.split("&"):
        kv = param.split("=")
        if len(kv) == 2:
            info[kv[0].upper()] = kv[1]
    return info


# ---------------------------------------------------------------------------
# GEOTEKNISK PROFILPLOT
# ---------------------------------------------------------------------------

def plot_soil_profile(layers: list[dict], groundwater_depth_m: float = None,
                      title: str = "Geoteknisk profil") -> plt.Figure:
    """
    Tegner et enkelt geoteknisk profil.
    layers: liste av dict med 'soil_type' og 'thickness_m'
    """
    fig, ax = plt.subplots(figsize=(4, 7))
    depth = 0.0
    legend_patches = []
    seen = set()
    for layer in layers:
        soil = SOIL_DATABASE.get(layer["soil_type"], SOIL_DATABASE["Ukjent"])
        thickness = layer["thickness_m"]
        color = soil["color"]
        rect = mpatches.FancyBboxPatch(
            (0, -(depth + thickness)), 1, thickness,
            boxstyle="square,pad=0", linewidth=0.5,
            edgecolor="#444", facecolor=color, alpha=0.85
        )
        ax.add_patch(rect)
        ax.text(0.5, -(depth + thickness / 2), f"{soil['label']}\n{thickness:.1f} m",
                ha="center", va="center", fontsize=8, fontweight="bold")
        if layer["soil_type"] not in seen:
            legend_patches.append(mpatches.Patch(color=color, label=soil["label"]))
            seen.add(layer["soil_type"])
        depth += thickness

    if groundwater_depth_m is not None and groundwater_depth_m <= depth:
        gw_y = -groundwater_depth_m
        ax.axhline(y=gw_y, color="blue", linewidth=2, linestyle="--", alpha=0.7)
        ax.text(1.05, gw_y, f"GVN: {groundwater_depth_m:.1f} m", color="blue", fontsize=8, va="center")

    ax.set_xlim(0, 1.5)
    ax.set_ylim(-depth - 0.5, 0.5)
    ax.set_xticks([])
    ax.set_yticks([round(-d, 1) for d in [i * (depth / 5) for i in range(6)]])
    ax.set_yticklabels([f"{abs(d):.1f} m" for d in ax.get_yticks()])
    ax.set_title(title, fontsize=10)
    ax.legend(handles=legend_patches, loc="lower right", fontsize=7)
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    return fig


def plot_bearing_capacity_chart(soil_types: list, foundation_area_m2: float) -> plt.Figure:
    """Stolpediagram over bæreevne for ulike jordarter."""
    labels = [SOIL_DATABASE[s]["label"] for s in soil_types if s in SOIL_DATABASE]
    bearing = [SOIL_DATABASE[s]["bearing_capacity_kPa"] for s in soil_types if s in SOIL_DATABASE]
    colors = [SOIL_DATABASE[s]["color"] for s in soil_types if s in SOIL_DATABASE]
    fig, ax = plt.subplots(figsize=(7, 4))
    bars = ax.bar(labels, bearing, color=colors, edgecolor="#444", linewidth=0.5)
    ax.set_ylabel("Bæreevne [kPa]")
    ax.set_title("Bæreevne per jordart")
    ax.tick_params(axis="x", rotation=25)
    for bar, val in zip(bars, bearing):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 5,
                f"{val:.0f}", ha="center", va="bottom", fontsize=8)
    plt.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# GRUNNVANNSNIVÅ
# ---------------------------------------------------------------------------

def groundwater_risk_assessment(gwl_depth_m: float, foundation_depth_m: float) -> dict:
    """Vurderer risiko knyttet til grunnvannsnivå."""
    margin = gwl_depth_m - foundation_depth_m
    if margin > 1.0:
        risk = "Lav"
        color = "green"
        note = "Grunnvannsnivå er godt over fundamentnivå. Liten risiko for oppdrift og fuktskader."
    elif margin > 0:
        risk = "Middels"
        color = "orange"
        note = "Grunnvannsnivå er nær fundamentnivå. Vurder drenering og membran."
    elif margin > -0.5:
        risk = "Høy"
        color = "red"
        note = "Fundamentet er nær eller under grunnvannsnivå. Oppdrift og vanninntrenging er en risiko."
    else:
        risk = "Kritisk"
        color = "darkred"
        note = "Fundamentet er under grunnvannsnivå. Permanent pumpeanlegg, injeksjon eller senking av GVN er nødvendig."
    return {
        "grunnvannsnivå_m": gwl_depth_m,
        "fundamentdybde_m": foundation_depth_m,
        "margin_m": round(margin, 2),
        "risikonivå": risk,
        "merknad": note,
        "color": color,
    }


# ---------------------------------------------------------------------------
# EKSPORT
# ---------------------------------------------------------------------------

def build_ground_report_df(layers: list[dict], gwl: float, foundation_key: str,
                            qty: float, pile_lm: float,
                            cut_m3: float, fill_m3: float, area_m2: float) -> pd.DataFrame:
    """Samler alle grunndata i én rapport-DataFrame."""
    rows = []
    depth = 0.0
    for i, layer in enumerate(layers):
        soil = SOIL_DATABASE.get(layer["soil_type"], SOIL_DATABASE["Ukjent"])
        depth += layer["thickness_m"]
        rows.append({
            "Lag nr.": i + 1,
            "Jordart": soil["label"],
            "Tykkelse [m]": layer["thickness_m"],
            "Dybde til bunn [m]": round(depth, 2),
            "Bæreevne [kPa]": soil["bearing_capacity_kPa"],
            "Friktionsvinkel [°]": soil["friction_angle_deg"],
            "Kohesjon [kPa]": soil["cohesion_kPa"],
            "CO₂ utgraving [kg/m³]": soil["co2_excavation_kgm3"],
        })
    return pd.DataFrame(rows)
