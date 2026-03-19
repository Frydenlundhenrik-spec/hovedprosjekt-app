import io
import math
import os
import re
import tempfile
from pathlib import Path
from datetime import datetime

import matplotlib.pyplot as plt
import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

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
    page_title="byggTotal – Excel og IFC til kalkyle, mengder og CO₂",
    page_icon="🏗️",
    layout="wide",
)

st.markdown("""
<style>
    .main {
        background-color: #f6f7fb;
    }

    .block-container {
        padding-top: 2rem;
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
</style>
""", unsafe_allow_html=True)

DEFAULT_WORKBOOK = Path(__file__).with_name("example_model.xlsx")
STEEL_DENSITY = 7850
GLULAM_DENSITY = 460
CLT_DENSITY = 500
CONCRETE_DENSITY = 2400


@st.cache_data(show_spinner=False)
def load_sheet_df(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl")


def clean_dataframe(df: pd.DataFrame, required_cols=None) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")
    if required_cols:
        for col in required_cols:
            if col in df.columns:
                df = df[df[col].notna()]
    return df.reset_index(drop=True)


def parse_profile(profile: str):
    text = str(profile or "")
    material = "Ukjent"

    if "stål" in text.lower():
        material = "Stål"
    elif "limtre" in text.lower():
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

    merged["Volum [m3]"] = merged["Lengde [m]"] * merged["areal_m2"]
    merged["Vekt [kg]"] = merged.apply(
        lambda row: row["Volum [m3]"] * STEEL_DENSITY
        if row["materiale"] == "Stål"
        else (
            row["Volum [m3]"] * GLULAM_DENSITY
            if row["materiale"] == "Limtre"
            else (
                row["Volum [m3]"] * CLT_DENSITY
                if row["materiale"] == "Massivtre"
                else math.nan
            )
        ),
        axis=1,
    )

    knutepunkter.columns = [str(c).strip() for c in knutepunkter.columns]
    return merged, knutepunkter, forside


def get_quantity_from_element(element, quantity_names):
    try:
        for rel in getattr(element, "IsDefinedBy", []) or []:
            definition = getattr(rel, "RelatingPropertyDefinition", None)
            if not definition or not definition.is_a("IfcElementQuantity"):
                continue

            for qty in getattr(definition, "Quantities", []) or []:
                qname = getattr(qty, "Name", "")
                if qname not in quantity_names:
                    continue

                for attr in ["LengthValue", "AreaValue", "VolumeValue", "CountValue", "WeightValue"]:
                    if hasattr(qty, attr):
                        return getattr(qty, attr)
    except Exception:
        pass
    return None


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


def classify_material(material_text):
    text = str(material_text or "").lower()

    if any(x in text for x in ["stål", "steel", "s355", "s235", "kfh", "vfh", "rhs", "shs", "ihe", "hea", "heb", "ipe"]):
        return "Stål"

    if any(x in text for x in ["limtre", "glulam", "glt"]):
        return "Limtre"

    if any(x in text for x in ["massivtre", "clt", "cross laminated timber", "krysslaminert"]):
        return "Massivtre"

    if any(x in text for x in ["concrete", "betong", "hulldekke", "hd", "in-situ", "cast in place", "prefab concrete"]):
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


def build_dataset_from_ifc(ifc_bytes: bytes):
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert. Kjør: py -m pip install ifcopenshell")

    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp:
            tmp.write(ifc_bytes)
            temp_path = tmp.name

        model = ifcopenshell.open(temp_path)

        rows = []
        target_types = [
            "IfcBeam",
            "IfcColumn",
            "IfcSlab",
            "IfcWall",
            "IfcWallStandardCase",
            "IfcRoof",
            "IfcMember",
            "IfcFooting",
        ]

        for type_name in target_types:
            for el in model.by_type(type_name):
                global_id = getattr(el, "GlobalId", None)
                name = getattr(el, "Name", None) or global_id or "Ukjent"
                object_type = getattr(el, "ObjectType", None) or ""
                predefined = getattr(el, "PredefinedType", None) or ""

                material_raw = get_ifc_material_name(el)
                materiale = classify_material(material_raw)

                length_m = get_quantity_from_element(el, ["Length", "NetLength", "GrossLength"])
                area_m2 = get_quantity_from_element(el, ["Area", "NetArea", "GrossArea"])
                volume_m3 = get_quantity_from_element(el, ["Volume", "NetVolume", "GrossVolume"])

                length_m = pd.to_numeric(length_m, errors="coerce")
                area_m2 = pd.to_numeric(area_m2, errors="coerce")
                volume_m3 = pd.to_numeric(volume_m3, errors="coerce")

                if materiale == "Stål" and pd.notna(volume_m3):
                    vekt_kg = volume_m3 * STEEL_DENSITY
                elif materiale == "Limtre" and pd.notna(volume_m3):
                    vekt_kg = volume_m3 * GLULAM_DENSITY
                elif materiale == "Massivtre" and pd.notna(volume_m3):
                    vekt_kg = volume_m3 * CLT_DENSITY
                elif materiale == "Betong" and pd.notna(volume_m3):
                    vekt_kg = volume_m3 * CONCRETE_DENSITY
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
def extract_ifc_meshes(ifc_bytes: bytes):
    if ifcopenshell is None:
        raise ImportError("ifcopenshell er ikke installert.")

    temp_path = None
    meshes = []

    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".ifc") as tmp:
            tmp.write(ifc_bytes)
            temp_path = tmp.name

        model = ifcopenshell.open(temp_path)

        settings = ifcopenshell.geom.settings()
        settings.set(settings.USE_WORLD_COORDS, True)

        target_types = [
            "IfcBeam",
            "IfcColumn",
            "IfcSlab",
            "IfcWall",
            "IfcWallStandardCase",
            "IfcRoof",
            "IfcMember",
            "IfcFooting",
        ]

        for type_name in target_types:
            for el in model.by_type(type_name):
                try:
                    shape = ifcopenshell.geom.create_shape(settings, el)
                    geom = shape.geometry

                    verts = geom.verts
                    faces = geom.faces

                    if not verts or not faces:
                        continue

                    x = verts[0::3]
                    y = verts[1::3]
                    z = verts[2::3]

                    i_idx = faces[0::3]
                    j_idx = faces[1::3]
                    k_idx = faces[2::3]

                    material_raw = get_ifc_material_name(el)
                    materiale = classify_material(material_raw)

                    meshes.append(
                        {
                            "global_id": getattr(el, "GlobalId", ""),
                            "name": getattr(el, "Name", "") or getattr(el, "GlobalId", "") or "Ukjent",
                            "type": map_ifc_type(type_name),
                            "ifc_type": type_name,
                            "materiale": materiale,
                            "x": x,
                            "y": y,
                            "z": z,
                            "i": i_idx,
                            "j": j_idx,
                            "k": k_idx,
                        }
                    )
                except Exception:
                    continue

        return meshes

    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def material_color(materiale: str):
    mapping = {
        "Stål": "#4F81BD",
        "Betong": "#A6A6A6",
        "Limtre": "#C58C4B",
        "Massivtre": "#8CBF3F",
        "Tre": "#B97A57",
        "Ukjent": "#D9D9D9",
    }
    return mapping.get(materiale, "#D9D9D9")


def build_ifc_3d_figure(meshes, visible_ids=None, max_elements=300):
    fig = go.Figure()
    count = 0

    for mesh in meshes:
        gid = mesh["global_id"]

        if visible_ids is not None and gid not in visible_ids:
            continue

        fig.add_trace(
            go.Mesh3d(
                x=mesh["x"],
                y=mesh["y"],
                z=mesh["z"],
                i=mesh["i"],
                j=mesh["j"],
                k=mesh["k"],
                color=material_color(mesh["materiale"]),
                opacity=0.95,
                flatshading=True,
                name=f"{mesh['type']} – {mesh['materiale']}",
                hovertext=(
                    f"Navn: {mesh['name']}<br>"
                    f"Type: {mesh['type']}<br>"
                    f"IFC-type: {mesh['ifc_type']}<br>"
                    f"Materiale: {mesh['materiale']}<br>"
                    f"GlobalId: {mesh['global_id']}"
                ),
                hoverinfo="text",
                showscale=False,
            )
        )
        count += 1
        if count >= max_elements:
            break

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
        height=750,
    )

    return fig


def cost_for_row(
    row,
    steel_price,
    glulam_price,
    concrete_price_m3,
    clt_price_m3,
    timber_price_m3,
    unknown_price_m3,
):
    materiale = row.get("materiale", "Ukjent")
    vekt = row.get("Vekt [kg]", math.nan)
    volum = row.get("Volum [m3]", math.nan)

    if materiale == "Stål":
        return (vekt if pd.notna(vekt) else 0) * steel_price
    if materiale == "Limtre":
        return (volum if pd.notna(volum) else 0) * glulam_price
    if materiale == "Massivtre":
        return (volum if pd.notna(volum) else 0) * clt_price_m3
    if materiale == "Betong":
        return (volum if pd.notna(volum) else 0) * concrete_price_m3
    if materiale == "Tre":
        return (volum if pd.notna(volum) else 0) * timber_price_m3
    return (volum if pd.notna(volum) else 0) * unknown_price_m3


def co2_for_row(
    row,
    steel_co2_factor,
    glulam_co2_factor,
    concrete_co2_factor,
    clt_co2_factor,
    timber_co2_factor,
    unknown_co2_factor,
):
    materiale = row.get("materiale", "Ukjent")
    vekt = row.get("Vekt [kg]", math.nan)
    volum = row.get("Volum [m3]", math.nan)

    if materiale == "Stål":
        return (vekt if pd.notna(vekt) else 0) * steel_co2_factor
    if materiale == "Limtre":
        return (volum if pd.notna(volum) else 0) * glulam_co2_factor
    if materiale == "Massivtre":
        return (volum if pd.notna(volum) else 0) * clt_co2_factor
    if materiale == "Betong":
        return (volum if pd.notna(volum) else 0) * concrete_co2_factor
    if materiale == "Tre":
        return (volum if pd.notna(volum) else 0) * timber_co2_factor
    return (volum if pd.notna(volum) else 0) * unknown_co2_factor


def parse_rect_profile_mm(profile_text: str):
    text = str(profile_text or "")
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+[\.,]?\d*", text)]
    if len(nums) >= 2:
        return nums[-2], nums[-1]
    return None, None


def calculate_material_swap(
    source_df: pd.DataFrame,
    selected_type: str,
    from_material: str,
    to_material: str,
    new_profile_text: str,
    new_density: float,
    new_price_per_m3: float,
    steel_co2_factor: float,
    glulam_co2_factor: float,
    concrete_co2_factor: float,
    clt_co2_factor: float,
):
    df = source_df.copy()

    matched = df[(df["Type"] == selected_type) & (df["materiale"] == from_material)].copy()

    if matched.empty:
        return matched

    bredde_mm, hoyde_mm = parse_rect_profile_mm(new_profile_text)

    if bredde_mm and hoyde_mm:
        nytt_areal_m2 = (bredde_mm * hoyde_mm) / 1_000_000
    else:
        nytt_areal_m2 = math.nan

    matched["Gammelt materiale"] = matched["materiale"]
    matched["Nytt materiale"] = to_material
    matched["Nytt tverrsnitt"] = new_profile_text
    matched["Ny bredde [mm]"] = bredde_mm
    matched["Ny høyde [mm]"] = hoyde_mm
    matched["Gammel kostnad [kr]"] = matched["Kostnad [kr]"]
    matched["Gammel vekt [kg]"] = matched["Vekt [kg]"]
    matched["Gammelt volum [m3]"] = matched["Volum [m3]"]
    matched["Gammel CO2 [kgCO2e]"] = matched["CO2 [kgCO2e]"]

    matched["Nytt volum [m3]"] = matched.apply(
        lambda row: row["Lengde [m]"] * nytt_areal_m2
        if pd.notna(row["Lengde [m]"]) and pd.notna(nytt_areal_m2)
        else row["Volum [m3]"],
        axis=1,
    )

    matched["Ny vekt [kg]"] = matched["Nytt volum [m3]"] * new_density

    if to_material in ["Limtre", "Massivtre", "Betong", "Stål"]:
        matched["Ny kostnad [kr]"] = matched["Nytt volum [m3]"] * new_price_per_m3
    else:
        matched["Ny kostnad [kr]"] = math.nan

    if to_material == "Limtre":
        matched["Ny CO2 [kgCO2e]"] = matched["Nytt volum [m3]"] * glulam_co2_factor
    elif to_material == "Massivtre":
        matched["Ny CO2 [kgCO2e]"] = matched["Nytt volum [m3]"] * clt_co2_factor
    elif to_material == "Betong":
        matched["Ny CO2 [kgCO2e]"] = matched["Nytt volum [m3]"] * concrete_co2_factor
    elif to_material == "Stål":
        matched["Ny CO2 [kgCO2e]"] = matched["Ny vekt [kg]"] * steel_co2_factor
    else:
        matched["Ny CO2 [kgCO2e]"] = math.nan

    matched["Kostnadsendring [kr]"] = matched["Ny kostnad [kr]"] - matched["Gammel kostnad [kr]"]
    matched["Vektendring [kg]"] = matched["Ny vekt [kg]"] - matched["Gammel vekt [kg]"]
    matched["CO2-endring [kgCO2e]"] = matched["Ny CO2 [kgCO2e]"] - matched["Gammel CO2 [kgCO2e]"]

    matched["Byttemetode"] = matched.apply(
        lambda row: "Lengde × nytt tverrsnitt"
        if pd.notna(row["Lengde [m]"]) and pd.notna(nytt_areal_m2)
        else "Fallback til gammelt volum",
        axis=1,
    )

    return matched


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


def make_report_summary_dict(filename, filtered_df):
    return {
        "Fil": filename,
        "Antall elementer": int(len(filtered_df)),
        "Total lengde [m]": float(pd.to_numeric(filtered_df["Lengde [m]"], errors="coerce").fillna(0).sum()),
        "Total volum [m3]": float(pd.to_numeric(filtered_df["Volum [m3]"], errors="coerce").fillna(0).sum()),
        "Total vekt [kg]": float(pd.to_numeric(filtered_df["Vekt [kg]"], errors="coerce").fillna(0).sum()),
        "Total kostnad [kr]": float(pd.to_numeric(filtered_df["Kostnad [kr]"], errors="coerce").fillna(0).sum()),
        "Total CO2 [kgCO2e]": float(pd.to_numeric(filtered_df["CO2 [kgCO2e]"], errors="coerce").fillna(0).sum()),
    }


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
        doc.add_paragraph("Før/etter-analyse av valgt materialbytte.")
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
        Prototype til hovedoppgaven – leser Excel- eller IFC-modell og gjør den om til en brukbar app for mengder, materialer, kostnader, CO₂-regnskap, materialbytte og 3D-visning.
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

    st.subheader("Priser")
    steel_price = st.number_input("Pris stål (kr/kg)", min_value=0.0, value=24.0, step=1.0)
    glulam_price = st.number_input("Pris limtre (kr/m³)", min_value=0.0, value=6500.0, step=100.0)
    clt_price_m3 = st.number_input("Pris massivtre / CLT (kr/m³)", min_value=0.0, value=7200.0, step=100.0)
    timber_price_m3 = st.number_input("Pris tre (kr/m³)", min_value=0.0, value=5000.0, step=100.0)
    concrete_price_m3 = st.number_input("Pris betong (kr/m³)", min_value=0.0, value=1800.0, step=100.0)
    unknown_price_m3 = st.number_input("Pris ukjent materiale (kr/m³)", min_value=0.0, value=1000.0, step=100.0)

    st.subheader("Tetthet")
    glulam_density = st.number_input("Tetthet limtre (kg/m³)", min_value=100.0, value=460.0, step=10.0)
    clt_density = st.number_input("Tetthet massivtre / CLT (kg/m³)", min_value=100.0, value=500.0, step=10.0)

    st.subheader("CO₂-faktorer")
    steel_co2_factor = st.number_input("Stål (kg CO₂e/kg)", min_value=0.0, value=1.90, step=0.05)
    glulam_co2_factor = st.number_input("Limtre (kg CO₂e/m³)", min_value=0.0, value=80.0, step=5.0)
    clt_co2_factor = st.number_input("Massivtre / CLT (kg CO₂e/m³)", min_value=0.0, value=110.0, step=5.0)
    timber_co2_factor = st.number_input("Tre (kg CO₂e/m³)", min_value=0.0, value=120.0, step=5.0)
    concrete_co2_factor = st.number_input("Betong (kg CO₂e/m³)", min_value=0.0, value=350.0, step=10.0)
    unknown_co2_factor = st.number_input("Ukjent materiale (kg CO₂e/m³)", min_value=0.0, value=200.0, step=10.0)

    show_raw = st.toggle("Vis rådata-tabeller", value=False)

GLULAM_DENSITY = glulam_density
CLT_DENSITY = clt_density

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
    elif DEFAULT_WORKBOOK.exists():
        filename = DEFAULT_WORKBOOK.name
        data, nodes, forside = build_dataset_from_excel(DEFAULT_WORKBOOK.read_bytes())
    else:
        st.warning("Last opp en Excel-fil eller IFC-fil i sidepanelet.")
        st.stop()
except Exception as e:
    st.error(f"Kunne ikke lese filen: {e}")
    st.stop()

a, b, c = st.columns([1.6, 1.2, 1.2])
with a:
    st.success(f"Aktiv fil: **{filename}**")
with b:
    if uploaded_excel is None and uploaded_ifc is None and DEFAULT_WORKBOOK.exists():
        st.info("Viser medfølgende eksempelmodell")
with c:
    st.write("")

for col in [
    "Segment",
    "Type",
    "Knutepunkter",
    "Material / Tverrsnitt",
    "Lengde [m]",
    "Volum [m3]",
    "Vekt [kg]",
    "materiale",
]:
    if col not in data.columns:
        data[col] = pd.NA

data["Kostnad [kr]"] = data.apply(
    lambda row: cost_for_row(
        row,
        steel_price,
        glulam_price,
        concrete_price_m3,
        clt_price_m3,
        timber_price_m3,
        unknown_price_m3,
    ),
    axis=1,
)

data["CO2 [kgCO2e]"] = data.apply(
    lambda row: co2_for_row(
        row,
        steel_co2_factor,
        glulam_co2_factor,
        concrete_co2_factor,
        clt_co2_factor,
        timber_co2_factor,
        unknown_co2_factor,
    ),
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
    filtered.groupby("materiale", dropna=False)
    .agg(
        antall=("Segment", "count"),
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

if valg == "Mengder":
    st.header("📊 Mengder")

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1:
        metric_card("Elementer", f"{len(filtered):,}".replace(",", " "))
    with k2:
        metric_card("Total lengde", f"{filtered['Lengde [m]'].sum():,.1f} m".replace(",", " "))
    with k3:
        metric_card("Stålvekt", f"{filtered.loc[filtered['materiale']=='Stål', 'Vekt [kg]'].sum():,.0f} kg".replace(",", " "))
    with k4:
        metric_card("Limtrevolum", f"{filtered.loc[filtered['materiale']=='Limtre', 'Volum [m3]'].sum():,.2f} m³".replace(",", " "))
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
                    "lengde_m": "{:.1f}",
                    "volum_m3": "{:.3f}",
                    "vekt_kg": "{:.0f}",
                    "kostnad_kr": "{:,.0f}",
                    "co2_kg": "{:,.0f}",
                }
            ),
            width="stretch",
            hide_index=True,
        )

        st.subheader("Største profiler / tverrsnitt")
        profiles = (
            filtered.groupby("Material / Tverrsnitt", dropna=False)
            .agg(
                antall=("Segment", "count"),
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
                    "lengde_m": "{:.1f}",
                    "kostnad_kr": "{:,.0f}",
                    "co2_kg": "{:,.0f}",
                }
            ),
            width="stretch",
            hide_index=True,
        )

    with right:
        st.subheader("Fordeling av kostnad")
        pie_data = summary[summary["kostnad_kr"] > 0].copy()

        if not pie_data.empty:
            pie_data["navn"] = pie_data["Type"].fillna("Ukjent") + " – " + pie_data["materiale"].fillna("Ukjent")
            fig1, ax1 = plt.subplots(figsize=(6, 5))
            ax1.pie(pie_data["kostnad_kr"], labels=pie_data["navn"], autopct="%1.1f%%", startangle=90)
            ax1.axis("equal")
            st.pyplot(fig1)
        else:
            st.info("Ingen kostnader å vise for valgt filter.")

        st.subheader("CO₂ per materiale")
        co2_data = filtered.groupby("materiale", dropna=False)["CO2 [kgCO2e]"].sum().reset_index()
        if not co2_data.empty:
            fig2, ax2 = plt.subplots(figsize=(6, 4))
            ax2.bar(co2_data["materiale"].fillna("Ukjent"), co2_data["CO2 [kgCO2e]"])
            ax2.set_ylabel("kg CO₂e")
            ax2.set_xlabel("Materiale")
            st.pyplot(fig2)
        else:
            st.info("Ingen CO₂-data å vise.")

    if all(col in filtered.columns for col in ["X1", "Y1", "X2", "Y2"]):
        st.subheader("Modellvisning i 2D (plan)")
        plot_df = filtered[["X1", "Y1", "X2", "Y2", "Type", "materiale"]].dropna()

        if not plot_df.empty:
            fig3, ax3 = plt.subplots(figsize=(10, 6))
            for row in plot_df.itertuples(index=False):
                ax3.plot([row.X1, row.X2], [row.Y1, row.Y2], linewidth=1)
            ax3.set_xlabel("X [mm]")
            ax3.set_ylabel("Y [mm]")
            ax3.set_title("Segmenter i plan")
            ax3.axis("equal")
            st.pyplot(fig3)

    st.subheader("Filtrerte elementer")
    show_cols = [
        c
        for c in [
            "Segment",
            "Type",
            "Knutepunkter",
            "Material / Tverrsnitt",
            "materiale",
            "Lengde [m]",
            "Areal [m2]",
            "Volum [m3]",
            "Vekt [kg]",
            "Kostnad [kr]",
            "CO2 [kgCO2e]",
            "IFC Type",
            "IFC GlobalId",
            "X1",
            "Y1",
            "Z1",
            "X2",
            "Y2",
            "Z2",
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
        width="stretch",
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
            st.dataframe(param_df, width="stretch", hide_index=True)
        else:
            st.info("Fant ingen tydelige prosjektparametere.")

    if show_raw:
        with st.expander("Rådata"):
            st.dataframe(data, width="stretch")

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
        filtered.groupby(["materiale", "Material / Tverrsnitt"], dropna=False)
        .agg(
            antall=("Segment", "count"),
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
                "lengde_m": "{:.1f}",
                "volum_m3": "{:.3f}",
                "vekt_kg": "{:.0f}",
                "kostnad_kr": "{:,.0f}",
                "co2_kg": "{:,.0f}",
            }
        ),
        width="stretch",
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
        st.info("Ingen materialdata å vise.")

    st.subheader("Kostnad per type")
    kost_type = filtered.groupby("Type", dropna=False)["Kostnad [kr]"].sum()
    if not kost_type.empty:
        st.bar_chart(kost_type)
    else:
        st.info("Ingen kostnadsdata å vise.")

    st.subheader("Lengde per materiale")
    lengde_mat = filtered.groupby("materiale", dropna=False)["Lengde [m]"].sum()
    if not lengde_mat.empty:
        st.bar_chart(lengde_mat)
    else:
        st.info("Ingen lengdedata å vise.")

    st.subheader("CO₂ per type")
    co2_type = filtered.groupby("Type", dropna=False)["CO2 [kgCO2e]"].sum()
    if not co2_type.empty:
        st.bar_chart(co2_type)
    else:
        st.info("Ingen CO₂-data å vise.")

elif valg == "Materialbytte":
    st.header("🔁 Materialbytte")
    st.info(
        "Denne siden gjør en kalkylemessig erstatning av elementer. "
        "Den skriver ikke tilbake til IFC ennå, og den verifiserer ikke bæreevne eller brannkrav."
    )

    if data.empty:
        st.warning("Ingen data tilgjengelig.")
        st.stop()

    col1, col2, col3 = st.columns(3)

    with col1:
        available_types = sorted([x for x in data["Type"].dropna().unique().tolist()])
        if not available_types:
            st.warning("Fant ingen elementtyper i datasettet.")
            st.stop()
        default_type = "Søyle" if "Søyle" in available_types else available_types[0]
        selected_swap_type = st.selectbox(
            "Elementtype som skal byttes",
            available_types,
            index=available_types.index(default_type),
        )

    with col2:
        available_materials = sorted([x for x in data["materiale"].dropna().unique().tolist()])
        if not available_materials:
            st.warning("Fant ingen materialer i datasettet.")
            st.stop()
        default_from = "Stål" if "Stål" in available_materials else available_materials[0]
        from_material = st.selectbox(
            "Nåværende materiale",
            available_materials,
            index=available_materials.index(default_from),
        )

    with col3:
        to_material = st.selectbox("Nytt materiale", ["Limtre", "Massivtre", "Betong", "Stål"], index=0)

    col4, col5, col6 = st.columns(3)

    with col4:
        new_profile_text = st.text_input("Nytt tverrsnitt", value="115x360")

    with col5:
        if to_material == "Limtre":
            new_density = st.number_input("Ny tetthet (kg/m³)", min_value=100.0, value=460.0, step=10.0, key="swap_density")
        elif to_material == "Massivtre":
            new_density = st.number_input("Ny tetthet (kg/m³)", min_value=100.0, value=500.0, step=10.0, key="swap_density")
        elif to_material == "Betong":
            new_density = st.number_input("Ny tetthet (kg/m³)", min_value=500.0, value=2400.0, step=50.0, key="swap_density")
        else:
            new_density = st.number_input("Ny tetthet (kg/m³)", min_value=1000.0, value=7850.0, step=50.0, key="swap_density")

    with col6:
        if to_material == "Limtre":
            new_price_per_m3 = st.number_input("Ny pris (kr/m³)", min_value=0.0, value=6500.0, step=100.0, key="swap_price")
        elif to_material == "Massivtre":
            new_price_per_m3 = st.number_input("Ny pris (kr/m³)", min_value=0.0, value=7200.0, step=100.0, key="swap_price")
        elif to_material == "Betong":
            new_price_per_m3 = st.number_input("Ny pris (kr/m³)", min_value=0.0, value=1800.0, step=100.0, key="swap_price")
        else:
            new_price_per_m3 = st.number_input("Ny pris (kr/m³)", min_value=0.0, value=15000.0, step=500.0, key="swap_price")

    swap_df = calculate_material_swap(
        source_df=data,
        selected_type=selected_swap_type,
        from_material=from_material,
        to_material=to_material,
        new_profile_text=new_profile_text,
        new_density=new_density,
        new_price_per_m3=new_price_per_m3,
        steel_co2_factor=steel_co2_factor,
        glulam_co2_factor=glulam_co2_factor,
        concrete_co2_factor=concrete_co2_factor,
        clt_co2_factor=clt_co2_factor,
    )

    if swap_df.empty:
        st.warning("Fant ingen elementer som matcher valgt elementtype og materiale.")
    else:
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
                "Material / Tverrsnitt",
                "Nytt tverrsnitt",
                "Lengde [m]",
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
            ]
            if c in swap_df.columns
        ]

        st.dataframe(
            swap_df[show_swap_cols].style.format(
                {
                    "Lengde [m]": "{:.2f}",
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
                }
            ),
            width="stretch",
            hide_index=True,
        )

        swap_csv = swap_df[show_swap_cols].to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Last ned endringsliste som CSV",
            swap_csv,
            file_name="materialbytte.csv",
            mime="text/csv",
        )

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
        filtered.groupby("materiale", dropna=False)
        .agg(
            antall=("Segment", "count"),
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
                "volum_m3": "{:.3f}",
                "vekt_kg": "{:.0f}",
                "co2_kg": "{:,.0f}",
                "kostnad_kr": "{:,.0f}",
            }
        ),
        width="stretch",
        hide_index=True,
    )

    left, right = st.columns(2)

    with left:
        if not co2_material.empty:
            fig4, ax4 = plt.subplots(figsize=(6, 4))
            ax4.bar(co2_material["materiale"].fillna("Ukjent"), co2_material["co2_kg"])
            ax4.set_ylabel("kg CO₂e")
            ax4.set_xlabel("Materiale")
            ax4.set_title("CO₂ per materiale")
            st.pyplot(fig4)

    with right:
        if not co2_material.empty:
            fig5, ax5 = plt.subplots(figsize=(6, 4))
            ax5.bar(co2_material["materiale"].fillna("Ukjent"), co2_material["kostnad_kr"])
            ax5.set_ylabel("kr")
            ax5.set_xlabel("Materiale")
            ax5.set_title("Kostnad per materiale")
            st.pyplot(fig5)

elif valg == "3D-modell":
    st.header("🧊 3D-modellvisning")

    if uploaded_ifc is None:
        st.info("3D-visning er tilgjengelig når du laster opp en IFC-fil.")
        st.stop()

    st.markdown(
        "Denne visningen viser IFC-elementene som er lest inn i appen. "
        "Du kan velge å vise kun filtrerte elementer, slik at du ser hva som faktisk er med i beregningen."
    )

    left, right = st.columns(2)

    with left:
        visning = st.radio(
            "Velg visning",
            ["Kun filtrerte elementer", "Alle elementer"],
            horizontal=True,
        )

    with right:
        max_elements_3d = st.slider(
            "Maks antall elementer i 3D-visning",
            min_value=50,
            max_value=1000,
            value=300,
            step=50,
        )

    try:
        meshes = extract_ifc_meshes(uploaded_ifc.getvalue())
    except Exception as e:
        st.error(f"Kunne ikke lage 3D-visning: {e}")
        st.stop()

    if not meshes:
        st.warning("Fant ingen 3D-geometri i IFC-filen.")
        st.stop()

    if visning == "Kun filtrerte elementer":
        visible_ids = set(filtered["IFC GlobalId"].dropna().astype(str).tolist())
    else:
        visible_ids = None

    fig3d = build_ifc_3d_figure(
        meshes=meshes,
        visible_ids=visible_ids,
        max_elements=max_elements_3d,
    )

    st.plotly_chart(fig3d, use_container_width=True)

    st.subheader("Forklaring")
    legend_df = pd.DataFrame(
        {
            "Materiale": ["Stål", "Betong", "Limtre", "Massivtre", "Tre", "Ukjent"],
            "Farge": ["Blå", "Grå", "Brun", "Grønn", "Trebrun", "Lys grå"],
        }
    )
    st.dataframe(legend_df, width="stretch", hide_index=True)

    if visning == "Kun filtrerte elementer":
        st.success(
            f"Viser {min(len(set(filtered['IFC GlobalId'].dropna())), max_elements_3d)} filtrerte elementer i 3D."
        )
    else:
        st.info(f"Viser opptil {max_elements_3d} elementer fra IFC-modellen.")

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
    st.dataframe(report_df, width="stretch", hide_index=True)

    st.subheader("Materialoversikt")
    export_material_summary = material_summary.copy()
    st.dataframe(
        export_material_summary.style.format(
            {
                "lengde_m": "{:.1f}",
                "volum_m3": "{:.3f}",
                "vekt_kg": "{:.0f}",
                "kostnad_kr": "{:,.0f}",
                "co2_kg": "{:,.0f}",
            }
        ),
        width="stretch",
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
                    "Gammel kostnad [kr]",
                    "Ny kostnad [kr]",
                    "Gammel CO2 [kgCO2e]",
                    "Ny CO2 [kgCO2e]",
                    "Kostnadsendring [kr]",
                    "CO2-endring [kgCO2e]",
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
            st.info("Installer python-docx for Word-eksport.")

    with col_b:
        if pdf_bytes is not None:
            st.download_button(
                "Last ned rapport som PDF",
                data=pdf_bytes,
                file_name="byggtotal_rapport.pdf",
                mime="application/pdf",
            )
        else:
            st.info("Installer reportlab for PDF-eksport.")

st.markdown("---")
st.markdown(
    "**Neste steg for hovedoppgaven:** legg til IFC-geometri, NS3420-koder, EPD-baserte CO₂-faktorer og eksport av endringsrapport per fag.**"
)