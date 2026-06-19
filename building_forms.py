"""
building_forms.py – Fri bygningsform for byggTotal Bygggenerator.

Støtter:
  • Regulært polygon (trekant, pentagon, hexagon, oktagon, n-kant)
  • Oval / Ellipse
  • Sirkel
  • Frihånd polygon (brukeren oppgir XY-hjørner)

Leverer:
  • generate_free_form_vertices()  → liste av (x, y) i meter
  • free_form_frame_export()       → DataFrame med søyler + bjelker per etasje
  • free_form_slab_export()        → DataFrame med dekker per etasje
  • plot_free_form_plan()          → matplotlib-figur
  • free_form_area()               → float areal m²
  • free_form_perimeter()          → float omkrets m
"""
from __future__ import annotations

import math
from typing import Sequence

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import pandas as pd


# ── Geometri-generatorer ──────────────────────────────────────────────────────

def _regular_polygon(n: int, radius: float, rotation_deg: float = 0.0) -> list[tuple[float, float]]:
    """Regulært n-kant sentrert i (radius, radius)."""
    pts = []
    for i in range(n):
        angle = math.radians(rotation_deg + 360 * i / n)
        pts.append((radius + radius * math.cos(angle),
                    radius + radius * math.sin(angle)))
    return pts


def _ellipse(a: float, b: float, n_pts: int = 64) -> list[tuple[float, float]]:
    """Ellipse med halvakser a (X) og b (Y), sentrert i (a, b)."""
    pts = []
    for i in range(n_pts):
        angle = 2 * math.pi * i / n_pts
        pts.append((a + a * math.cos(angle),
                    b + b * math.sin(angle)))
    return pts


def _circle(radius: float, n_pts: int = 64) -> list[tuple[float, float]]:
    return _ellipse(radius, radius, n_pts)


def generate_free_form_vertices(
    shape: str,
    params: dict,
) -> list[tuple[float, float]]:
    """
    Returnerer en liste med (x, y) polygon-hjørner i meter.

    shape:
      'polygon_n'   – regulært n-kant
      'ellipse'     – ellipse
      'circle'      – sirkel
      'freehand'    – frihånd (brukeren oppgir hjørner)
    """
    if shape == "polygon_n":
        n      = max(3, int(params.get("poly_n_sides", 6)))
        radius = max(0.5, float(params.get("poly_radius_m", 10.0)))
        rot    = float(params.get("poly_rotation_deg", 0.0))
        return _regular_polygon(n, radius, rot)

    elif shape == "ellipse":
        a = max(1.0, float(params.get("ellipse_a_m", 15.0)))
        b = max(1.0, float(params.get("ellipse_b_m", 10.0)))
        n = max(12, int(params.get("ellipse_segments", 48)))
        return _ellipse(a, b, n)

    elif shape == "circle":
        r = max(1.0, float(params.get("circle_r_m", 10.0)))
        n = max(12, int(params.get("circle_segments", 48)))
        return _circle(r, n)

    elif shape == "freehand":
        raw = params.get("freehand_vertices", [])
        pts = []
        for row in raw:
            try:
                x = float(row.get("X [m]", 0))
                y = float(row.get("Y [m]", 0))
                pts.append((x, y))
            except Exception:
                continue
        if len(pts) < 3:
            # Fallback: 6-kant
            return _regular_polygon(6, 10.0)
        return pts

    else:
        return _regular_polygon(4, 10.0)


# ── Geometri-hjelpere ─────────────────────────────────────────────────────────

def free_form_area(vertices: list[tuple[float, float]]) -> float:
    """Shoelace-formel."""
    n = len(vertices)
    if n < 3:
        return 0.0
    area = 0.0
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        area += x1 * y2 - x2 * y1
    return abs(area) / 2.0


def free_form_perimeter(vertices: list[tuple[float, float]]) -> float:
    n = len(vertices)
    if n < 2:
        return 0.0
    total = 0.0
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        total += math.hypot(x2 - x1, y2 - y1)
    return total


def _bbox(vertices):
    xs = [v[0] for v in vertices]
    ys = [v[1] for v in vertices]
    return min(xs), max(xs), min(ys), max(ys)


# ── Polygon-geometri-hjelp ────────────────────────────────────────────────────

def _point_in_polygon(x: float, y: float, vertices: list[tuple[float, float]]) -> bool:
    """Ray-casting punkt-i-polygon test."""
    n = len(vertices)
    inside = False
    x1, y1 = vertices[-1]
    for x2, y2 in vertices:
        if ((y1 > y) != (y2 > y)) and (x < (x2 - x1) * (y - y1) / (y2 - y1) + x1):
            inside = not inside
        x1, y1 = x2, y2
    return inside


def _scanline_x(vertices: list[tuple[float, float]], y: float) -> list[float]:
    """Sorterte x-kryss der horisontal linje y krysser polygonkantene."""
    xs = []
    n = len(vertices)
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        if (y1 <= y < y2) or (y2 <= y < y1):
            xs.append(x1 + (y - y1) * (x2 - x1) / (y2 - y1))
    xs.sort()
    return xs


def _scanline_y(vertices: list[tuple[float, float]], x: float) -> list[float]:
    """Sorterte y-kryss der vertikal linje x krysser polygonkantene."""
    ys = []
    n = len(vertices)
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        if abs(x2 - x1) < 1e-9:
            continue
        if (x1 <= x < x2) or (x2 <= x < x1):
            ys.append(y1 + (x - x1) * (y2 - y1) / (x2 - x1))
    ys.sort()
    return ys


def _grid_lines(
    vertices: list[tuple[float, float]],
    max_span_m: float,
) -> tuple[
    list[tuple[tuple[float, float], tuple[float, float]]],
    list[tuple[tuple[float, float], tuple[float, float]]],
    list[tuple[float, float]],
]:
    """
    Bygg innvendig strukturgrid klippet mot polygonen.
    Returnerer (x_beams, y_beams, col_points):
      x_beams  – horisontale bjelkestrekk inne i polygon
      y_beams  – vertikale bjelkestrekk inne i polygon
      col_pts  – skjæringspunkter i rutenettet som er inne i polygonen
    """
    coord_x = [v[0] for v in vertices]
    coord_y = [v[1] for v in vertices]
    x_min, x_max = min(coord_x), max(coord_x)
    y_min, y_max = min(coord_y), max(coord_y)

    # Gridlinjer i begge retninger (starter max_span inn fra kanten)
    grid_ys = []
    y = y_min + max_span_m
    while y < y_max - 1e-6:
        grid_ys.append(y)
        y += max_span_m

    grid_xs = []
    x = x_min + max_span_m
    while x < x_max - 1e-6:
        grid_xs.append(x)
        x += max_span_m

    # Horisontale bjelker (langs X) ved hvert grid_y
    x_beams: list[tuple[tuple[float, float], tuple[float, float]]] = []
    for gy in grid_ys:
        xs = _scanline_x(vertices, gy)
        for j in range(0, len(xs) - 1, 2):
            xa, xb = xs[j], xs[j + 1]
            if xb - xa > 0.05:
                x_beams.append(((xa, gy), (xb, gy)))

    # Vertikale bjelker (langs Y) ved hvert grid_x
    y_beams: list[tuple[tuple[float, float], tuple[float, float]]] = []
    for gx in grid_xs:
        ys = _scanline_y(vertices, gx)
        for j in range(0, len(ys) - 1, 2):
            ya, yb = ys[j], ys[j + 1]
            if yb - ya > 0.05:
                y_beams.append(((gx, ya), (gx, yb)))

    # Søyler ved alle skjæringspunkter i rutenettet som er inne i polygonen
    col_pts: list[tuple[float, float]] = []
    for gx in grid_xs:
        for gy in grid_ys:
            if _point_in_polygon(gx, gy, vertices):
                col_pts.append((gx, gy))

    return x_beams, y_beams, col_pts


# ── Bæresystem fra polygon ────────────────────────────────────────────────────

def _edge_column_points(
    vertices: list[tuple[float, float]],
    max_span_m: float,
) -> list[list[tuple[float, float]]]:
    """
    For hver kant i polygonen: returner en liste med (x,y)-punkter langs kanten
    inkl. start- og sluttpunktet, med maksimalt `max_span_m` mellom søylene.

    Eksempel: kant 8.5 m, max_span 6 m → 2 segmenter à 4.25 m → 3 punkter.
    """
    n = len(vertices)
    edge_points = []
    for i in range(n):
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        L = math.hypot(x2 - x1, y2 - y1)
        if L < 0.01:
            edge_points.append([(x1, y1), (x2, y2)])
            continue
        n_segs = max(1, math.ceil(L / max_span_m))
        pts = []
        for k in range(n_segs + 1):
            t = k / n_segs
            pts.append((x1 + t * (x2 - x1), y1 + t * (y2 - y1)))
        edge_points.append(pts)
    return edge_points


def free_form_frame_export(
    vertices: list[tuple[float, float]],
    n_levels: int,
    floor_h_m: float,
    col_mat: str,
    col_qual: str,
    col_prof: str,
    beam_mat: str,
    beam_qual: str,
    beam_prof: str,
    max_span_m: float = 6.0,
    col_cost_per_m: float = 8500.0,
    col_co2_per_m3: float = 7200.0,
    beam_cost_per_m: float = 6200.0,
    beam_co2_per_m3: float = 6800.0,
    interior_beam_mat: str | None = None,
    interior_beam_qual: str | None = None,
    interior_beam_prof: str | None = None,
) -> pd.DataFrame:
    """
    Genererer søyler og bjelker langs polygon-kantene per etasje.
    Bjelkespenn begrenses til max_span_m (standard 6 m) ved å sette inn
    mellomliggende søyler der kanten er lengre enn max_span_m.
    """
    # Bygg kant-punkter og innvendig strukturgrid
    edge_pts                          = _edge_column_points(vertices, max_span_m)
    x_beams, y_beams, grid_col_pts   = _grid_lines(vertices, max_span_m)
    interior_lines                    = x_beams + y_beams

    # Samle unike søyleposisjoner (alle punkt fra alle kanter)
    seen_cols: set[tuple[float, float]] = set()
    all_col_positions: list[tuple[float, float]] = []
    for pts in edge_pts:
        for p in pts:
            key = (round(p[0], 4), round(p[1], 4))
            if key not in seen_cols:
                seen_cols.add(key)
                all_col_positions.append(p)

    def _prof_area_m2(prof: str) -> float:
        if "HEB" in prof or "HEA" in prof:
            size = "".join(filter(str.isdigit, prof))
            return (float(size) / 1000) ** 2 * 0.25 if size else 0.02
        try:
            parts = prof.replace("mm", "").split("x")
            if len(parts) == 2:
                return (float(parts[0]) / 1000) * (float(parts[1]) / 1000)
        except Exception:
            pass
        return 0.02

    # Innvendige bjelker kan ha separat profil (f.eks. HSQ ved hulldekke)
    int_b_mat  = interior_beam_mat  or beam_mat
    int_b_qual = interior_beam_qual or beam_qual
    int_b_prof = interior_beam_prof or beam_prof

    col_area      = _prof_area_m2(col_prof)
    beam_area     = _prof_area_m2(beam_prof)
    int_beam_area = _prof_area_m2(int_b_prof)

    rows = []
    col_id = 1
    beam_id = 1

    def _col_row(seg_id, x, y, z_bot, z_top, level):
        length = z_top - z_bot
        vol = col_area * length
        return {
            "ID": seg_id, "Segment": seg_id, "Type": "Søyle", "Nivå": level,
            "X1 [m]": round(x, 4), "Y1 [m]": round(y, 4), "Z1 [m]": round(z_bot, 4),
            "X2 [m]": round(x, 4), "Y2 [m]": round(y, 4), "Z2 [m]": round(z_top, 4),
            "Knutepunkter": f"({x:.2f},{y:.2f},{z_bot:.2f})-({x:.2f},{y:.2f},{z_top:.2f})",
            "Material / Tverrsnitt": f"{col_mat} {col_qual} / {col_prof}",
            "Lengde [m]": round(length, 3), "Areal [m2]": round(col_area, 5),
            "Volum [m3]": round(vol, 4), "Vekt [kg]": round(vol * _density(col_mat), 1),
            "materiale": col_mat, "Materialkvalitet": col_qual,
            "Mengdegrunnlag": "Fri form", "Endret IFC": False,
            "Kostnad [kr]": round(length * col_cost_per_m, 0),
            "CO2 [kgCO2e]": round(vol * col_co2_per_m3, 1),
        }

    def _beam_row(seg_id, typ, ax_, ay_, bx_, by_, z_lev, level):
        seg_len = math.hypot(bx_ - ax_, by_ - ay_)
        vol = beam_area * seg_len
        return {
            "ID": seg_id, "Segment": seg_id, "Type": typ, "Nivå": level,
            "X1 [m]": round(ax_, 4), "Y1 [m]": round(ay_, 4), "Z1 [m]": round(z_lev, 4),
            "X2 [m]": round(bx_, 4), "Y2 [m]": round(by_, 4), "Z2 [m]": round(z_lev, 4),
            "Knutepunkter": f"({ax_:.2f},{ay_:.2f},{z_lev:.2f})-({bx_:.2f},{by_:.2f},{z_lev:.2f})",
            "Material / Tverrsnitt": f"{beam_mat} {beam_qual} / {beam_prof}",
            "Lengde [m]": round(seg_len, 3), "Areal [m2]": round(beam_area, 5),
            "Volum [m3]": round(vol, 4), "Vekt [kg]": round(vol * _density(beam_mat), 1),
            "materiale": beam_mat, "Materialkvalitet": beam_qual,
            "Mengdegrunnlag": "Fri form", "Endret IFC": False,
            "Kostnad [kr]": round(seg_len * beam_cost_per_m, 0),
            "CO2 [kgCO2e]": round(vol * beam_co2_per_m3, 1),
        }

    for level in range(1, n_levels + 1):
        z0 = (level - 1) * floor_h_m
        z1 = level * floor_h_m

        # Innvendige grid-søyler
        for x, y in grid_col_pts:
            rows.append(_col_row(f"SI{col_id}", x, y, z0, z1, level))
            col_id += 1

        # Perimetre-søyler langs kantene
        for x, y in all_col_positions:
            rows.append(_col_row(f"S{col_id}", x, y, z0, z1, level))
            col_id += 1

        # Bjelker mellom nabokolonner langs kantene
        for pts in edge_pts:
            for j in range(len(pts) - 1):
                ax_, ay_ = pts[j]
                bx_, by_ = pts[j + 1]
                if math.hypot(bx_ - ax_, by_ - ay_) < 0.01:
                    continue
                rows.append(_beam_row(f"B{beam_id}", "Bjelke", ax_, ay_, bx_, by_, z1, level))
                beam_id += 1

        # Innvendige bjelker for å holde dekke-spenn ≤ max_span_m
        for (ax_, ay_), (bx_, by_) in interior_lines:
            if math.hypot(bx_ - ax_, by_ - ay_) < 0.05:
                continue
            seg_len = math.hypot(bx_ - ax_, by_ - ay_)
            vol = int_beam_area * seg_len
            rows.append({
                "ID": f"BI{beam_id}", "Segment": f"BI{beam_id}", "Type": "Innv. bjelke", "Nivå": level,
                "X1 [m]": round(ax_, 4), "Y1 [m]": round(ay_, 4), "Z1 [m]": round(z1, 4),
                "X2 [m]": round(bx_, 4), "Y2 [m]": round(by_, 4), "Z2 [m]": round(z1, 4),
                "Knutepunkter": f"({ax_:.2f},{ay_:.2f},{z1:.2f})-({bx_:.2f},{by_:.2f},{z1:.2f})",
                "Material / Tverrsnitt": f"{int_b_mat} {int_b_qual} / {int_b_prof}",
                "Lengde [m]": round(seg_len, 3), "Areal [m2]": round(int_beam_area, 5),
                "Volum [m3]": round(vol, 4), "Vekt [kg]": round(vol * _density(int_b_mat), 1),
                "materiale": int_b_mat, "Materialkvalitet": int_b_qual,
                "Mengdegrunnlag": "Fri form", "Endret IFC": False,
                "Kostnad [kr]": round(seg_len * beam_cost_per_m, 0),
                "CO2 [kgCO2e]": round(vol * beam_co2_per_m3, 1),
            })
            beam_id += 1

    return pd.DataFrame(rows)


def free_form_slab_export(
    vertices: list[tuple[float, float]],
    n_levels: int,
    floor_h_m: float,
    slab_thk_m: float,
    slab_mat: str,
    slab_qual: str,
    cost_per_m2: float = 1650.0,
    co2_per_m3: float = 380.0,
) -> pd.DataFrame:
    """Ett dekke per etasje = polygon-arealet. Inkluderer polygon-geometri for IFC-eksport."""
    import json as _json
    area = free_form_area(vertices)
    # Polygon-koordinater som JSON for presis IFC-geometri
    poly_json = _json.dumps([[round(x, 4), round(y, 4)] for x, y in vertices])

    rows = []
    for level in range(1, n_levels + 1):
        z_mm = int(round(level * floor_h_m * 1000))
        vol    = area * slab_thk_m
        weight = vol * _density(slab_mat)
        cost   = area * cost_per_m2
        co2    = vol * co2_per_m3
        rows.append({
            "DeckID": f"D{level}",
            "Type": "Dekke",
            "Nivå": level,
            # Polygon-form for IFC-eksport
            "poly_pts_json": poly_json,
            "Z [mm]": z_mm,
            "Knutepunkter": f"z={level*floor_h_m:.2f}m",
            "Material / Tverrsnitt": f"{slab_mat} {slab_qual} / t={round(slab_thk_m*1000)} mm",
            "Lengde [m]": float("nan"),
            "Areal [m2]": round(area, 2),
            "Volum [m3]": round(vol, 3),
            "Vekt [kg]": round(weight, 1),
            "materiale": slab_mat,
            "Materialkvalitet": slab_qual,
            "Mengdegrunnlag": "Fri form",
            "Endret IFC": False,
            "Kostnad [kr]": round(cost, 0),
            "CO2 [kgCO2e]": round(co2, 1),
        })
    return pd.DataFrame(rows)


def _density(material: str) -> float:
    return {"Stål": 7850.0, "Betong": 2400.0, "Limtre": 480.0,
            "Massivtre": 500.0, "CLT": 500.0}.get(material, 2400.0)


# ── Plotting ──────────────────────────────────────────────────────────────────

def plot_free_form_plan(
    vertices: list[tuple[float, float]],
    shape_label: str = "",
    col_indices: list[int] | None = None,
    edge_col_pts: list[list[tuple[float, float]]] | None = None,
    interior_lines: list[tuple[tuple[float, float], tuple[float, float]]] | None = None,
    grid_col_pts: list[tuple[float, float]] | None = None,
) -> plt.Figure:
    """2D plantegning med polygon-omriss, mål på alle kanter og søylemarkering."""
    fig, ax = plt.subplots(figsize=(9, 7))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("#f8f9fb")

    n = len(vertices)
    xs = [v[0] for v in vertices] + [vertices[0][0]]
    ys = [v[1] for v in vertices] + [vertices[0][1]]

    # Fylte polygon
    ax.fill(xs[:-1], ys[:-1], color="#ddeeff", alpha=0.7, zorder=1)
    ax.plot(xs, ys, color="#1a3a5c", linewidth=2.0, zorder=2)

    # Innvendige bjelkelinjer (stiplede) – begge retninger
    if interior_lines:
        for (x_a, y_a), (x_b, y_b) in interior_lines:
            ax.plot([x_a, x_b], [y_a, y_b],
                    color="#4a90d9", linewidth=1.1,
                    linestyle="--", alpha=0.8, zorder=3)

    # Innvendige grid-søyler
    if grid_col_pts:
        gx_arr = [p[0] for p in grid_col_pts]
        gy_arr = [p[1] for p in grid_col_pts]
        ax.scatter(gx_arr, gy_arr, marker="s", s=28,
                   color="#1a3a5c", zorder=7, label="Søyle (innv.)")

    # Kant-mål – vis bare for polygon med få sider, ellers bare total omkrets
    _kant_iter = range(n) if n <= 24 else []
    for i in _kant_iter:
        x1, y1 = vertices[i]
        x2, y2 = vertices[(i + 1) % n]
        L = math.hypot(x2 - x1, y2 - y1)
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        angle  = math.degrees(math.atan2(y2 - y1, x2 - x1))
        # Normaliser tekst-rotasjon
        if angle > 90:
            angle -= 180
        elif angle < -90:
            angle += 180
        # Perpendikulær offset (utover)
        dx = (x2 - x1) / L; dy = (y2 - y1) / L
        nx, ny = -dy, dx
        cx = [v[0] for v in vertices]
        cy = [v[1] for v in vertices]
        center_x, center_y = sum(cx)/n, sum(cy)/n
        # Sjekk retning (vekk fra sentrum)
        if nx * (mx - center_x) + ny * (my - center_y) < 0:
            nx, ny = -nx, -ny
        off = max(free_form_perimeter(vertices) * 0.04, 1.5)
        ax.text(mx + nx * off, my + ny * off,
                f"{L:.1f} m",
                ha="center", va="center", fontsize=7.5,
                rotation=angle, color="#1a3a5c",
                bbox=dict(fc="white", ec="none", pad=0.4))

    # Hjørne-vinkler – bare for polygon med få sider (<=16), ellers blir det for tett
    show_angles = n <= 16
    if show_angles:
        for i in range(n):
            pc = np.array(vertices[i])
            pm = np.array(vertices[(i - 1) % n])
            pn = np.array(vertices[(i + 1) % n])
            v1 = pm - pc; v2 = pn - pc
            L1 = np.linalg.norm(v1); L2 = np.linalg.norm(v2)
            if L1 < 0.01 or L2 < 0.01:
                continue
            dot = float(np.clip(np.dot(v1/L1, v2/L2), -1.0, 1.0))
            ang = math.degrees(math.acos(dot))
            a1 = math.degrees(math.atan2(float(v1[1]/L1), float(v1[0]/L1)))
            a2 = math.degrees(math.atan2(float(v2[1]/L2), float(v2[0]/L2)))
            # Sikre at begge verdier er endelige tall
            if not (math.isfinite(a1) and math.isfinite(a2) and math.isfinite(ang)):
                continue
            t1, t2 = sorted([a1, a2])
            if t2 - t1 > 180:
                t1, t2 = t2, t1 + 360
            # Unngå degenerate arc (theta1 == theta2)
            if abs(t2 - t1) < 0.1:
                continue
            r = min(max(free_form_perimeter(vertices) / n * 0.15, 0.8), 3.0)
            try:
                from matplotlib.patches import Arc
                ax.add_patch(Arc(tuple(pc.tolist()), r * 2, r * 2, angle=0,
                                 theta1=float(t1), theta2=float(t2),
                                 color="#c0392b", lw=0.9, zorder=5))
            except Exception:
                pass
            a_mid = math.radians((t1 + t2) / 2)
            ax.text(pc[0] + (r + max(r * 0.6, 0.8)) * math.cos(a_mid),
                    pc[1] + (r + max(r * 0.6, 0.8)) * math.sin(a_mid),
                    f"{ang:.1f}°",
                    ha="center", va="center", fontsize=6.5,
                    color="#c0392b", fontweight="bold",
                    bbox=dict(fc="white", ec="none", pad=0.2))

    # Søyle-markeringer (bruker edge_col_pts hvis tilgjengelig)
    if edge_col_pts is not None:
        seen_sq: set[tuple[float, float]] = set()
        for pts in edge_col_pts:
            for x, y in pts:
                key = (round(x, 3), round(y, 3))
                if key not in seen_sq:
                    seen_sq.add(key)
                    ax.plot(x, y, "s", color="#1a3a5c", markersize=6, zorder=6)
    elif col_indices is not None:
        for ci in col_indices:
            x, y = vertices[ci]
            ax.plot(x, y, "s", color="#1a3a5c", markersize=6, zorder=6)
    else:
        for x, y in vertices:
            ax.plot(x, y, "s", color="#1a3a5c", markersize=6, zorder=6)

    # Hjørnepunkt-numre (bare for polygon med få sider)
    if n <= 24:
        for i, (x, y) in enumerate(vertices):
            ax.text(x, y, f" {i+1}", fontsize=5.5, color="#555",
                    va="bottom", zorder=7)

    # Areal og omkrets
    A = free_form_area(vertices)
    P = free_form_perimeter(vertices)
    ax.text(0.01, 0.99,
            f"Areal: {A:,.1f} m²\nOmkrets: {P:,.1f} m\n{shape_label}",
            transform=ax.transAxes, va="top", ha="left",
            fontsize=8, color="#1a3a5c",
            bbox=dict(fc="white", ec="#1a3a5c", pad=5, linewidth=0.7))

    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title("Plantegning – fri form", fontsize=11, fontweight="bold",
                 color="#1a3a5c", pad=10)
    fig.tight_layout()
    return fig


def plot_free_form_3d(
    vertices: list[tuple[float, float]],
    n_levels: int,
    floor_h_m: float,
    frame_df: pd.DataFrame | None = None,
    edge_col_pts: list[list[tuple[float, float]]] | None = None,
    grid_col_pts: list[tuple[float, float]] | None = None,
    interior_lines: list[tuple[tuple[float, float], tuple[float, float]]] | None = None,
) -> "go.Figure":
    """Enkel 3D Plotly-modell: etasje-omriss stablet oppover."""
    import plotly.graph_objects as go

    fig = go.Figure()
    colors = ["#4a90d9", "#5ba35e", "#e67e22", "#9b59b6",
              "#e74c3c", "#1abc9c", "#f39c12", "#2c3e50"]

    for level in range(n_levels + 1):
        z = level * floor_h_m
        xs = [v[0] for v in vertices] + [vertices[0][0]]
        ys = [v[1] for v in vertices] + [vertices[0][1]]
        zs = [z] * len(xs)
        col = colors[level % len(colors)]
        name = "Fundament" if level == 0 else f"Etasje {level} – z={z:.1f}m"
        fig.add_trace(go.Scatter3d(
            x=xs, y=ys, z=zs,
            mode="lines",
            line=dict(color=col, width=5),
            name=name,
        ))
        # Fylte dekke-polygon (triangulert via fan fra sentrum)
        if level > 0:
            cx = sum(v[0] for v in vertices) / len(vertices)
            cy = sum(v[1] for v in vertices) / len(vertices)
            x_tri = [cx] + [v[0] for v in vertices]
            y_tri = [cy] + [v[1] for v in vertices]
            z_tri = [z]  * (len(vertices) + 1)
            i_idx = [0] * len(vertices)
            j_idx = list(range(1, len(vertices) + 1))
            k_idx = list(range(2, len(vertices) + 1)) + [1]
            fig.add_trace(go.Mesh3d(
                x=x_tri, y=y_tri, z=z_tri,
                i=i_idx, j=j_idx, k=k_idx,
                color=col, opacity=0.18,
                showlegend=False,
            ))

    # Vertikale søylelinjer – hent unike posisjoner fra edge_col_pts
    if edge_col_pts is not None:
        seen_col3d: set[tuple[float, float]] = set()
        col_positions_3d = []
        for pts in edge_col_pts:
            for p in pts:
                key = (round(p[0], 3), round(p[1], 3))
                if key not in seen_col3d:
                    seen_col3d.add(key)
                    col_positions_3d.append(p)
    else:
        n_v = len(vertices)
        perim = free_form_perimeter(vertices)
        step  = max(1, round(perim / (6.0 * n_v)) if n_v > 12 else 1)
        col_positions_3d = [vertices[ci] for ci in range(0, n_v, step)]

    z_top = n_levels * floor_h_m

    for x, y in col_positions_3d:
        fig.add_trace(go.Scatter3d(
            x=[x, x], y=[y, y], z=[0, z_top],
            mode="lines", line=dict(color="#1a3a5c", width=3),
            showlegend=False,
        ))

    # Innvendige grid-søyler
    if grid_col_pts:
        for x, y in grid_col_pts:
            fig.add_trace(go.Scatter3d(
                x=[x, x], y=[y, y], z=[0, z_top],
                mode="lines", line=dict(color="#1a3a5c", width=3),
                showlegend=False,
            ))

    # Innvendige bjelker per etasje
    if interior_lines:
        for level in range(1, n_levels + 1):
            z = level * floor_h_m
            for (x_a, y_a), (x_b, y_b) in interior_lines:
                fig.add_trace(go.Scatter3d(
                    x=[x_a, x_b], y=[y_a, y_b], z=[z, z],
                    mode="lines", line=dict(color="#4a90d9", width=2),
                    showlegend=False,
                ))

    fig.update_layout(
        title=f"3D – fri form  ({n_levels} etasjer)",
        height=600,
        margin=dict(l=0, r=0, t=40, b=0),
        scene=dict(
            xaxis_title="X [m]", yaxis_title="Y [m]", zaxis_title="Z [m]",
            aspectmode="data",
            camera=dict(eye=dict(x=1.5, y=-1.8, z=1.1)),
            bgcolor="#f7f9fc",
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, x=0),
    )
    return fig
