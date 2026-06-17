# geometry.py
# Geometri- og terrengfunksjoner for byggTotal.
# Håndterer konveks hull, terrengberegning, plotting og grunnprisberegning.

import math

import matplotlib.pyplot as plt
import pandas as pd


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


def generate_ground_obj(points_df: pd.DataFrame) -> bytes:
    """Genererer en OBJ-terrengmodell fra stikningspunkter."""
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
