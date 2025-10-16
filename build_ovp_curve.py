#!/usr/bin/env python3
"""
Outputs ONE worksheet 'curve_data' with columns:
A x        (dense X grid)
B y        (smoothed curve)
C-G 2–3 .. 6–7 (band series; values only inside band)
H labels   (only 5 midpoint labels, e.g., '49%', blank elsewhere)

Usage: run, enter five midpoint (x,y) pairs and points-per-segment.
"""

from dataclasses import dataclass
from typing import List, Tuple
import numpy as np
import pandas as pd

# ---------- Spline helpers (Excel-like 'Smoothed line') ----------

@dataclass
class Pt:
    x: float
    y: float

def catmull_rom_to_bezier(points: List[Pt]) -> List[Tuple[Pt, Pt, Pt, Pt]]:
    """Open Catmull–Rom (tension=0) → list of cubic Bezier segments."""
    n = len(points)
    if n < 2:
        raise ValueError("Need at least two points")

    # Extend endpoints by mirroring to mimic Excel's smoothed line
    ext = [None] * (n + 2)
    ext[1:-1] = points[:]
    ext[0]  = Pt(points[0].x - (points[1].x - points[0].x),
                 points[0].y - (points[1].y - points[0].y))
    ext[-1] = Pt(points[-1].x + (points[-1].x - points[-2].x),
                 points[-1].y + (points[-1].y - points[-2].y))

    segs: List[Tuple[Pt, Pt, Pt, Pt]] = []
    for i in range(1, n):
        P0, P1, P2, P3 = ext[i-1], ext[i], ext[i+1], ext[i+2]
        C1 = Pt(P1.x + (P2.x - P0.x)/6.0, P1.y + (P2.y - P0.y)/6.0)
        C2 = Pt(P2.x - (P3.x - P1.x)/6.0, P2.y - (P3.y - P1.y)/6.0)
        segs.append((P1, C1, C2, P2))
    return segs

def sample_bezier(P1: Pt, C1: Pt, C2: Pt, P2: Pt, n_pts: int):
    """Sample one cubic Bezier at n_pts (including ends)."""
    t = np.linspace(0.0, 1.0, n_pts)
    b0 = (1 - t) ** 3
    b1 = 3 * (1 - t) ** 2 * t
    b2 = 3 * (1 - t) * t ** 2
    b3 = t ** 3
    xs = b0 * P1.x + b1 * C1.x + b2 * C2.x + b3 * P2.x
    ys = b0 * P1.y + b1 * C1.y + b2 * C2.y + b3 * P2.y
    return xs, ys

# ---------- Main ----------

def main():
    print("\nOVP curve generator (single-sheet output)")
    print("Enter FIVE midpoint (x, y) pairs, ascending by x (e.g., 2.5 0.04)")

    mids: List[Pt] = []
    for i in range(5):
        while True:
            try:
                sx, sy = input(f"Point {i+1} (x y): ").strip().split()
                mids.append(Pt(float(sx), float(sy)))
                break
            except Exception:
                print("  Please enter two numbers like: 3.5 0.49")

    mids.sort(key=lambda p: p.x)

    while True:
        try:
            n_per_seg = int(input("Points per segment (50–100 typical): ").strip())
            if n_per_seg < 4:
                print("  Use at least 4."); continue
            break
        except Exception:
            print("  Enter an integer, e.g., 60")

    # Build Excel-like smoothed curve through midpoints
    beziers = catmull_rom_to_bezier(mids)
    xs_all, ys_all = [], []
    for idx, (P1, C1, C2, P2) in enumerate(beziers):
        xs, ys = sample_bezier(P1, C1, C2, P2, n_per_seg)
        if idx > 0:
            xs, ys = xs[1:], ys[1:]  # avoid double-counting the join point
        xs_all.extend(xs)
        ys_all.extend(ys)

    df = pd.DataFrame({"x": xs_all, "y": ys_all})

    # Define band edges (edit here if you ever change bins)
    bands = [(2.0, 3.0, "2–3"),
             (3.0, 4.0, "3–4"),
             (4.0, 5.0, "4–5"),
             (5.0, 6.0, "5–6"),
             (6.0, 7.0, "6–7")]

    # Create columns C–G for bands: values only inside range, else blank (NaN)
    for i, (lo, hi, label) in enumerate(bands):
        # last band inclusive of hi
        mask = (df["x"] >= lo) & (df["x"] <= (hi if i == len(bands)-1 else hi - 1e-12))
        df[label] = np.where(mask, df["y"], np.nan)

    # Column H 'labels': only five midpoint labels (e.g., "49%")
    labels = [""] * len(df)
    for pt in mids:
        # locate nearest dense x to the midpoint x
        idx = int((df["x"] - pt.x).abs().idxmin())
        labels[idx] = f"{pt.y*100:.0f}%"
    df["labels"] = labels

    # Reorder columns: A..H
    out_df = df[["x", "y", "2–3", "3–4", "4–5", "5–6", "6–7", "labels"]]

    out_file = "Company Job Titles - ovp_curve.xlsx"
    with pd.ExcelWriter(out_file, engine="xlsxwriter") as w:
        out_df.to_excel(w, sheet_name="curve_data", index=False)

    print(f"\nSaved: {out_file}\n"
          "In Excel:\n"
          "  • Insert → Line (Smoothed) using columns A:B for the curve\n"
          "  • Add each band (C–G) as Area (regular Area, not Stacked), same X list (A)\n"
          "  • Add data labels to the curve → Value From Cells → select column H\n"
          "  • Format: areas 60–70% transparency; Y-axis percent; move vertical axis: "
          "Format X-axis → Vertical axis crosses → At maximum category\n")
    print("")
    print("WORK LOCOMOTION: Make Potential Actual")
    print("")
    return 0

if __name__ == "__main__":
    main()


