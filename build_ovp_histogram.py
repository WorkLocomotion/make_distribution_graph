#!/usr/bin/env python3
from pathlib import Path
import sys
import numpy as np
import pandas as pd

def main():
    print("\nOVP histogram (bins + midpoints + % of employees)")
    print("Assumes columns named exactly 'Average OVP' and 'Headcount'")
    print("-----------------------------------------------------------")

    # 1) Ask for file path
    while True:
        p = input('Enter or drag/drop the Excel file path (.xlsx): ').strip().strip('"')
        if not p:
            print("  ⚠️ Please enter a path."); continue
        path = Path(p)
        if path.exists() and path.suffix.lower() == ".xlsx":
            break
        print("  ❌ Not found or not an .xlsx file. Try again.")

    # 2) Read first sheet
    try:
        df = pd.read_excel(path)
    except Exception as e:
        print(f"  ❌ Failed to read Excel: {e}")
        sys.exit(1)

    # 3) Validate required columns
    required = {"Average OVP", "Headcount"}
    missing = required - set(df.columns)
    if missing:
        print(f"  ❌ Missing required column(s): {', '.join(sorted(missing))}")
        print(f"  Columns found: {list(df.columns)}")
        sys.exit(1)

    # 4) Clean/coerce
    work = df[["Average OVP", "Headcount"]].copy()
    work["Average OVP"] = pd.to_numeric(work["Average OVP"], errors="coerce")
    work["Headcount"]   = pd.to_numeric(work["Headcount"], errors="coerce")
    work = work.dropna(subset=["Average OVP", "Headcount"])
    work = work[work["Headcount"] > 0]

    if work.empty:
        print("  ❌ No valid rows after cleaning (check data).")
        sys.exit(1)

    # 5) Define bins explicitly and compute with boolean masks
    edges = [2.0, 3.0, 4.0, 5.0, 6.0, 7.0]
    bins = [(edges[i], edges[i+1]) for i in range(len(edges)-1)]

    results = []
    total_hc = 0
    # First pass to get total_hc in-range (so % excludes out-of-range)
    for i, (lo, hi) in enumerate(bins):
        if i < len(bins)-1:
            mask = (work["Average OVP"] >= lo) & (work["Average OVP"] < hi)
        else:
            mask = (work["Average OVP"] >= lo) & (work["Average OVP"] <= hi)
        hc_sum = float(work.loc[mask, "Headcount"].sum())
        total_hc += hc_sum
        results.append((lo, hi, hc_sum))

    # Build output rows with shares and midpoints
    rows = []
    for (lo, hi, hc_sum) in results:
        share = (hc_sum / total_hc) if total_hc > 0 else 0.0
        midpoint = (lo + hi) / 2.0
        label = f"{int(lo)}–{int(hi)}"
        rows.append({
            "Bin Lower": lo,
            "Bin Upper": hi,
            "Midpoint": midpoint,
            "Range Label": label,
            "Headcount": hc_sum,
            "Share of Total (%)": share
        })

    out_df = pd.DataFrame(rows)

    # 6) Write output
    out_file = "Company Job Titles - ovp_histogram.xlsx"
    with pd.ExcelWriter(out_file, engine="xlsxwriter") as w:
        out_df.to_excel(w, sheet_name="ovp_histogram", index=False)
        wb = w.book
        ws = w.sheets["ovp_histogram"]
        pct = wb.add_format({"num_format": "0%"})
        ws.set_column("A:C", 12)   # bounds & midpoint
        ws.set_column("D:D", 14)   # label
        ws.set_column("E:E", 14)   # headcount
        ws.set_column("F:F", 16, pct)  # share %

    print(f"\n✅ Saved: {out_file}")
    print(f"   In-range total headcount: {int(total_hc)}\n")
    print("")
    print("WORK LOCOMOTION: Make Potential Actual")
    print("")
    return 0

if __name__ == "__main__":
    main()
