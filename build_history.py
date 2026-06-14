"""
Texas Power Guru — Rate History Tracker
========================================
Run this daily alongside build_from_excel.py.
Reads your Excel, calculates average rates per TDU, appends a snapshot
to rate_history.json, then rebuilds rate_tracker.html.

HOW TO USE:
  Add to the end of your existing daily Python script:

      from build_history import update_history
      update_history(excel_path=r"C:\\path\\to\\your_file.xlsx")
"""

import pandas as pd
import json
import os
import sys
from datetime import datetime, timezone

EXCEL_PATH           = "Premium_Model_App_Test.xlsx"
SHEET_NAME           = "PowertoChooseResults"
SCRIPT_DIR           = os.path.dirname(os.path.abspath(__file__))
HISTORY_PATH         = os.path.join(SCRIPT_DIR, "rate_history.json")
TRACKER_TEMPLATE     = os.path.join(SCRIPT_DIR, "tracker_template.html")
TRACKER_OUTPUT       = os.path.join(SCRIPT_DIR, "rate_tracker.html")
MAX_DAYS             = 365 * 3

TDU_LABELS = {
    "AEP TEXAS CENTRAL":                       "AEP Texas Central",
    "AEP TEXAS NORTH":                         "AEP Texas North",
    "CENTERPOINT ENERGY HOUSTON ELECTRIC LLC": "CenterPoint Houston",
    "ONCOR ELECTRIC DELIVERY COMPANY":         "Oncor",
    "TEXAS-NEW MEXICO POWER COMPANY":          "Texas-NM Power",
    "LUBBOCK POWER & LIGHT SYSTEM":            "Lubbock P&L",
}

def clean_numeric(val):
    if pd.isna(val):
        return None
    s = str(val).strip().replace("[", "").replace("]", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return None

def calc_tdu_averages(df):
    results = {}
    for tdu in df["TDU Area"].unique():
        tdu = str(tdu).strip()
        if not tdu or tdu.lower() == "nan":
            continue
        subset = df[df["TDU Area"] == tdu]
        r500  = [clean_numeric(v) for v in subset["500KWH"]  if clean_numeric(v) is not None]
        r1000 = [clean_numeric(v) for v in subset["1000KwH"] if clean_numeric(v) is not None]
        r2000 = [clean_numeric(v) for v in subset["2000KwH"] if clean_numeric(v) is not None]
        results[tdu] = {
            "avg500":  round(sum(r500) /len(r500),  2) if r500  else None,
            "avg1000": round(sum(r1000)/len(r1000), 2) if r1000 else None,
            "avg2000": round(sum(r2000)/len(r2000), 2) if r2000 else None,
            "count":   len(subset),
        }
    return results

def get_top_plans(df, term=12, n=3):
    """Get top N most affordable plans at 1000 kWh for each TDU for a given term."""
    def clean_str(val):
        if not isinstance(val, str) and pd.isna(val):
            return ""
        s = str(val).strip()
        return "" if s.lower() in ("nan", "none") else s

    results = {}
    for tdu in df["TDU Area"].unique():
        tdu = str(tdu).strip()
        if not tdu or tdu.lower() == "nan":
            continue
        subset = df[
            (df["TDU Area"] == tdu) &
            (df["Length of Plan"] == term)
        ].copy()
        if subset.empty:
            continue
        subset["_r1000"] = subset["1000KwH"].apply(clean_numeric)
        subset = subset.dropna(subset=["_r1000"])
        subset = subset.sort_values("_r1000").head(n)
        top = []
        for _, row in subset.iterrows():
            top.append({
                "rep":      clean_str(row["REP"]),
                "name":     clean_str(row["Plan Name"]),
                "rate500":  clean_numeric(row["500KWH"]),
                "rate1000": clean_numeric(row["1000KwH"]),
                "rate2000": clean_numeric(row["2000KwH"]),
                "cancel":   clean_str(row["Cancellation Fee"]),
                "renewable":clean_str(row["renewable %"]),
                "enrollUrl":clean_str(row["Enrollment Page"]),
            })
        results[tdu] = top
    return results


def load_history():
    if os.path.exists(HISTORY_PATH):
        with open(HISTORY_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"snapshots": []}

def save_history(history):
    history["snapshots"] = history["snapshots"][-MAX_DAYS:]
    with open(HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(history, f, separators=(",", ":"))

def build_tracker():
    if not os.path.exists(TRACKER_TEMPLATE):
        print("  WARNING: tracker_template.html not found.")
        return
    if not os.path.exists(HISTORY_PATH):
        print("  WARNING: rate_history.json not found.")
        return
    with open(TRACKER_TEMPLATE, "r", encoding="utf-8") as f:
        template = f.read()
    with open(HISTORY_PATH, "r", encoding="utf-8") as f:
        history_json = f.read()
    html = template.replace("HISTORY_DATA_PLACEHOLDER", history_json)
    with open(TRACKER_OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  rate_tracker.html rebuilt ({len(html)//1024} KB)")

def update_history(excel_path=EXCEL_PATH, sheet_name=SHEET_NAME):
    print("=" * 60)
    print("Texas Power Guru — Rate History Tracker")
    print("=" * 60)

    if not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found at {excel_path}", file=sys.stderr)
        sys.exit(1)

    print(f"\nReading {excel_path} ...")
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    print(f"  {len(df)} rows found")

    today    = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    averages = calc_tdu_averages(df)

    for tdu, vals in averages.items():
        label = TDU_LABELS.get(tdu, tdu)
        print(f"  {label}: 500={vals['avg500']}c  1000={vals['avg1000']}c  2000={vals['avg2000']}c  ({vals['count']} plans)")

    history = load_history()
    history["snapshots"] = [s for s in history["snapshots"] if s["date"] != today]
    top_plans = get_top_plans(df, term=12, n=3)
    history["snapshots"].append({"date": today, "tdu": averages, "top12": top_plans})
    history["snapshots"].sort(key=lambda s: s["date"])
    save_history(history)

    total = len(history["snapshots"])
    print(f"\n  History updated: {total} day(s) on record")
    print(f"  Saved to: {HISTORY_PATH}")

    print("\nRebuilding rate_tracker.html ...")
    build_tracker()
    print("\nDone!")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        update_history(excel_path=sys.argv[1])
    else:
        update_history()
