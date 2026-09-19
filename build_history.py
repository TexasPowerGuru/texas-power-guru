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

# ── TDU DELIVERY RATES, BY EFFECTIVE DATE ─────────────────────────────────────
# The tracker's REP-only cost = PowerToChoose all-in price minus TDU delivery
# (fixed $/month spread over usage + variable $/kWh). TDU rates change (usually
# March 1 and September 1), so each day is priced with the rates in effect ON that
# day. The schedule lives in tdu_rate_history.json; it is created from the seed
# below on first run. If your Excel TDURates sheet ever differs from the latest
# entry, a new entry is added automatically, effective the day it is noticed. If the
# tariff really took effect earlier (e.g. Sep 1), edit that entry's "effective"
# date in tdu_rate_history.json and run again; history is recomputed every run.
TDU_RATE_PATH = os.path.join(SCRIPT_DIR, "tdu_rate_history.json")
BRACKETS = (500, 1000, 2000)

def _r(fixed, variable):
    return {"fixed": fixed, "variable": variable}

TDU_RATE_SEED = [
    {"effective": "2000-01-01", "rates": {
        "ONCOR ELECTRIC DELIVERY COMPANY":         _r(4.06, 0.061196),
        "CENTERPOINT ENERGY HOUSTON ELECTRIC LLC": _r(4.90, 0.051461),
        "AEP TEXAS NORTH":                         _r(3.24, 0.056677),
        "AEP TEXAS CENTRAL":                       _r(3.24, 0.058272),
        "TEXAS-NEW MEXICO POWER COMPANY":          _r(7.85, 0.064665),
        "LUBBOCK POWER & LIGHT SYSTEM":            _r(0.00, 0.06312)}},
    {"effective": "2026-08-23", "rates": {
        "ONCOR ELECTRIC DELIVERY COMPANY":         _r(4.06, 0.060295),
        "CENTERPOINT ENERGY HOUSTON ELECTRIC LLC": _r(4.90, 0.049811),
        "AEP TEXAS NORTH":                         _r(3.24, 0.056677),
        "AEP TEXAS CENTRAL":                       _r(3.24, 0.058272),
        "TEXAS-NEW MEXICO POWER COMPANY":          _r(7.85, 0.064665),
        "LUBBOCK POWER & LIGHT SYSTEM":            _r(0.00, 0.06312)}},
    {"effective": "2026-09-01", "rates": {
        "ONCOR ELECTRIC DELIVERY COMPANY":         _r(4.06, 0.060295),
        "CENTERPOINT ENERGY HOUSTON ELECTRIC LLC": _r(4.90, 0.06413),
        "AEP TEXAS NORTH":                         _r(3.24, 0.056407),
        "AEP TEXAS CENTRAL":                       _r(3.24, 0.057554),
        "TEXAS-NEW MEXICO POWER COMPANY":          _r(7.85, 0.074022),
        "LUBBOCK POWER & LIGHT SYSTEM":            _r(0.00, 0.06312)}},
]

def load_rate_schedule(excel_path, today):
    """Load the dated TDU rate schedule; add an entry if Excel's rates changed."""
    if os.path.exists(TDU_RATE_PATH):
        with open(TDU_RATE_PATH, "r", encoding="utf-8") as f:
            schedule = json.load(f)
    else:
        schedule = json.loads(json.dumps(TDU_RATE_SEED))
    schedule.sort(key=lambda e: e["effective"])

    try:
        from build_from_excel import read_tdu_rates
        excel_rates = read_tdu_rates(excel_path)
    except Exception as e:
        print(f"  WARNING: could not read TDU rates from Excel ({e}).")
        excel_rates = {}

    latest = schedule[-1]["rates"]
    def differs(a, b):
        return abs(a["fixed"] - b["fixed"]) > 1e-9 or abs(a["variable"] - b["variable"]) > 1e-9
    changed = [t for t, r in excel_rates.items() if t in latest and differs(r, latest[t])]
    if changed:
        merged = dict(latest)
        merged.update({t: excel_rates[t] for t in changed})
        schedule.append({"effective": today, "rates": merged})
        print("  NOTICE: TDU delivery rates in Excel changed for: " +
              ", ".join(TDU_LABELS.get(t, t) for t in changed))
        print(f"  Recorded as effective {today}. If the tariff took effect earlier (e.g. Mar 1 / Sep 1),")
        print(f"  edit that entry's date in {TDU_RATE_PATH} and run again.")

    with open(TDU_RATE_PATH, "w", encoding="utf-8") as f:
        json.dump(schedule, f, indent=1)
    return schedule

def rates_on(schedule, date, tdu):
    """TDU rates in effect on a date (falls back to the earliest entry)."""
    chosen = schedule[0]
    for entry in schedule:
        if entry["effective"] <= date:
            chosen = entry
    return chosen["rates"].get(tdu)

def add_rep_costs(entry, tdu_rates_for_area):
    """Set rep500/rep1000/rep2000 on one TDU entry: the all-in average minus TDU
    delivery at that usage, in cents per kWh."""
    for kwh in BRACKETS:
        all_in = entry.get(f"avg{kwh}")
        if all_in is None or not tdu_rates_for_area:
            entry[f"rep{kwh}"] = None
            continue
        tdu_cents = (tdu_rates_for_area["fixed"] / kwh + tdu_rates_for_area["variable"]) * 100
        entry[f"rep{kwh}"] = round(all_in - tdu_cents, 2)

def recompute_rep_costs(history, schedule):
    """Recompute REP-only costs for EVERY snapshot using the rates in effect on
    each snapshot's own date, so a TDU rate change never shows up as a fake jump."""
    for snap in history["snapshots"]:
        for tdu, entry in snap.get("tdu", {}).items():
            add_rep_costs(entry, rates_on(schedule, snap["date"], tdu))

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

    schedule = load_rate_schedule(excel_path, today)
    for tdu, vals in averages.items():
        add_rep_costs(vals, rates_on(schedule, today, tdu))
        label = TDU_LABELS.get(tdu, tdu)
        print(f"  {label} REP-only: 500={vals['rep500']}c  1000={vals['rep1000']}c  2000={vals['rep2000']}c")

    history = load_history()
    history["snapshots"] = [s for s in history["snapshots"] if s["date"] != today]
    top_plans = get_top_plans(df, term=12, n=3)
    history["snapshots"].append({"date": today, "tdu": averages, "top12": top_plans})
    history["snapshots"].sort(key=lambda s: s["date"])
    recompute_rep_costs(history, schedule)
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
