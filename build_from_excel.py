"""
Texas Power Guru — Excel to HTML Builder
=========================================
Reads your updated Excel file and builds TWO versions of the app:

  index.html            — full version (plan name, enroll link, EFL shown)
  index_restricted.html — restricted version (plan name, enroll link, EFL hidden)

HOW TO USE:
  Add this to the end of your existing daily Python script:

      from build_from_excel import build_all
      build_all(excel_path=r"C:\path\to\your_file.xlsx")

  Or run directly in Spyder: open this file and press F5.

REQUIREMENTS:
  pip install pandas openpyxl
"""

import pandas as pd
import json
import re
import os
import sys
from datetime import datetime, timezone


# ── CONFIGURATION — edit these paths ──────────────────────────────────────────

EXCEL_PATH   = r"C:\Users\ZazuE\OneDrive - Zazu Energy\SpreadsheetWeb\Published\DualCalculators\AppHosting\PowertoChooseResults.xlsx"
SHEET_NAME   = "PowertoChooseResults"

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_FULL        = os.path.join(SCRIPT_DIR, "template.html")
TEMPLATE_RESTRICTED  = os.path.join(SCRIPT_DIR, "template_restricted.html")
OUTPUT_FULL          = os.path.join(SCRIPT_DIR, "index.html")
OUTPUT_RESTRICTED    = os.path.join(SCRIPT_DIR, "index_restricted.html")


# ── HELPERS ───────────────────────────────────────────────────────────────────

def clean_numeric(val):
    """Strip brackets from '[4.79]' style values and return float."""
    if pd.isna(val):
        return None
    s = str(val).strip().replace("[", "").replace("]", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def clean_str(val):
    """Return clean string or empty string for NaN/None."""
    if not isinstance(val, str) and pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none") else s


# ── READ TDU RATES FROM EXCEL ─────────────────────────────────────────────────

def read_tdu_rates(excel_path):
    """Read TDU fixed and variable delivery charges from the TDURates sheet."""
    try:
        df = pd.read_excel(excel_path, sheet_name="TDURates", header=None)
        # Find the header row containing "TDU"
        header_row = None
        for i, row in df.iterrows():
            if any(str(v).strip().upper() == "TDU" for v in row):
                header_row = i
                break
        if header_row is None:
            print("  WARNING: Could not find TDU header row — using hardcoded rates.")
            return {}
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        tdu_col   = next((c for c in df.columns if str(c).strip().upper() == "TDU"), None)
        fixed_col = next((c for c in df.columns if "FIXED" in str(c).upper()), None)
        var_col   = next((c for c in df.columns if "VARIABLE" in str(c).upper()), None)
        if not all([tdu_col, fixed_col, var_col]):
            print("  WARNING: Could not identify TDU rate columns — using hardcoded rates.")
            return {}
        rates = {}
        for _, row in df.iterrows():
            tdu = str(row[tdu_col]).strip()
            if not tdu or tdu.lower() == "nan":
                continue
            try:
                rates[tdu] = {
                    "fixed":    float(str(row[fixed_col]).replace(",", "") or 0),
                    "variable": float(str(row[var_col]).replace(",", "")  or 0),
                }
            except (ValueError, TypeError):
                continue
        return rates
    except Exception as e:
        print(f"  WARNING: Could not read TDURates sheet ({e}) — using hardcoded rates.")
        return {}


def build_tdu_rates_js(tdu_rates):
    """Build the JavaScript TDU_RATES object string from the dict."""
    if not tdu_rates:
        return None
    lines = ["var TDU_RATES = {"]
    for tdu, rates in tdu_rates.items():
        safe = tdu.replace('"', '\\"')
        lines.append(
            '  "' + safe + '": { fixed: ' +
            str(round(rates["fixed"], 2)) + ',  variable: ' +
            str(round(rates["variable"], 6)) + ' },'
        )
    lines.append("};")
    return "\n".join(lines)


# ── READ PLAN DATA FROM EXCEL ─────────────────────────────────────────────────

def read_plans(excel_path, sheet_name):
    """Read and clean plan rows from the PowertoChooseResults sheet."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    print(f"  {len(df)} rows found")

    required = ["TDU Area", "REP", "Plan Name", "Length of Plan",
                "500KWH", "1000KwH", "2000KwH", "Cancellation Fee",
                "renewable %", "Enrollment Number", "Enrollment Page",
                "EFL:", "BaseFee", "EnergyCharge"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"ERROR: Missing columns: {missing}", file=sys.stderr)
        sys.exit(1)

    plans, skipped = [], 0
    for _, row in df.iterrows():
        tdu  = clean_str(row["TDU Area"])
        rep  = clean_str(row["REP"])
        name = clean_str(row["Plan Name"])
        if not tdu or not rep or not name:
            skipped += 1
            continue
        try:
            term = int(float(str(row["Length of Plan"])))
        except (ValueError, TypeError):
            skipped += 1
            continue
        r500  = clean_numeric(row["500KWH"])
        r1000 = clean_numeric(row["1000KwH"])
        r2000 = clean_numeric(row["2000KwH"])
        if r500 is None and r1000 is None and r2000 is None:
            skipped += 1
            continue
        plans.append({
            "tdu":          tdu,
            "rep":          rep,
            "name":         name,
            "term":         term,
            "rate500":      r500,
            "rate1000":     r1000,
            "rate2000":     r2000,
            "cancelFee":    clean_str(row["Cancellation Fee"]),
            "renewable":    clean_str(row["renewable %"]),
            "phone":        clean_str(row["Enrollment Number"]),
            "enrollUrl":    clean_str(row["Enrollment Page"]),
            "eflUrl":       clean_str(row["EFL:"]),
            "baseFee":      clean_numeric(row["BaseFee"]),
            "energyCharge": clean_numeric(row["EnergyCharge"]),
        })
    print(f"  {len(plans)} valid plans · {skipped} rows skipped")
    return plans


# ── BUILD ONE HTML FILE ───────────────────────────────────────────────────────

def build_one(plans, tdu_rates_js, updated_at, template_path, output_path, label):
    """Inject plan data into a template and write the output file."""
    if not os.path.exists(template_path):
        print(f"  ERROR: template not found at '{template_path}'", file=sys.stderr)
        return False
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()
    if "PLANS_DATA_PLACEHOLDER" not in template:
        print(f"  ERROR: PLANS_DATA_PLACEHOLDER missing in {template_path}", file=sys.stderr)
        return False

    html = template.replace("PLANS_DATA_PLACEHOLDER", json.dumps(plans, separators=(",", ":")))

    # Inject fresh TDU rates if we got them from Excel
    if tdu_rates_js:
        html = re.sub(r"var TDU_RATES = \{.*?\};", tdu_rates_js, html, flags=re.DOTALL)

    html = html.replace(
        "Rates sourced from PowerToChoose",
        f"Last updated: {updated_at} · Rates sourced from PowerToChoose"
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  {label} written ({len(html)//1024} KB) → '{output_path}'")
    return True


# ── MAIN ──────────────────────────────────────────────────────────────────────

def build_all(excel_path=EXCEL_PATH, sheet_name=SHEET_NAME):
    print("=" * 60)
    print("Texas Power Guru — Excel to HTML Builder")
    print("=" * 60)

    if not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found at '{excel_path}'", file=sys.stderr)
        print("       Update EXCEL_PATH at the top of build_from_excel.py")
        sys.exit(1)

    # 1. Read plans
    print(f"\nReading plans from '{excel_path}' …")
    plans = read_plans(excel_path, sheet_name)
    if not plans:
        print("ERROR: No valid plans found.", file=sys.stderr)
        sys.exit(1)

    # 2. Read TDU rates
    print("\nReading TDU delivery rates …")
    tdu_rates = read_tdu_rates(excel_path)
    tdu_rates_js = build_tdu_rates_js(tdu_rates)
    for tdu, r in tdu_rates.items():
        print(f"  {tdu}: fixed=${r['fixed']:.2f}  variable=${r['variable']:.6f}")

    updated_at = datetime.now(timezone.utc).strftime("%B %d, %Y at %H:%M UTC")

    # 3. Build full version
    print("\nBuilding full version (index.html) …")
    build_one(plans, tdu_rates_js, updated_at, TEMPLATE_FULL, OUTPUT_FULL, "index.html")

    # 4. Build restricted version
    print("\nBuilding restricted version (index_restricted.html) …")
    build_one(plans, tdu_rates_js, updated_at, TEMPLATE_RESTRICTED, OUTPUT_RESTRICTED, "index_restricted.html")

    print(f"\nAll done! Timestamp: {updated_at}")


# ── RUN DIRECTLY ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        build_all(excel_path=sys.argv[1])
    else:
        build_all()
