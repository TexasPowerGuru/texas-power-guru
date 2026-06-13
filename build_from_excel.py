"""
Texas Power Guru — Excel to HTML Builder
=========================================
Reads your updated Excel file (PowertoChooseResults sheet) and rebuilds
index.html with fresh plan data.

HOW TO USE:
  Add one line at the end of your existing daily Python script:

      import subprocess
      subprocess.run(["python", r"C:/path\to\build_from_excel.py"])

  Or call build_html() directly if you prefer to import it:

      from build_from_excel import build_html
      build_html(excel_path="your_file.xlsx")

REQUIREMENTS:
  pip install pandas openpyxl
"""

import pandas as pd
import json
import re
import os
import sys
from datetime import datetime, timezone


# ── CONFIGURATION — edit these paths ─────────────────────────────────────────

# Path to your Excel file (update this to match where your script saves it)
EXCEL_PATH = "Premium_Model_App_Test.xlsx"

# Sheet name containing plan data
SHEET_NAME = "PowertoChooseResults"

# Template file (the app shell — sits next to this script)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.html")

# Output file (what GitHub Pages serves)
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), "index.html")


# ── COLUMN MAP — maps Excel columns to app field names ───────────────────────

COLUMN_MAP = {
    "TDU Area":          "tdu",
    "REP":               "rep",
    "Plan Name":         "name",
    "Length of Plan":    "term",
    "500KWH":            "rate500",
    "1000KwH":           "rate1000",
    "2000KwH":           "rate2000",
    "Cancellation Fee":  "cancelFee",
    "renewable %":       "renewable",
    "Enrollment Number": "phone",
    "Enrollment Page":   "enrollUrl",
    "EFL:":              "eflUrl",
    "BaseFee":           "baseFee",
    "EnergyCharge":      "energyCharge",
}


# ── HELPERS ───────────────────────────────────────────────────────────────────

def clean_numeric(val):
    """Strip brackets from values like '[4.79000000000000]' and return float."""
    if pd.isna(val):
        return None
    s = str(val).strip().replace("[", "").replace("]", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def clean_str(val):
    """Return a clean string or empty string for NaN/None."""
    if pd.isna(val) if not isinstance(val, str) else False:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("nan", "none") else s


# ── MAIN BUILD FUNCTION ───────────────────────────────────────────────────────

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
            print("WARNING: Could not find TDU header row — using hardcoded rates.")
            return {}
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        # Find the right column names (may vary slightly)
        tdu_col   = next((c for c in df.columns if str(c).strip().upper() == "TDU"), None)
        fixed_col = next((c for c in df.columns if "FIXED" in str(c).upper()), None)
        var_col   = next((c for c in df.columns if "VARIABLE" in str(c).upper()), None)
        if not all([tdu_col, fixed_col, var_col]):
            print("WARNING: Could not identify TDU rate columns — using hardcoded rates.")
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
        print(f"WARNING: Could not read TDURates sheet ({e}) — using hardcoded rates.")
        return {}


def build_tdu_rates_js(tdu_rates):
    """Build the JavaScript TDU_RATES object string from the dict."""
    if not tdu_rates:
        # Fall back to hardcoded values if sheet read failed
        return None
    lines = ["var TDU_RATES = {"]
    for tdu, rates in tdu_rates.items():
        safe_tdu = tdu.replace('"', '\"')
        lines.append(f'  "{safe_tdu}": {{ fixed: {rates["fixed"]:.2f},  variable: {rates["variable"]:.6f} }},')
    lines.append("};")
    return "\n".join(lines)


def build_html(excel_path=EXCEL_PATH,
               sheet_name=SHEET_NAME,
               template_path=TEMPLATE_PATH,
               output_path=OUTPUT_PATH):

    # ── 1. Read Excel ─────────────────────────────────────────────────────────
    print(f"Reading '{excel_path}' → sheet '{sheet_name}' …")
    if not os.path.exists(excel_path):
        print(f"ERROR: Excel file not found at '{excel_path}'", file=sys.stderr)
        print("       Update EXCEL_PATH in build_from_excel.py to the correct path.")
        sys.exit(1)

    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    print(f"  {len(df)} rows found")

    # ── 2. Check required columns ─────────────────────────────────────────────
    missing = [c for c in COLUMN_MAP if c not in df.columns]
    if missing:
        print(f"ERROR: Missing columns in Excel sheet: {missing}", file=sys.stderr)
        print(f"       Available columns: {df.columns.tolist()}")
        sys.exit(1)

    # ── 3. Convert rows to plan dicts ─────────────────────────────────────────
    plans = []
    skipped = 0

    for _, row in df.iterrows():
        tdu      = clean_str(row["TDU Area"])
        rep      = clean_str(row["REP"])
        name     = clean_str(row["Plan Name"])

        # Skip rows missing essential fields
        if not tdu or not rep or not name:
            skipped += 1
            continue

        try:
            term = int(float(str(row["Length of Plan"])))
        except (ValueError, TypeError):
            skipped += 1
            continue

        rate500  = clean_numeric(row["500KWH"])
        rate1000 = clean_numeric(row["1000KwH"])
        rate2000 = clean_numeric(row["2000KwH"])

        # Skip rows with no rate data at all
        if rate500 is None and rate1000 is None and rate2000 is None:
            skipped += 1
            continue

        plans.append({
            "tdu":          tdu,
            "rep":          rep,
            "name":         name,
            "term":         term,
            "rate500":      rate500,
            "rate1000":     rate1000,
            "rate2000":     rate2000,
            "cancelFee":    clean_str(row["Cancellation Fee"]),
            "renewable":    clean_str(row["renewable %"]),
            "phone":        clean_str(row["Enrollment Number"]),
            "enrollUrl":    clean_str(row["Enrollment Page"]),
            "eflUrl":       clean_str(row["EFL:"]),
            "baseFee":      clean_numeric(row["BaseFee"]),
            "energyCharge": clean_numeric(row["EnergyCharge"]),
        })

    print(f"  {len(plans)} valid plans · {skipped} rows skipped")

    if not plans:
        print("ERROR: No valid plans found. HTML not updated.", file=sys.stderr)
        sys.exit(1)

    # ── 4. Read template ──────────────────────────────────────────────────────
    if not os.path.exists(template_path):
        print(f"ERROR: template.html not found at '{template_path}'", file=sys.stderr)
        sys.exit(1)

    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    if "PLANS_DATA_PLACEHOLDER" not in template:
        print("ERROR: template.html is missing the PLANS_DATA_PLACEHOLDER token.", file=sys.stderr)
        sys.exit(1)

    # ── 5. Inject data + timestamp ────────────────────────────────────────────
    plans_json = json.dumps(plans, separators=(",", ":"))
    updated_at = datetime.now(timezone.utc).strftime("%B %d, %Y at %H:%M UTC")

    html = template.replace("PLANS_DATA_PLACEHOLDER", plans_json)
    html = html.replace(
        "Rates sourced from PowerToChoose",
        f"Last updated: {updated_at} · Rates sourced from PowerToChoose"
    )

    # ── 6. Write output ───────────────────────────────────────────────────────
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = len(html) / 1024
    print(f"  index.html written ({size_kb:.0f} KB) → '{output_path}'")
    print(f"  Timestamp: {updated_at}")
    print("Done!")


# ── RUN DIRECTLY ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Optional: pass excel path as command-line argument
    # e.g.  python build_from_excel.py "C:\data\my_plans.xlsx"
    if len(sys.argv) > 1:
        build_html(excel_path=sys.argv[1])
    else:
        build_html()
