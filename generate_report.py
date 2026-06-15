"""
Texas Power Guru — Client Energy Report Generator
===================================================
Generates a static HTML report showing the most affordable plan per
unique contract term in three categories:
  - Short-Term  (< 12 months)
  - Mid-Term    (12–24 months)
  - Long-Term   (36+ months)

HOW TO USE in Spyder:
    from generate_report import generate_report

    generate_report(
        client_name   = "John Smith",
        tdu           = "ONCOR ELECTRIC DELIVERY COMPANY",
        monthly_usage = [900,750,700,650,800,1200,1800,2000,1600,1000,750,800],
        output_path   = r"C:\\Users\\ZazuE\\Desktop\\john_smith_report.html"
    )

    The HTML file can be:
      - Opened directly in any browser
      - Printed to PDF via Ctrl+P → Save as PDF
      - Emailed as an attachment
      - Shared via a link if hosted online

PARAMETERS:
    client_name    : Client name shown in the report header
    tdu            : One of the 6 Texas TDU service area keys
    monthly_usage  : List of 12 kWh values (Jan–Dec). Use 0 or None to estimate.
    output_path    : Where to save the HTML file
    renewable_filter: "any" (default) or "green" (100% renewable only)
"""

import json, re, os, sys
from datetime import datetime

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML  = os.path.join(SCRIPT_DIR, "index.html")
DEFAULT_OUT = os.path.join(os.path.expanduser("~"), "Desktop", "texas_power_guru_report.html")

TDU_LABELS = {
    "AEP TEXAS CENTRAL":                       "AEP Texas Central",
    "AEP TEXAS NORTH":                         "AEP Texas North",
    "CENTERPOINT ENERGY HOUSTON ELECTRIC LLC": "CenterPoint Energy Houston",
    "ONCOR ELECTRIC DELIVERY COMPANY":         "Oncor Electric Delivery",
    "TEXAS-NEW MEXICO POWER COMPANY":          "Texas-New Mexico Power (TNMP)",
    "LUBBOCK POWER & LIGHT SYSTEM":            "Lubbock Power & Light",
}

MONTH_ABBR  = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
MONTH_NAMES = ["January","February","March","April","May","June",
               "July","August","September","October","November","December"]
SEASON_RATIO = [0.963890,0.796843,0.732663,0.710121,0.819878,1.167162,
                1.411973,1.471527,1.406559,1.003560,0.755894,0.759930]
HIST_AVG     = [890.18,735.91,676.64,655.82,757.18,1077.91,
                1304,1359,1299,926.82,698.09,701.82]


# ── SEASONAL PROJECTION ───────────────────────────────────────────────────────
def project_usage(monthly_usage):
    inputs = [float(v) if v and float(v) > 0 else None for v in monthly_usage]
    bases  = [inputs[i]/SEASON_RATIO[i] for i in range(12) if inputs[i] is not None]
    base   = sum(bases)/len(bases) if bases else None
    return [
        inputs[i] if inputs[i] is not None
        else (round(base * SEASON_RATIO[i]) if base else round(HIST_AVG[i]))
        for i in range(12)
    ]


# ── PLAN LOADING ──────────────────────────────────────────────────────────────
def load_plans():
    if not os.path.exists(INDEX_HTML):
        print(f"ERROR: index.html not found at {INDEX_HTML}"); sys.exit(1)
    with open(INDEX_HTML, "r", encoding="utf-8") as f:
        html = f.read()
    match = re.search(r'const PLANS = (\[.*?\]);', html, re.DOTALL)
    if not match:
        print("ERROR: Could not find plan data in index.html"); sys.exit(1)
    return json.loads(match.group(1))


# ── COST CALCULATION ──────────────────────────────────────────────────────────
def interpolate_rate(plan, kwh):
    r500, r1000, r2000 = plan.get("rate500"), plan.get("rate1000"), plan.get("rate2000")
    if kwh <= 500:  return r500 or r1000 or r2000
    if kwh <= 1000:
        if r500 is None or r1000 is None: return r1000 or r500
        return r500 + (kwh-500)/500 * (r1000-r500)
    if kwh <= 2000:
        if r1000 is None or r2000 is None: return r1000 or r2000
        return r1000 + (kwh-1000)/1000 * (r2000-r1000)
    return r2000 or r1000

def calc_costs(plan, projected):
    monthly = [(interpolate_rate(plan, kwh)/100)*kwh
               if interpolate_rate(plan, kwh) else None
               for kwh in projected]
    annual = sum(monthly) if all(c is not None for c in monthly) else None
    return annual, monthly

def parse_cancel(fee_str):
    s = str(fee_str).replace("Cancellation Fee:","").replace("$$","$").strip()
    if re.search(r"per\s*month|month\s*remain|remaining|/\s*month", s, re.I):
        m = re.search(r"\$(\d+)", s)
        return f"${m.group(1)}/mo remaining" if m else "Variable"
    m = re.search(r"\$([\d.]+)", s)
    if m:
        amt = float(m.group(1))
        return "None" if amt == 0 else f"${amt:.0f}"
    return s or "None"

def best_plan_per_term(plans, projected):
    by_term = {}
    for p in plans:
        annual, monthly = calc_costs(p, projected)
        if annual is None: continue
        p = dict(p, _annual=annual, _monthly=monthly)
        if p["term"] not in by_term or annual < by_term[p["term"]]["_annual"]:
            by_term[p["term"]] = p
    return sorted(by_term.values(), key=lambda x: x["_annual"])

def esc(s):
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")


# ── HTML BUILDER ──────────────────────────────────────────────────────────────
def build_plans_section(plans, monthly_usage, projected, annual_kwh, accent, label):
    if not plans:
        return f'<p style="color:#718096;font-size:13px;padding:12px 0">No plans available in this category for your service area.</p>'

    rows_html = ""
    medals = ["🥇","🥈","🥉","#4","#5","#6","#7"]

    for rank, p in enumerate(plans):
        annual  = p["_annual"]
        monthly = p["_monthly"]
        avg_rate = annual / annual_kwh * 100

        renew_pct = int(re.search(r"(\d+)%", p.get("renewable","0")).group(1)) \
                    if re.search(r"(\d+)%", p.get("renewable","")) else 0
        cancel_str = parse_cancel(p.get("cancelFee",""))
        enroll_url = p.get("enrollUrl","")
        efl_url    = p.get("eflUrl","")
        is_green   = renew_pct == 100

        # Monthly cost cells
        mo_cells = ""
        for i, cost in enumerate(monthly):
            is_entered = monthly_usage[i] and float(monthly_usage[i]) > 0
            style = "color:#1A202C" if is_entered else "color:#1A73C8;font-style:italic"
            mo_cells += f'<td style="text-align:center;padding:6px 4px;{style}">${cost:,.0f}</td>'

        enroll_btn = (f'<a href="{esc(enroll_url)}" target="_blank" rel="noopener" '
                      f'style="display:inline-block;padding:5px 14px;background:#0E4F8B;color:white;'
                      f'border-radius:6px;font-size:11px;font-weight:700;text-decoration:none;'
                      f'margin-right:6px">⚡ Enroll</a>') if enroll_url and enroll_url not in ("nan","None","") else ""

        efl_btn = (f'<a href="{esc(efl_url)}" target="_blank" rel="noopener" '
                   f'style="display:inline-block;padding:5px 14px;background:transparent;color:#0E4F8B;'
                   f'border:1.5px solid #0E4F8B;border-radius:6px;font-size:11px;font-weight:600;'
                   f'text-decoration:none">📄 EFL</a>') if efl_url and efl_url not in ("nan","None","") else ""

        renew_badge = (f'<span style="background:#F0FFF4;color:#276749;border:1px solid #C6F6D5;'
                       f'padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">'
                       f'🌱 {renew_pct}%</span>') if is_green else \
                      (f'<span style="background:#F7FAFC;color:#718096;border:1px solid #E2E8F0;'
                       f'padding:2px 8px;border-radius:10px;font-size:11px">{renew_pct}% Renewable</span>')

        row_bg = "#FAFBFC" if rank % 2 == 1 else "white"
        medal_str = medals[rank] if rank < len(medals) else f"#{rank+1}"

        rows_html += f"""
        <tr style="background:{row_bg}">
          <td style="padding:10px 8px;font-weight:700;color:#0E4F8B;text-align:center;
                     font-size:14px;white-space:nowrap">{medal_str}</td>
          <td style="padding:10px 8px">
            <div style="font-size:11px;font-weight:700;color:#718096;text-transform:uppercase;
                        letter-spacing:.4px;margin-bottom:2px">{esc(p['rep'])}</div>
            <div style="font-size:13px;font-weight:700;color:#1A202C">{esc(p['name'])}</div>
            <div style="margin-top:6px">{enroll_btn}{efl_btn}</div>
          </td>
          <td style="padding:10px 8px;text-align:center;font-weight:600">{p['term']} mo</td>
          <td style="padding:10px 8px;text-align:center">
            <div style="font-size:18px;font-weight:800;color:#0E4F8B">${annual:,.0f}</div>
            <div style="font-size:10px;color:#718096">per year</div>
          </td>
          <td style="padding:10px 8px;text-align:center">
            <div style="font-size:14px;font-weight:700;color:#1A202C">{avg_rate:.1f}¢</div>
            <div style="font-size:10px;color:#718096">avg/kWh</div>
          </td>
          <td style="padding:10px 8px;text-align:center;font-size:12px">{cancel_str}</td>
          <td style="padding:10px 8px;text-align:center">{renew_badge}</td>
          {mo_cells}
          <td style="padding:10px 8px;text-align:center;font-weight:700;color:#0E4F8B;
                     background:#EBF4FF;font-size:13px">${annual:,.0f}</td>
        </tr>"""

    mo_headers = "".join(f'<th style="padding:8px 4px;font-size:11px;font-weight:600;'
                         f'text-align:center">{m}</th>' for m in MONTH_ABBR)

    return f"""
    <div style="overflow-x:auto;margin-bottom:4px">
      <table style="width:100%;border-collapse:collapse;font-size:12px">
        <thead>
          <tr style="background:#0A3560;color:white">
            <th style="padding:8px 6px;text-align:center;width:36px"></th>
            <th style="padding:8px 8px;text-align:left;min-width:200px">Provider / Plan</th>
            <th style="padding:8px 6px;text-align:center">Term</th>
            <th style="padding:8px 6px;text-align:center">Annual Cost</th>
            <th style="padding:8px 6px;text-align:center">Avg Rate</th>
            <th style="padding:8px 6px;text-align:center">ETF</th>
            <th style="padding:8px 6px;text-align:center">Renewable</th>
            {mo_headers}
            <th style="padding:8px 6px;text-align:center;background:#1A73C8">Annual</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
    <p style="font-size:11px;color:#718096;margin:4px 0 0">
      <em>Blue italic monthly costs = estimated from seasonal averages based on your entered months.</em>
    </p>"""


def generate_report(client_name, tdu, monthly_usage,
                    output_path=DEFAULT_OUT, renewable_filter="any"):

    print(f"Generating report for {client_name}...")

    projected  = project_usage(monthly_usage)
    annual_kwh = sum(projected)
    avg_kwh    = annual_kwh / 12
    entered    = sum(1 for v in monthly_usage if v and float(v) > 0)

    all_plans = load_plans()
    filtered  = [p for p in all_plans if p["tdu"] == tdu]
    if renewable_filter == "green":
        filtered = [p for p in filtered if "100%" in p.get("renewable","")]

    cat_short = best_plan_per_term([p for p in filtered if p["term"] < 12],  projected)
    cat_mid   = best_plan_per_term([p for p in filtered if 12 <= p["term"] <= 24], projected)
    cat_long  = best_plan_per_term([p for p in filtered if p["term"] > 36],  projected)

    now = datetime.now().strftime("%B %d, %Y")

    # Monthly usage summary table
    mo_usage_rows = ""
    for i, kwh in enumerate(projected):
        is_entered = monthly_usage[i] and float(monthly_usage[i]) > 0
        style = "color:#1A202C;font-weight:600" if is_entered else "color:#1A73C8;font-style:italic"
        status = "Entered" if is_entered else "Estimated"
        status_style = "color:#276749;font-weight:600" if is_entered else "color:#1A73C8"
        mo_usage_rows += f"""
        <td style="padding:8px 6px;text-align:center;border-right:1px solid #E2E8F0">
          <div style="{style}">{round(kwh):,}</div>
          <div style="font-size:10px;{status_style}">{status}</div>
        </td>"""

    cats = [
        ("Short-Term Plans",  "Less than 12 Months",  cat_short, "#CD7F32"),
        ("Mid-Term Plans",    "12 to 24 Months",       cat_mid,   "#A8A9AD"),
        ("Long-Term Plans",   "36+ Months",            cat_long,  "#D4AF37"),
    ]

    sections_html = ""
    for title, subtitle, plans, accent in cats:
        plan_html = build_plans_section(plans, monthly_usage, projected,
                                        annual_kwh, accent, title)
        sections_html += f"""
      <div style="margin-bottom:28px">
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:10px">
          <div style="width:5px;height:36px;background:{accent};border-radius:3px;flex-shrink:0"></div>
          <div>
            <div style="font-family:'Sora',sans-serif;font-size:16px;font-weight:700;
                        color:#0A3560">{title}</div>
            <div style="font-size:12px;color:#718096">{subtitle} — Best plan per unique term length</div>
          </div>
        </div>
        {plan_html}
      </div>"""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Texas Power Guru — Report for {esc(client_name)}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Sora:wght@700;800&display=swap" rel="stylesheet">
<style>
  *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:'Inter',sans-serif;background:#F0F4F8;color:#1A202C;font-size:14px;line-height:1.5}}
  @media print{{
    body{{background:white}}
    .no-print{{display:none}}
    .page-break{{page-break-before:always}}
  }}
</style>
</head>
<body>

<div class="no-print" style="background:#0A3560;color:white;padding:10px 24px;
     display:flex;justify-content:space-between;align-items:center">
  <span style="font-size:13px;opacity:.85">Texas Power Guru — Client Report</span>
  <button onclick="window.print()" style="padding:6px 16px;background:#1A73C8;color:white;
    border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">
    🖨 Print / Save as PDF
  </button>
</div>

<div style="max-width:1100px;margin:0 auto;padding:24px 20px 48px">

  <!-- HEADER -->
  <div style="background:linear-gradient(135deg,#0A3560,#1A73C8);border-radius:12px;
       padding:24px 28px 28px;margin-bottom:20px;position:relative;overflow:hidden">
    <div style="position:absolute;top:-40px;right:-40px;width:160px;height:160px;
         background:rgba(255,255,255,.06);border-radius:50%"></div>
    <div style="font-family:'Sora',sans-serif;font-size:24px;font-weight:800;
         color:white;margin-bottom:4px">⚡ Texas Power Guru</div>
    <div style="font-size:14px;color:rgba(255,255,255,.8)">
      Personalized Electricity Plan Comparison Report
    </div>
  </div>

  <!-- CLIENT INFO -->
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px">
    <div style="background:white;border-radius:10px;padding:14px 18px;
         box-shadow:0 1px 3px rgba(0,0,0,.08)">
      <div style="font-size:11px;font-weight:700;color:#718096;text-transform:uppercase;
           letter-spacing:.5px;margin-bottom:4px">Client</div>
      <div style="font-size:16px;font-weight:700;color:#0A3560">{esc(client_name)}</div>
      <div style="font-size:12px;color:#718096;margin-top:2px">{now}</div>
    </div>
    <div style="background:white;border-radius:10px;padding:14px 18px;
         box-shadow:0 1px 3px rgba(0,0,0,.08)">
      <div style="font-size:11px;font-weight:700;color:#718096;text-transform:uppercase;
           letter-spacing:.5px;margin-bottom:4px">Service Area</div>
      <div style="font-size:14px;font-weight:700;color:#1A202C">{esc(TDU_LABELS.get(tdu,tdu))}</div>
    </div>
    <div style="background:white;border-radius:10px;padding:14px 18px;
         box-shadow:0 1px 3px rgba(0,0,0,.08)">
      <div style="font-size:11px;font-weight:700;color:#718096;text-transform:uppercase;
           letter-spacing:.5px;margin-bottom:4px">Projected Usage</div>
      <div style="font-size:16px;font-weight:800;color:#0A3560;font-family:'Sora',sans-serif">
        {round(annual_kwh):,} <span style="font-size:12px;font-weight:600">kWh/year</span>
      </div>
      <div style="font-size:12px;color:#718096">{round(avg_kwh):,} kWh/mo avg
        &nbsp;·&nbsp; {entered}/12 months entered</div>
    </div>
  </div>

  <!-- MONTHLY USAGE -->
  <div style="background:white;border-radius:10px;padding:16px 18px;margin-bottom:20px;
       box-shadow:0 1px 3px rgba(0,0,0,.08)">
    <div style="font-family:'Sora',sans-serif;font-size:13px;font-weight:700;color:#0A3560;
         margin-bottom:10px">Projected Monthly Usage (kWh)</div>
    <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse;min-width:600px">
        <thead>
          <tr style="background:#0A3560;color:white">
            {"".join(f'<th style="padding:7px 6px;font-size:11px;text-align:center;font-weight:600">{m}</th>' for m in MONTH_ABBR)}
            <th style="padding:7px 6px;font-size:11px;text-align:center;font-weight:600;background:#1A73C8">Annual</th>
          </tr>
        </thead>
        <tbody>
          <tr style="background:#F7FAFC">
            {mo_usage_rows}
            <td style="padding:8px 6px;text-align:center;font-weight:800;color:#0A3560;
                background:#EBF4FF;font-size:14px">{round(annual_kwh):,}</td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>

  <!-- PLAN SECTIONS -->
  <div style="background:white;border-radius:10px;padding:20px 18px;
       box-shadow:0 1px 3px rgba(0,0,0,.08)">
    {sections_html}

    <!-- DISCLAIMER -->
    <div style="border-top:1px solid #E2E8F0;margin-top:16px;padding-top:12px;
         font-size:11px;color:#718096;line-height:1.6;text-align:center">
      Texas Power Guru is independently operated and not affiliated with any electricity provider.
      No commissions collected from any source. All plan data sourced from the Public Utility Commission
      of Texas website PowerToChoose.org. Cost estimates use projected monthly usage and published EFL rates.
      Actual bills may vary. Always review the Electricity Facts Label (EFL) before enrolling.
    </div>
  </div>

</div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Report saved: {output_path}")
    return output_path


# ── EXAMPLE ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    generate_report(
        client_name   = "Sample Client",
        tdu           = "ONCOR ELECTRIC DELIVERY COMPANY",
        monthly_usage = [900,750,700,650,800,1200,1800,2000,1600,1000,750,800],
        output_path   = "/mnt/user-data/outputs/sample_energy_report.html"
    )
