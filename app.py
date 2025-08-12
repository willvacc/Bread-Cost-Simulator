# app.py — Bread Cost Simulator (tailored + tooltips)
from __future__ import annotations
import io, re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(page_title="Bread Cost Simulator", page_icon="🥖", layout="wide")

# ==============================
# Utilities
# ==============================
def to_float(x):
    if x is None:
        return np.nan
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x)
    s = s.replace(",", "").replace("$", "").strip()
    try:
        return float(s)
    except Exception:
        m = re.match(r"^\s*([+-]?\d+(\.\d+)?)\s*%\s*$", s)
        if m:
            return float(m.group(1)) / 100.0
        return np.nan

def read_any_excel(uploaded):
    name = uploaded.name.lower()
    try:
        if name.endswith(".xlsx"):
            return pd.read_excel(uploaded, header=None, engine="openpyxl")
        return pd.read_excel(uploaded, header=None)
    except Exception:
        return pd.read_excel(uploaded, header=None, engine="openpyxl")

def find_row(df, pattern):
    rx = re.compile(pattern, re.I)
    for i in range(df.shape[0]):
        row = " | ".join([str(x) for x in df.iloc[i].fillna("").tolist()])
        if rx.search(row):
            return i
    return -1

def next_section_or_blank(df, start_row, stop_patterns):
    stop_rx = [re.compile(p, re.I) for p in stop_patterns]
    r = start_row + 1
    last_nonempty = start_row
    while r < df.shape[0]:
        row_vals = [str(x).strip() for x in df.iloc[r].tolist()]
        row_join = " | ".join(row_vals)
        is_all_blank = all(v in ("", "nan", "None") for v in row_vals)
        is_stop = any(rx.search(row_join) for rx in stop_rx)
        if is_stop:
            break
        if not is_all_blank:
            last_nonempty = r
        r += 1
    return last_nonempty

def slice_table(df, header_row, start_row, end_row, expected_cols):
    header = df.iloc[header_row].astype(str).str.strip().tolist()
    tmp = df.iloc[start_row:end_row+1].copy()
    tmp.columns = header + [f"__extra_{i}" for i in range(len(tmp.columns)-len(header))]
    keep = [c for c in expected_cols if c in tmp.columns]
    return tmp[keep].copy()

def nonempty_rows(tbl):
    mask = ~(tbl.astype(str).apply(lambda s: s.str.strip().isin(["", "nan", "None"])).all(axis=1))
    return tbl.loc[mask]

# ==============================
# Parsing your sheet
# ==============================
def parse_costing_sheet(df):
    # Standard Formula Cost
    sfc_row = find_row(df, r"Standard\s+Formula\s+Cost")
    ing_header_row = sfc_row + 1
    ing_end = next_section_or_blank(df, sfc_row, [
        r"Processing\s+Materials", r"Packing\s+Materials", r"Standard\s+Direct\s+Labor", r"Indirect\s+Labor",
        r"Variable\s+Manufacturing\s+Cost", r"Total\s+Formula"
    ])
    ing_tbl = slice_table(
        df, ing_header_row, ing_header_row+1, ing_end,
        expected_cols=["Item #","Raw Material","UOM","Cost per Unit Weight","Formula Requirements","Formula Cost"]
    )
    ing_tbl = nonempty_rows(ing_tbl)
    if ing_tbl.empty:
        ing_tbl = slice_table(
            df, ing_header_row, ing_header_row+1, ing_end,
            expected_cols=["Item Number","Raw Material","UOM","Cost per Unit Weight","Formula Requirements","Formula Cost"]
        )
        ing_tbl = nonempty_rows(ing_tbl)

    # Units per Batch
    units_row = find_row(df, r"Units\s+per\s+Batch")
    units_per_batch = None
    if units_row >= 0:
        vals = [to_float(v) for v in df.iloc[units_row].tolist()]
        nums = [v for v in vals if isinstance(v, float) and not np.isnan(v)]
        if nums:
            units_per_batch = nums[-1]

    # Processing Materials
    pm_row = find_row(df, r"Processing\s+Materials")
    pm_header_row = pm_row + 1
    pm_end = next_section_or_blank(df, pm_row, [
        r"Packing\s+Materials", r"Standard\s+Direct\s+Labor", r"Indirect\s+Labor", r"Variable\s+Manufacturing\s+Cost"
    ])
    pm_tbl = slice_table(
        df, pm_header_row, pm_header_row+1, pm_end,
        expected_cols=["Item #","Material","UOM","Cost per Unit","Usage","Total Cost"]
    )
    if pm_tbl.empty:
        pm_tbl = slice_table(
            df, pm_header_row, pm_header_row+1, pm_end,
            expected_cols=["Item Number","Material","UOM","Cost per Unit","Usage","Total Cost"]
        )
    pm_tbl = nonempty_rows(pm_tbl)

    # Packing Materials
    pack_row = find_row(df, r"Packing\s+Materials")
    pack_header_row = pack_row + 1
    pack_end = next_section_or_blank(df, pack_row, [
        r"Standard\s+Direct\s+Labor", r"Indirect\s+Labor", r"Variable\s+Manufacturing\s+Cost", r"Total\s+Packaging\s+Cost"
    ])
    pack_tbl = slice_table(
        df, pack_header_row, pack_header_row+1, pack_end,
        expected_cols=["Item #","Description","UOM","Cost Per Unit","Number of Units","Total Packing Cost"]
    )
    if pack_tbl.empty:
        pack_tbl = slice_table(
            df, pack_header_row, pack_header_row+1, pack_end,
            expected_cols=["Item Number","Description","UOM","Cost Per Unit","Number of Units","Total Packing Cost"]
        )
    pack_tbl = nonempty_rows(pack_tbl)

    # Direct Labor
    dl_row = find_row(df, r"Standard\s+Direct\s+Labor\s+Cost")
    dl_header_row = dl_row + 1
    dl_end = next_section_or_blank(df, dl_row, [r"Indirect\s+Labor", r"Variable\s+Manufacturing\s+Cost"])
    dl_tbl = slice_table(
        df, dl_header_row, dl_header_row+1, dl_end,
        expected_cols=["Item #","Job Description","Hourly Rate","Number of Employees","Total D/L Cost"]
    )
    if dl_tbl.empty:
        dl_tbl = slice_table(
            df, dl_header_row, dl_header_row+1, dl_end,
            expected_cols=["Item Number","Job Description","Hourly Rate","Number of Employees","Total D/L Cost"]
        )
    dl_tbl = nonempty_rows(dl_tbl)

    units_per_hr = None
    labor_per_unit = None
    for r in range(dl_row, min(dl_row+25, df.shape[0])):
        row_vals = [str(x) for x in df.iloc[r].tolist()]
        row_join = " | ".join(row_vals)
        if re.search(r"Units\s+per\s+HR", row_join, re.I):
            nums = [to_float(v) for v in row_vals if not pd.isna(to_float(v))]
            if nums:
                units_per_hr = nums[-1]
        if re.search(r"Labor\s+Cost\s+per\s+Unit", row_join, re.I):
            nums = [to_float(v) for v in row_vals if not pd.isna(to_float(v))]
            if nums:
                labor_per_unit = nums[-1]
    if pd.isna(labor_per_unit) or labor_per_unit is None:
        total_dl_per_hr = dl_tbl["Total D/L Cost"].map(to_float).sum()
        if units_per_hr and units_per_hr > 0:
            labor_per_unit = float(total_dl_per_hr) / float(units_per_hr)

    # Indirect Labor
    il_row = find_row(df, r"Indirect\s+Labor\s+Cost")
    il_header_row = il_row + 1
    il_end = next_section_or_blank(df, il_row, [r"Variable\s+Manufacturing\s+Cost"])
    il_tbl = slice_table(
        df, il_header_row, il_header_row+1, il_end,
        expected_cols=["Item #","Description","Rate","Number","Cost"]
    )
    il_tbl = nonempty_rows(il_tbl)

    indirect_per_unit = None
    for r in range(il_row, min(il_row+25, df.shape[0])):
        row_vals = [str(x) for x in df.iloc[r].tolist()]
        if re.search(r"Total\s+Indirect\s+Labor\s+Cost\s+per\s+Unit", " | ".join(row_vals), re.I):
            nums = [to_float(v) for v in row_vals if not pd.isna(to_float(v))]
            if nums:
                indirect_per_unit = nums[-1]

    # Variable Mfg Cost (VME)
    vm_row = find_row(df, r"Variable\s+Manufacturing\s+Cost")
    vm_header_row = vm_row + 1
    vm_end = next_section_or_blank(df, vm_row, [r"Total\s+Material\s+Cost", r"Total\s+Conversion\s+Cost", r"Over\s*Head\s*Cost\s*per\s*Unit"])
    vm_tbl = slice_table(
        df, vm_header_row, vm_header_row+1, vm_end,
        expected_cols=["Item #","Description","Rate","Cost"]
    )
    vm_tbl = nonempty_rows(vm_tbl)

    efficiency = 0.90
    scrap = 0.05
    for _, row in vm_tbl.iterrows():
        name = str(row.get("Description", "")).strip().lower()
        rate = to_float(row.get("Rate", np.nan))
        if "eff" in name and not pd.isna(rate):
            efficiency = rate if rate <= 1.0 else rate/100.0
        if "scrap" in name and not pd.isna(rate):
            scrap = rate if rate <= 1.0 else rate/100.0

    vme_per_unit = None
    for r in range(vm_row, min(vm_row+40, df.shape[0])):
        row_vals = [str(x) for x in df.iloc[r].tolist()]
        if re.search(r"\bVME\s+Per\s+Unit\b", " | ".join(row_vals), re.I):
            nums = [to_float(v) for v in row_vals if not pd.isna(to_float(v))]
            if nums:
                vme_per_unit = nums[-1]

    # Overhead & Contingency
    overhead_per_unit = None
    contingency_per_unit = None
    for r in range(vm_end, min(vm_end+40, df.shape[0])):
        row_join = " | ".join([str(x) for x in df.iloc[r].tolist()])
        if re.search(r"Over\s*Head\s*Cost\s*per\s*Unit", row_join, re.I):
            nums = [to_float(v) for v in df.iloc[r].tolist() if not pd.isna(to_float(v))]
            if nums:
                overhead_per_unit = nums[-1]
        if re.search(r"Contingency\s+Cost\s+per\s+Package", row_join, re.I):
            nums = [to_float(v) for v in df.iloc[r].tolist() if not pd.isna(to_float(v))]
            if nums:
                contingency_per_unit = nums[-1]

    units_per_batch = units_per_batch or 651.5294118
    efficiency = efficiency if efficiency is not None else 0.90
    scrap = scrap if scrap is not None else 0.05

    # Normalize
    usage_rows = []
    for _, r in ing_tbl.iterrows():
        item = str(r.get("Raw Material","")).strip()
        cpu = to_float(r.get("Cost per Unit Weight", np.nan))
        use = to_float(r.get("Formula Requirements", np.nan))
        if item and not pd.isna(cpu) and not pd.isna(use):
            usage_rows.append(["Ingredients", item, float(cpu), float(use)])

    for _, r in pm_tbl.iterrows():
        item = str(r.get("Material","")).strip()
        cpu = to_float(r.get("Cost per Unit", np.nan))
        use = to_float(r.get("Usage", np.nan))
        if item and not pd.isna(cpu) and not pd.isna(use):
            usage_rows.append(["Processing Materials", item, float(cpu), float(use)])

    for _, r in pack_tbl.iterrows():
        item = str(r.get("Description","")).strip()
        cpu = to_float(r.get("Cost Per Unit", np.nan))
        num_units = to_float(r.get("Number of Units", np.nan))
        if item and not pd.isna(cpu) and not pd.isna(num_units):
            usage_rows.append(["Packaging", item, float(cpu), float(num_units * (units_per_batch or 1.0))])

    usage_df = pd.DataFrame(usage_rows, columns=["Category","Item","CostPerUnit","UsagePerBatch"])

    per_unit_rows = []
    if labor_per_unit is not None:
        per_unit_rows.append(["Direct Labor", "Direct Labor per unit", float(labor_per_unit)])
    if indirect_per_unit is not None:
        per_unit_rows.append(["Indirect Labor", "Indirect Labor per unit", float(indirect_per_unit)])
    if vme_per_unit is not None:
        per_unit_rows.append(["Variable Mfg (VME)", "VME per unit", float(vme_per_unit)])
    if overhead_per_unit is not None:
        per_unit_rows.append(["Overhead", "Overhead per unit", float(overhead_per_unit)])
    if contingency_per_unit is not None:
        per_unit_rows.append(["Contingency", "Contingency per unit", float(contingency_per_unit)])
    per_unit_df = pd.DataFrame(per_unit_rows, columns=["Category","Item","UnitCostPerUnit"])

    meta = {"units_per_batch": float(units_per_batch), "efficiency": float(efficiency), "scrap": float(scrap)}
    return usage_df, per_unit_df, meta

# ==============================
# Cost model & charts
# ==============================
CATEGORIES_ORDER = ["Ingredients","Processing Materials","Packaging","Direct Labor","Indirect Labor","Variable Mfg (VME)","Overhead","Contingency"]

def compute_unit_cost(rows: pd.DataFrame, efficiency: float, scrap: float, units_per_batch: float):
    df = rows.copy()
    df["CostPerBatch"] = np.where(df["UsagePerBatch"].notna(),
                                  df["CostPerUnit"] * df["UsagePerBatch"],
                                  0.0)
    good_units = max(1.0, units_per_batch * efficiency * (1.0 - scrap))
    per_unit_batch = df.groupby("Category")["CostPerBatch"].sum() / good_units
    per_unit_adders = rows["UnitCostPerUnit"].fillna(0).groupby(rows["Category"]).sum()
    breakdown = (per_unit_batch.add(per_unit_adders, fill_value=0.0)).reindex(CATEGORIES_ORDER, fill_value=0.0)
    total = breakdown.sum()
    return total, breakdown, good_units

def tornado_data(base_breakdown, swing_pct=0.10, top_n=6):
    s = base_breakdown.sort_values(key=lambda x: np.abs(x), ascending=False).head(top_n)
    return pd.DataFrame({"Category": s.index, "Low": s.values * (1 - swing_pct), "High": s.values * (1 + swing_pct)})

def monte_carlo(rows, efficiency, scrap, units_per_batch, shocks, runs=3000, seed=42):
    rng = np.random.default_rng(seed)
    totals = []
    for _ in range(runs):
        jitter = rows.copy()
        m = jitter["CostPerUnit"].notna()
        jitter.loc[m, "CostPerUnit"] = jitter.loc[m].apply(
            lambda r: r["CostPerUnit"] * (1 + rng.normal(0, shocks.get(r["Category"], 0.0))), axis=1
        )
        m2 = jitter["UnitCostPerUnit"].notna()
        jitter.loc[m2, "UnitCostPerUnit"] = jitter.loc[m2].apply(
            lambda r: r["UnitCostPerUnit"] * (1 + rng.normal(0, shocks.get(r["Category"], 0.0))), axis=1
        )
        tot, _, _ = compute_unit_cost(jitter, efficiency, scrap, units_per_batch)
        totals.append(tot)
    return np.array(totals)

def fmt_money(x):
    return f"${x:,.4f}" if abs(x) < 10 else f"${x:,.2f}"

# ==============================
# Sidebar controls (with tooltips)
# ==============================
st.sidebar.header("Upload & Assumptions")

# Glossary popover
with st.sidebar.popover("ℹ️ Glossary / Help"):
    st.markdown(
        "**VME (Variable Manufacturing Expense):** variable factory costs tied to output (OT premium, benefits %, supplies, variable utilities).\n\n"
        "**Contingency:** planned buffer added per unit to cover uncertainty (price spikes, scrap, downtime).\n\n"
        "**Efficiency:** share of time/materials that becomes good output (e.g., 0.90 = 90%).\n\n"
        "**Scrap:** proportion of units lost or reworked.\n\n"
        "**Breakdown chart:** where each $/unit goes by category.\n\n"
        "**Tornado:** ranks categories by impact when each moves ±X% independently.\n\n"
        "**Monte Carlo:** simulates random shocks to estimate best/median/worst case unit cost."
    )

file = st.sidebar.file_uploader(
    "Upload your costing sheet (.xls/.xlsx)",
    type=["xls","xlsx"],
    help="Drop in your existing workbook. The app auto-detects sections by name (Ingredients, Processing, Packing, Labor, VME, Overhead, Contingency)."
)

with st.sidebar.expander("Global Tweaks (% change)", expanded=False):
    pct_ing = st.number_input("Ingredients %", value=0.0, step=1.0, help="Applies to all ingredient line items.")
    pct_proc = st.number_input("Processing Materials %", value=0.0, step=1.0, help="Dusting flour, corn meal, oils, etc.")
    pct_pack = st.number_input("Packaging %", value=0.0, step=1.0, help="Sleeves, film, trays, etc.")
    pct_labor = st.number_input("Labor %", value=0.0, step=1.0, help="Direct + Indirect labor per unit.")
    pct_vme  = st.number_input("VME %", value=0.0, step=1.0, help="Variable manufacturing expenses per unit.")
    pct_over = st.number_input("Overhead %", value=0.0, step=1.0, help="Allocated overhead per unit.")
    pct_cont = st.number_input("Contingency %", value=0.0, step=1.0, help="Buffer per unit for uncertainty.")

eff = st.sidebar.slider(
    "Efficiency",
    min_value=0.5, max_value=1.0, value=0.90, step=0.01,
    help="Share of nominal output that becomes good sellable units. Higher efficiency lowers $/unit."
)
scrap = st.sidebar.slider(
    "Scrap",
    min_value=0.0, max_value=0.20, value=0.05, step=0.005,
    help="Fraction of units lost. Higher scrap spreads batch cost across fewer good units → higher $/unit."
)

with st.sidebar.expander("Sensitivity Settings", expanded=False):
    tornado_pct = st.number_input(
        "Tornado swing ±%",
        min_value=1.0, max_value=50.0, value=10.0, step=1.0,
        help="How much to move each category up/down (independently) to rank impact on total cost."
    )

with st.sidebar.expander("Monte Carlo (risk)", expanded=False):
    runs = st.number_input("Simulations", min_value=500, max_value=20000, value=3000, step=100,
                           help="More runs = smoother distribution, slower compute.")
    vol_ing = st.number_input("Ingredients σ%", value=5.0, step=0.5, help="Volatility applied to ingredient unit costs.")
    vol_proc = st.number_input("Processing σ%", value=3.0, step=0.5)
    vol_pack = st.number_input("Packaging σ%", value=2.0, step=0.5)
    vol_labor = st.number_input("Labor σ%", value=2.0, step=0.5)
    vol_over = st.number_input("Overhead σ%", value=2.0, step=0.5)
    vol_vme = st.number_input("VME σ%", value=2.0, step=0.5)
    vol_cont = st.number_input("Contingency σ%", value=1.0, step=0.5)

# ==============================
# Load & parse (multi-sheet select)
# ==============================
if not file:
    st.info("Upload your `.xls/.xlsx` costing sheet to parse it automatically.")
    st.stop()

# Read all sheets, let user choose
name = file.name.lower()
try:
    if name.endswith(".xlsx"):
        all_sheets = pd.read_excel(file, sheet_name=None, header=None, engine="openpyxl")
    else:
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)  # requires xlrd locally for .xls
except Exception:
    all_sheets = pd.read_excel(file, sheet_name=None, header=None, engine="openpyxl")

sheet_names = list(all_sheets.keys())
selected_sheet = st.selectbox("Select a sheet to parse", sheet_names, help="One product per sheet. Switch to compare products quickly.")
raw_df = all_sheets[selected_sheet]

usage_df, per_unit_df, meta = parse_costing_sheet(raw_df)

# ==============================
# Review/edit parsed tables
# ==============================
st.subheader(f"Parsed from workbook — **{selected_sheet}**")
c1, c2 = st.columns(2)
with c1:
    st.caption("Usage-based (spread over batch). Edit any cell to fine-tune.")
    usage_df = st.data_editor(usage_df, num_rows="dynamic", hide_index=True, key="usage_tbl")
with c2:
    st.caption("Per-unit adders (already $/unit).")
    per_unit_df = st.data_editor(per_unit_df, num_rows="dynamic", hide_index=True, key="unit_tbl")

units_per_batch = st.number_input("Units per Batch (detected)", value=float(meta["units_per_batch"]), step=1.0,
                                  help="Nominal finished units per batch at 100% efficiency and 0% scrap.")
eff = st.slider("Efficiency (detected default)", 0.5, 1.0, float(meta["efficiency"]), 0.01,
                help="From your sheet’s Variable Manufacturing section (e.g., 90%).")
scrap = st.slider("Scrap (detected default)", 0.0, 0.2, float(meta["scrap"]), 0.005,
                  help="From your sheet’s Variable Manufacturing section (e.g., 5%).")

# Apply global % tweaks
def apply_pct(df, pct_map):
    out = df.copy()
    if "CostPerUnit" in out.columns:
        out["CostPerUnit"] = out.apply(lambda r: r["CostPerUnit"] * (1 + pct_map.get(r["Category"],0.0)), axis=1)
    if "UnitCostPerUnit" in out.columns:
        out["UnitCostPerUnit"] = out.apply(lambda r: r["UnitCostPerUnit"] * (1 + pct_map.get(r["Category"],0.0)), axis=1)
    return out

pct_map = {
    "Ingredients": pct_ing/100.0,
    "Processing Materials": pct_proc/100.0,
    "Packaging": pct_pack/100.0,
    "Direct Labor": pct_labor/100.0,
    "Indirect Labor": pct_labor/100.0,
    "Overhead": pct_over/100.0,
    "Variable Mfg (VME)": pct_vme/100.0,
    "Contingency": pct_cont/100.0,
}

usage_adj = apply_pct(usage_df, pct_map)
per_unit_adj = apply_pct(per_unit_df, pct_map)

# Merge compute frame
rows = pd.concat([
    usage_adj.assign(UnitCostPerUnit=np.nan),
    per_unit_adj.assign(CostPerUnit=np.nan, UsagePerBatch=np.nan),
], ignore_index=True)

# ==============================
# Compute & display
# ==============================
total, breakdown, good_units = compute_unit_cost(rows, eff, scrap, units_per_batch)

st.subheader("Results")
k1, k2, k3 = st.columns(3)
k1.metric("Total Cost / Unit", fmt_money(total))
k2.metric("Good Units / Batch", f"{good_units:,.0f}")
k3.metric("Nominal Units / Batch", f"{units_per_batch:,.0f}")

# Breakdown
fig = go.Figure()
fig.add_bar(
    x=breakdown.index,
    y=breakdown.values,
    name="Per-Unit Cost",
    hovertemplate="<b>%{x}</b><br>$ per unit: %{y:.4f}<extra></extra>"
)
fig.update_layout(
    title="Per-Unit Cost Breakdown",
    yaxis_title="$ per unit",
    xaxis_title="",
)
st.plotly_chart(fig, use_container_width=True)
st.caption("Breakdown shows where each dollar per unit goes. Use the % tweaks and efficiency/scrap sliders to see shifts.")

# Tornado sensitivity (±X%)
st.markdown("#### Sensitivity (Tornado)")
tor = tornado_data(breakdown, swing_pct=tornado_pct/100.0, top_n=6)
ft = go.Figure()
ft.add_bar(
    y=tor["Category"], x=tor["High"]-tor["Low"], base=tor["Low"], orientation="h",
    hovertemplate="<b>%{y}</b><br>Low: $%{base:.4f}<br>High: $%{x:%f}+base<extra></extra>"
)
ft.update_layout(
    xaxis_title="$ per unit",
    yaxis_title="",
    title=f"Tornado Chart — impact of ±{tornado_pct:.0f}% per category (independently)"
)
st.plotly_chart(ft, use_container_width=True)
st.caption("Wider bars = bigger impact on total cost when that category moves ± the chosen %.")

# Monte Carlo
st.markdown("#### Risk (Monte Carlo)")
shocks = {
    "Ingredients": vol_ing/100.0,
    "Processing Materials": vol_proc/100.0,
    "Packaging": vol_pack/100.0,
    "Direct Labor": vol_labor/100.0,
    "Indirect Labor": vol_labor/100.0,
    "Overhead": vol_over/100.0,
    "Variable Mfg (VME)": vol_vme/100.0,
    "Contingency": vol_cont/100.0,
}
mc = monte_carlo(rows, eff, scrap, units_per_batch, shocks, runs=int(runs))
p5, p50, p95 = np.percentile(mc, [5,50,95])

fh = go.Figure()
fh.add_histogram(x=mc, nbinsx=40, name="Simulated $/unit", hovertemplate="$%{x:.4f}<extra></extra>")
fh.add_vline(x=p5, line_dash="dot", annotation_text=f"P5 {fmt_money(p5)}")
fh.add_vline(x=p50, line_dash="dash", annotation_text=f"Median {fmt_money(p50)}")
fh.add_vline(x=p95, line_dash="dot", annotation_text=f"P95 {fmt_money(p95)}")
fh.update_layout(title="Monte Carlo: Distribution of Cost per Unit", xaxis_title="$ per unit", yaxis_title="Count")
st.plotly_chart(fh, use_container_width=True)
st.info(
    f"Risk band (5–95%): {fmt_money(p5)} – {fmt_money(p95)} • Median: {fmt_money(p50)}\n\n"
    f"**P5** = 5th percentile (best case, only 5% chance cost is lower)\n"
    f"**P50** = Median (most likely cost)\n"
    f"**P95** = 95th percentile (worst case, only 5% chance cost is higher)"
)

# Export
st.subheader("Export Scenario")
if st.button("Export Excel", help="Download current scenario, breakdown, and risk runs for sharing/audit."):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as xw:
        usage_adj.to_excel(xw, index=False, sheet_name="UsageBased")
        per_unit_adj.to_excel(xw, index=False, sheet_name="PerUnit")
        pd.DataFrame(
            {"Metric":["TotalCostPerUnit","GoodUnitsPerBatch","Efficiency","Scrap","Timestamp"],
             "Value":[total, good_units, eff, scrap, datetime.now().isoformat()]}
        ).to_excel(xw, index=False, sheet_name="Summary")
        breakdown.reset_index().rename(columns={"index":"Category",0:"PerUnit"}).to_excel(xw, index=False, sheet_name="Breakdown")
        pd.DataFrame({"MC_Run":np.arange(len(mc)), "CostPerUnit":mc}).to_excel(xw, index=False, sheet_name="MonteCarlo")
    st.download_button("Download bread_scenario.xlsx", data=out.getvalue(),
                       file_name="bread_scenario.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
