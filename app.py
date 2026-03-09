"""
FBA Smart Supply Planner
========================
Upload:
  1. Amazon Inventory Ledger  (CSV / ZIP / XLSX)
  2. MTR Sales Report         (CSV / ZIP / XLSX, multiple files OK)

Outputs:
  - FBA / FBM dispatch plan with safety stock & priority scores
  - FC-wise & cluster-wise allocation
  - Risk alerts (critical / dead / slow / excess)
  - Damaged stock report
  - Amazon FBA flat file for bulk shipment creation
  - Full Excel intelligence report
"""

import math
import io
import zipfile
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

# ──────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FBA Smart Supply Planner",
    layout="wide",
    page_icon="📦",
)

st.markdown(
    """
<style>
    [data-testid="stMetricValue"] { font-size: 1.55rem; font-weight: 700; }
    .pill-red    { display:inline-block; background:#fed7d7; color:#9b2c2c;
                   border-radius:12px; padding:2px 10px; font-size:12px; font-weight:700; }
    .pill-green  { display:inline-block; background:#c6f6d5; color:#276221;
                   border-radius:12px; padding:2px 10px; font-size:12px; font-weight:700; }
    .pill-yellow { display:inline-block; background:#fefcbf; color:#744210;
                   border-radius:12px; padding:2px 10px; font-size:12px; font-weight:700; }
    .alert-red    { background:#fff5f5; border-left:4px solid #e53e3e; color:#742a2a;
                    padding:.65rem 1rem; border-radius:6px; margin:.35rem 0; font-weight:600; }
    .alert-yellow { background:#fffff0; border-left:4px solid #d69e2e; color:#744210;
                    padding:.65rem 1rem; border-radius:6px; margin:.35rem 0; font-weight:600; }
    .alert-green  { background:#f0fff4; border-left:4px solid #38a169; color:#1c4532;
                    padding:.65rem 1rem; border-radius:6px; margin:.35rem 0; font-weight:600; }
    .stTabs [data-baseweb="tab"] { font-size:13px; font-weight:600; }
    div[data-testid="stExpander"] summary { font-weight:600; }
</style>
""",
    unsafe_allow_html=True,
)

st.title("📦 FBA Smart Supply Planner")
st.caption(
    "Amazon Inventory Ledger + MTR Sales → Dispatch plan · FC allocation · Risk alerts · Amazon flat file"
)

# ──────────────────────────────────────────────────────────────
# FC MASTER  (Amazon India)
# ──────────────────────────────────────────────────────────────
FC_DATA = {
    # Delhi NCR
    "DEX3": {"name": "New Delhi FC (DEX3)",       "city": "New Delhi",  "state": "DELHI",          "cluster": "Delhi NCR"},
    "DEX8": {"name": "New Delhi FC (DEX8)",       "city": "New Delhi",  "state": "DELHI",          "cluster": "Delhi NCR"},
    "DEL4": {"name": "Delhi North FC (DEL4)",     "city": "Delhi",      "state": "DELHI",          "cluster": "Delhi NCR"},
    "DEL5": {"name": "Delhi North FC (DEL5)",     "city": "Delhi",      "state": "DELHI",          "cluster": "Delhi NCR"},
    "DEL6": {"name": "Manesar FC (DEL6)",         "city": "Manesar",    "state": "HARYANA",        "cluster": "Delhi NCR"},
    "DEL7": {"name": "Bilaspur FC (DEL7)",        "city": "Bilaspur",   "state": "HARYANA",        "cluster": "Delhi NCR"},
    "XDEL": {"name": "Delhi XL FC (XDEL)",        "city": "New Delhi",  "state": "DELHI",          "cluster": "Delhi NCR"},
    # North
    "LDH1": {"name": "Ludhiana FC (LDH1)",        "city": "Ludhiana",   "state": "PUNJAB",         "cluster": "North"},
    "JAI1": {"name": "Jaipur FC (JAI1)",          "city": "Jaipur",     "state": "RAJASTHAN",      "cluster": "North"},
    "LKO1": {"name": "Lucknow FC (LKO1)",         "city": "Lucknow",    "state": "UTTAR PRADESH",  "cluster": "North"},
    "AGR1": {"name": "Agra FC (AGR1)",            "city": "Agra",       "state": "UTTAR PRADESH",  "cluster": "North"},
    # Mumbai / West
    "BOM1": {"name": "Bhiwandi FC (BOM1)",        "city": "Bhiwandi",   "state": "MAHARASHTRA",    "cluster": "Mumbai West"},
    "BOM3": {"name": "Nashik FC (BOM3)",          "city": "Nashik",     "state": "MAHARASHTRA",    "cluster": "Mumbai West"},
    "BOM4": {"name": "Vasai FC (BOM4)",           "city": "Vasai",      "state": "MAHARASHTRA",    "cluster": "Mumbai West"},
    "BOM5": {"name": "Bhiwandi 2 FC (BOM5)",      "city": "Bhiwandi",   "state": "MAHARASHTRA",    "cluster": "Mumbai West"},
    "SAMB": {"name": "Mumbai West FC (SAMB)",     "city": "Mumbai",     "state": "MAHARASHTRA",    "cluster": "Mumbai West"},
    "PNQ1": {"name": "Pune FC (PNQ1)",            "city": "Pune",       "state": "MAHARASHTRA",    "cluster": "Mumbai West"},
    "PNQ2": {"name": "Pune FC (PNQ2)",            "city": "Pune",       "state": "MAHARASHTRA",    "cluster": "Mumbai West"},
    "XBOM": {"name": "Mumbai XL FC (XBOM)",       "city": "Bhiwandi",   "state": "MAHARASHTRA",    "cluster": "Mumbai West"},
    # Bangalore
    "BLR5": {"name": "Bangalore South FC (BLR5)", "city": "Bangalore",  "state": "KARNATAKA",      "cluster": "Bangalore"},
    "BLR6": {"name": "Bangalore FC (BLR6)",       "city": "Bangalore",  "state": "KARNATAKA",      "cluster": "Bangalore"},
    "SCJA": {"name": "Bangalore FC (SCJA)",       "city": "Bangalore",  "state": "KARNATAKA",      "cluster": "Bangalore"},
    "XSAB": {"name": "Bangalore XS FC (XSAB)",   "city": "Bangalore",  "state": "KARNATAKA",      "cluster": "Bangalore"},
    "SBLA": {"name": "Bangalore SBLA FC (SBLA)",  "city": "Bangalore",  "state": "KARNATAKA",      "cluster": "Bangalore"},
    # Chennai
    "MAA4": {"name": "Chennai FC (MAA4)",         "city": "Chennai",    "state": "TAMIL NADU",     "cluster": "Chennai"},
    "MAA5": {"name": "Chennai FC (MAA5)",         "city": "Chennai",    "state": "TAMIL NADU",     "cluster": "Chennai"},
    "SMAB": {"name": "Chennai SMAB FC (SMAB)",    "city": "Chennai",    "state": "TAMIL NADU",     "cluster": "Chennai"},
    # Hyderabad
    "HYD7": {"name": "Hyderabad FC (HYD7)",       "city": "Hyderabad",  "state": "TELANGANA",      "cluster": "Hyderabad"},
    "HYD8": {"name": "Hyderabad FC (HYD8)",       "city": "Hyderabad",  "state": "TELANGANA",      "cluster": "Hyderabad"},
    "HYD9": {"name": "Hyderabad FC (HYD9)",       "city": "Hyderabad",  "state": "TELANGANA",      "cluster": "Hyderabad"},
    # Kolkata / East
    "CCU1": {"name": "Kolkata FC (CCU1)",         "city": "Kolkata",    "state": "WEST BENGAL",    "cluster": "Kolkata East"},
    "CCU2": {"name": "Kolkata FC (CCU2)",         "city": "Kolkata",    "state": "WEST BENGAL",    "cluster": "Kolkata East"},
    "PAT1": {"name": "Patna FC (PAT1)",           "city": "Patna",      "state": "BIHAR",          "cluster": "Kolkata East"},
    # Gujarat
    "AMD1": {"name": "Ahmedabad FC (AMD1)",       "city": "Ahmedabad",  "state": "GUJARAT",        "cluster": "Gujarat West"},
    "AMD2": {"name": "Ahmedabad FC (AMD2)",       "city": "Ahmedabad",  "state": "GUJARAT",        "cluster": "Gujarat West"},
    "SUB1": {"name": "Surat FC (SUB1)",           "city": "Surat",      "state": "GUJARAT",        "cluster": "Gujarat West"},
}


def _fc(code: str, key: str, fallback: str) -> str:
    return FC_DATA.get(str(code).upper().strip(), {}).get(key, fallback)


def fc_name(code):    return _fc(code, "name",    f"FC {code}")
def fc_cluster(code): return _fc(code, "cluster", "Other")
def fc_state(code):   return _fc(code, "state",   "Unknown")
def fc_city(code):    return _fc(code, "city",    "Unknown")


# ──────────────────────────────────────────────────────────────
# UTILITY HELPERS
# ──────────────────────────────────────────────────────────────
def read_file(uploaded) -> pd.DataFrame:
    """Read CSV / ZIP(CSV) / XLSX into a DataFrame."""
    try:
        name = uploaded.name.lower()
        if name.endswith(".zip"):
            with zipfile.ZipFile(uploaded) as zf:
                for member in zf.namelist():
                    if member.lower().endswith(".csv"):
                        return pd.read_csv(zf.open(member), low_memory=False)
            return pd.DataFrame()
        if name.endswith((".xlsx", ".xls")):
            return pd.read_excel(uploaded)
        return pd.read_csv(uploaded, low_memory=False)
    except Exception as exc:
        st.error(f"Cannot read **{uploaded.name}**: {exc}")
        return pd.DataFrame()


def find_col(df: pd.DataFrame, exact: list, fuzzy: list = None) -> str:
    """Return the first matching column name (exact, then fuzzy)."""
    lower_map = {c.strip().lower(): c for c in df.columns}
    for e in exact:
        if e.strip().lower() in lower_map:
            return lower_map[e.strip().lower()]
    if fuzzy:
        for col in df.columns:
            col_lower = col.lower()
            if any(sub in col_lower for sub in fuzzy):
                return col
    return ""


def add_sno(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.drop(columns=["S.No"], errors="ignore", inplace=True)
    df.insert(0, "S.No", range(1, len(df) + 1))
    return df


def fmt(n) -> str:
    try:
        return f"{int(n):,}"
    except Exception:
        return str(n)


def health_tag(doc: float, pd_: int) -> str:
    if doc <= 0:        return "⚫ No Stock"
    if doc < 14:        return "🔴 Critical"
    if doc < 30:        return "🟠 Low"
    if doc < pd_:       return "🟢 Healthy"
    if doc < pd_ * 2:   return "🟡 Excess"
    return "🔵 Overstocked"


def velocity_tag(avg: float) -> str:
    if avg <= 0:   return "⚫ Dead"
    if avg < 0.5:  return "🔵 Slow"
    if avg < 2:    return "🟡 Medium"
    if avg < 5:    return "🟢 Fast"
    return "🔥 Top Seller"


def trunc(s, n=60) -> str:
    s = str(s)
    return s[:n] + "…" if len(s) > n else s


# ──────────────────────────────────────────────────────────────
# SIDEBAR  — planning controls & shipment settings
# ──────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Planning Controls")

    planning_days = st.number_input(
        "Planning Days (Coverage Target)", min_value=7, max_value=180, value=60, step=1
    )
    service_level = st.selectbox("Service Level", ["90%", "95%", "98%"])
    Z_MAP = {"90%": 1.28, "95%": 1.65, "98%": 2.05}
    z_val = Z_MAP[service_level]

    sales_basis = st.selectbox(
        "Sales Basis for Planning",
        ["Total Sales (full history)", "Last 30 Days", "Last 60 Days", "Last 90 Days"],
    )

    st.divider()
    st.subheader("🚚 Shipment Settings")
    ship_from_name = st.text_input("Ship-From Warehouse Name", placeholder="My Warehouse")
    ship_from_addr = st.text_area(
        "Ship-From Address", placeholder="123 Street, City, State - 400001"
    )
    default_case_qty = st.number_input("Default Units Per Carton", 1, 500, 12)
    case_packed  = st.checkbox("Case-Packed Shipment", value=False)
    prep_owner   = st.selectbox("Prep Ownership",  ["AMAZON", "SELLER"])
    label_owner  = st.selectbox("Label Ownership", ["AMAZON", "SELLER"])

    st.divider()
    st.subheader("🎯 Display Filters")
    min_dispatch  = st.number_input("Min Dispatch Units to Show", 0, 99999, 1)
    focus_cluster = st.multiselect(
        "Focus Clusters", sorted({v["cluster"] for v in FC_DATA.values()})
    )

# ──────────────────────────────────────────────────────────────
# FILE UPLOADS
# ──────────────────────────────────────────────────────────────
st.markdown("### 📁 Upload Files")
up1, up2 = st.columns(2)
with up1:
    mtr_files = st.file_uploader(
        "📊 MTR / Sales Report — multiple files OK (CSV, ZIP, XLSX)",
        type=["csv", "zip", "xlsx"],
        accept_multiple_files=True,
    )
with up2:
    inv_file = st.file_uploader(
        "🏭 Amazon Inventory Ledger (CSV, ZIP, XLSX)",
        type=["csv", "zip", "xlsx"],
    )

with st.expander("📋 Expected File Formats"):
    st.markdown(
        """
**Inventory Ledger** — Seller Central → Reports → Fulfillment → Inventory Ledger:

| Column | Example |
|--------|---------|
| `Date` | 02/12/2026 |
| `FNSKU` | X00275KJZP |
| `MSKU` | BR-899 |
| `Title` | GLOOYA Bracelet... |
| `Disposition` | SELLABLE / CUSTOMER_DAMAGED |
| `Ending Warehouse Balance` | 1 |
| `Location` | LKO1 |

**MTR / Sales Report** — Seller Central → Reports → Tax → MTR:

| Column | Example |
|--------|---------|
| `Sku` / `MSKU` | BR-899 |
| `Quantity` | 5 |
| `Shipment Date` | 2026-01-15 |
| `Ship To State` | UTTAR PRADESH |
| `Fulfilment` | AFN / MFN |
"""
    )

# ──────────────────────────────────────────────────────────────
# EARLY EXIT — show FC reference if files not yet uploaded
# ──────────────────────────────────────────────────────────────
if not (mtr_files and inv_file):
    st.info("👆 Upload both files above to start the planner.")
    with st.expander("🗺️ Supported Amazon India FC Locations"):
        st.dataframe(
            pd.DataFrame(
                [{"FC Code": k, "FC Name": v["name"], "City": v["city"],
                  "State": v["state"], "Cluster": v["cluster"]}
                 for k, v in FC_DATA.items()]
            ),
            use_container_width=True,
            hide_index=True,
        )
    st.stop()

# ══════════════════════════════════════════════════════════════
# ①  LOAD & PARSE  INVENTORY LEDGER
# ══════════════════════════════════════════════════════════════
inv_raw = read_file(inv_file)
if inv_raw.empty:
    st.error("Inventory file could not be read. Please check the file.")
    st.stop()

inv_raw.columns = inv_raw.columns.str.strip()

# Detect columns
C_INV_SKU   = find_col(inv_raw, ["MSKU", "Sku", "SKU", "ASIN"],                          ["msku", "sku", "asin"])
C_INV_TITLE = find_col(inv_raw, ["Title", "Product Name", "Description"],                 ["title", "product"])
C_INV_QTY   = find_col(inv_raw, ["Ending Warehouse Balance", "Quantity", "Qty"],          ["ending", "balance"])
C_INV_LOC   = find_col(inv_raw, ["Location", "Warehouse Code", "FC Code", "FC"],          ["location", "warehouse code", "fc code"])
C_INV_DISP  = find_col(inv_raw, ["Disposition"],                                          ["disposition"])
C_INV_FNSKU = find_col(inv_raw, ["FNSKU"],                                                ["fnsku"])
C_INV_DATE  = find_col(inv_raw, ["Date"],                                                 ["date"])

# Validate required columns
inv_missing = [(n, c) for n, c in [("SKU/MSKU", C_INV_SKU),
                                    ("Ending Warehouse Balance", C_INV_QTY),
                                    ("Location/FC", C_INV_LOC)] if not c]
if inv_missing:
    st.error("Inventory file is missing required columns: " +
             ", ".join(n for n, _ in inv_missing))
    with st.expander("Detected columns in your inventory file"):
        st.write(list(inv_raw.columns))
    st.stop()

# Build working inventory frame
inv = inv_raw.copy()
inv["MSKU"]        = inv[C_INV_SKU].astype(str).str.strip()
inv["Stock"]       = pd.to_numeric(inv[C_INV_QTY], errors="coerce").fillna(0)
inv["FC Code"]     = inv[C_INV_LOC].astype(str).str.upper().str.strip()
inv["Title"]       = inv[C_INV_TITLE].astype(str).str.strip() if C_INV_TITLE else ""
inv["FNSKU"]       = inv[C_INV_FNSKU].astype(str).str.strip() if C_INV_FNSKU else ""
inv["Disposition"] = (inv[C_INV_DISP].astype(str).str.upper().str.strip()
                      if C_INV_DISP else "SELLABLE")

# ── CRITICAL FIX ──────────────────────────────────────────────
# Amazon Inventory Ledger has ONE ROW PER DAY per SKU/Disposition/FC.
# Summing all rows multiplies stock by the number of days in the report.
# Correct approach: keep only the LATEST date's row per group.
# ──────────────────────────────────────────────────────────────
if C_INV_DATE:
    inv["_dt"] = pd.to_datetime(inv[C_INV_DATE], dayfirst=True, errors="coerce")
    latest_inv_date = inv["_dt"].max()
    inv = (
        inv.sort_values("_dt")
           .groupby(["MSKU", "Disposition", "FC Code"], as_index=False)
           .last()
           .drop(columns=["_dt"])
    )
    inv_date_label = latest_inv_date.strftime("%d %b %Y") if pd.notna(latest_inv_date) else "unknown"
else:
    # No date column — just deduplicate by keeping last row
    inv = inv.groupby(["MSKU", "Disposition", "FC Code"], as_index=False).last()
    inv_date_label = "latest row"

# Enrich with FC master data
inv["FC Name"]    = inv["FC Code"].apply(fc_name)
inv["FC Cluster"] = inv["FC Code"].apply(fc_cluster)
inv["FC State"]   = inv["FC Code"].apply(fc_state)

# Lookup dictionaries
fnsku_map = (
    inv[inv["FNSKU"].notna() & ~inv["FNSKU"].isin(["nan", ""])]
    .drop_duplicates("MSKU")
    .set_index("MSKU")["FNSKU"]
    .to_dict()
)
title_map = inv.drop_duplicates("MSKU").set_index("MSKU")["Title"].to_dict()

# Split sellable vs damaged/non-sellable
inv_sellable = inv[inv["Disposition"] == "SELLABLE"].copy()
inv_damaged  = inv[inv["Disposition"] != "SELLABLE"].copy()

# Stock aggregates
fc_stock = (
    inv_sellable
    .groupby(["MSKU", "FC Code", "FC Name", "FC Cluster", "FC State"])["Stock"]
    .sum()
    .reset_index()
    .rename(columns={"Stock": "FC Stock"})
)

sku_sellable_stock = (
    inv_sellable.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock": "Sellable Stock"})
)

sku_damaged_stock = (
    inv_damaged.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock": "Damaged Stock"})
)

disp_breakdown = (
    inv.groupby(["MSKU", "Disposition"])["Stock"]
    .sum().unstack(fill_value=0).reset_index()
)

damaged_report = (
    inv_damaged.groupby(["MSKU", "Disposition"])["Stock"]
    .sum().reset_index().rename(columns={"Stock": "Qty"})
)
damaged_report["Title"] = damaged_report["MSKU"].map(title_map).fillna("")

# ══════════════════════════════════════════════════════════════
# ②  LOAD & PARSE  SALES / MTR
# ══════════════════════════════════════════════════════════════
raw_list = []
for f in mtr_files:
    df = read_file(f)
    if not df.empty:
        raw_list.append(df)

if not raw_list:
    st.error("Could not read any MTR / sales file.")
    st.stop()

sales_raw = pd.concat(raw_list, ignore_index=True)

C_SKU   = find_col(sales_raw, ["Sku", "SKU", "MSKU", "Item SKU"],              ["sku", "msku"])
C_QTY   = find_col(sales_raw, ["Quantity", "Qty", "Units"],                    ["qty", "unit", "quant"])
C_DATE  = find_col(sales_raw, ["Shipment Date", "Purchase Date", "Order Date", "Date"], ["date"])
C_STATE = find_col(sales_raw, ["Ship To State", "Shipping State"],              ["ship to", "state"])
C_FT    = find_col(sales_raw,
                   ["Fulfilment", "Fulfillment", "Fulfilment Channel", "Fulfillment Channel"],
                   ["fulfil", "channel"])

sales_missing = [(n, c) for n, c in [("SKU", C_SKU), ("Quantity", C_QTY), ("Date", C_DATE)]
                 if not c]
if sales_missing:
    st.error("Sales file is missing required columns: " +
             ", ".join(n for n, _ in sales_missing))
    with st.expander("Detected columns in your sales file"):
        st.write(list(sales_raw.columns))
    st.stop()

sales = sales_raw.copy()
sales = sales.rename(columns={C_SKU: "MSKU", C_QTY: "Qty", C_DATE: "Date"})
sales["MSKU"] = sales["MSKU"].astype(str).str.strip()
sales["Qty"]  = pd.to_numeric(sales["Qty"], errors="coerce").fillna(0)
sales["Date"] = pd.to_datetime(sales["Date"], dayfirst=True, errors="coerce")
sales = sales.dropna(subset=["Date"])
sales = sales[sales["Qty"] > 0].copy()

sales["Ship To State"] = (
    sales[C_STATE].astype(str).str.upper().str.strip() if C_STATE else "UNKNOWN"
)

FT_MAP = {
    "AFN": "FBA", "FBA": "FBA", "AMAZON_FULFILLED": "FBA",
    "MFN": "FBM", "FBM": "FBM", "MERCHANT_FULFILLED": "FBM",
}
if C_FT:
    sales["Channel"] = (
        sales[C_FT].astype(str).str.upper().str.strip().map(FT_MAP).fillna("FBA")
    )
else:
    sales["Channel"] = "FBA"

# ── Sales window ────────────────────────────────────────────
s_min = sales["Date"].min()
s_max = sales["Date"].max()
uploaded_days = max((s_max - s_min).days + 1, 1)

if   "30" in sales_basis: window = 30
elif "60" in sales_basis: window = 60
elif "90" in sales_basis: window = 90
else:                     window = uploaded_days

cutoff    = s_max - pd.Timedelta(days=window - 1)
hist      = sales[sales["Date"] >= cutoff].copy() if window < uploaded_days else sales.copy()
hist_days = max((hist["Date"].max() - hist["Date"].min()).days + 1, 1)

# ── Demand aggregates ────────────────────────────────────────
ch_hist  = (hist.groupby(["MSKU", "Channel"])["Qty"].sum()
            .reset_index().rename(columns={"Qty": "Hist Sales"}))
ch_full  = (sales.groupby(["MSKU", "Channel"])["Qty"].sum()
            .reset_index().rename(columns={"Qty": "All-Time Sales"}))
ch_daily = hist.groupby(["MSKU", "Channel", "Date"])["Qty"].sum().reset_index()
ch_std   = (ch_daily.groupby(["MSKU", "Channel"])["Qty"].std()
            .reset_index().rename(columns={"Qty": "Demand StdDev"}))

sku_top_states = (
    hist.groupby(["MSKU", "Ship To State"])["Qty"].sum()
    .reset_index()
    .sort_values("Qty", ascending=False)
    .groupby("MSKU")["Ship To State"]
    .apply(lambda x: ", ".join(list(x)[:3]))
    .reset_index()
    .rename(columns={"Ship To State": "Top States"})
)

# ══════════════════════════════════════════════════════════════
# ③  PLANNING TABLE
# ══════════════════════════════════════════════════════════════
plan_keys = pd.concat(
    [ch_hist[["MSKU", "Channel"]],
     sku_sellable_stock.assign(Channel="FBA")[["MSKU", "Channel"]]],
    ignore_index=True,
).drop_duplicates()

plan = plan_keys.copy()
plan = plan.merge(ch_hist,             on=["MSKU", "Channel"], how="left")
plan = plan.merge(ch_full,             on=["MSKU", "Channel"], how="left")
plan = plan.merge(ch_std,              on=["MSKU", "Channel"], how="left")
plan = plan.merge(sku_sellable_stock,  on="MSKU",              how="left")
plan = plan.merge(sku_damaged_stock,   on="MSKU",              how="left")
plan = plan.merge(sku_top_states,      on="MSKU",              how="left")

for col in ["Hist Sales", "All-Time Sales", "Sellable Stock",
            "Damaged Stock", "Demand StdDev"]:
    plan[col] = pd.to_numeric(plan[col], errors="coerce").fillna(0)

plan["Title"]            = plan["MSKU"].map(title_map).fillna("")
plan["FNSKU"]            = plan["MSKU"].map(fnsku_map).fillna("")
plan["Sales Days Used"]  = hist_days
plan["Planning Days"]    = planning_days
plan["Avg Daily Sale"]   = (plan["Hist Sales"] / hist_days).round(4)
plan["Safety Stock"]     = (z_val * plan["Demand StdDev"] * math.sqrt(planning_days)).round(0)
plan["Base Requirement"] = (plan["Avg Daily Sale"] * planning_days).round(0)
plan["Required Stock"]   = (plan["Base Requirement"] + plan["Safety Stock"]).round(0)
plan["Dispatch Needed"]  = (plan["Required Stock"] - plan["Sellable Stock"]).clip(lower=0).round(0)

plan["Days of Cover"] = np.where(
    plan["Avg Daily Sale"] > 0,
    (plan["Sellable Stock"] / plan["Avg Daily Sale"]).round(1),
    np.where(plan["Sellable Stock"] > 0, 9999, 0),
)

plan["Health"]   = plan.apply(lambda r: health_tag(r["Days of Cover"], planning_days), axis=1)
plan["Velocity"] = plan["Avg Daily Sale"].apply(velocity_tag)

_max_avg = plan["Avg Daily Sale"].max() or 1
plan["Priority Score"] = plan.apply(
    lambda r: round(
        (max(0, (planning_days - min(r["Days of Cover"], planning_days)) / planning_days) * 0.65
         + min(r["Avg Daily Sale"] / _max_avg, 1) * 0.35) * 100, 1
    ) if r["Avg Daily Sale"] > 0 else 0,
    axis=1,
)

fba_plan = plan[plan["Channel"] == "FBA"].copy()
fbm_plan = plan[plan["Channel"] == "FBM"].copy()

DISP_COLS = ["MSKU", "FNSKU", "Title", "Avg Daily Sale", "Sellable Stock",
             "Safety Stock", "Required Stock", "Dispatch Needed",
             "Days of Cover", "Health", "Velocity", "Priority Score", "Top States"]

# ══════════════════════════════════════════════════════════════
# ④  FC-WISE PLAN
# ══════════════════════════════════════════════════════════════
fc_plan = fc_stock.merge(
    plan[["MSKU", "Channel", "Avg Daily Sale", "Dispatch Needed",
          "Sellable Stock", "Required Stock", "Hist Sales",
          "Title", "FNSKU", "Priority Score", "Days of Cover"]],
    on="MSKU", how="left",
)

fc_plan["FC Days of Cover"] = np.where(
    fc_plan["Avg Daily Sale"] > 0,
    (fc_plan["FC Stock"] / fc_plan["Avg Daily Sale"]).round(1),
    np.where(fc_plan["FC Stock"] > 0, 9999, 0),
)
fc_plan["FC Health"]   = fc_plan.apply(
    lambda r: health_tag(r["FC Days of Cover"], planning_days), axis=1
)
fc_plan["FC Priority"] = fc_plan["FC Days of Cover"].apply(
    lambda d: "🔴 Urgent" if d < 14 else ("🟠 Soon" if d < 30 else "🟢 OK")
)

_fc_total = (fc_plan.groupby(["MSKU", "Channel"])["FC Stock"].sum()
             .reset_index().rename(columns={"FC Stock": "_total_fc"}))
fc_plan = fc_plan.merge(_fc_total, on=["MSKU", "Channel"], how="left")
fc_plan["FC Share"]    = (fc_plan["FC Stock"] / fc_plan["_total_fc"].replace(0, 1)).fillna(1.0)
fc_plan["FC Dispatch"] = (fc_plan["Dispatch Needed"] * fc_plan["FC Share"]).round(0)
fc_plan.drop(columns=["_total_fc"], inplace=True)

# filtered view
fc_view = fc_plan.copy()
if focus_cluster:
    fc_view = fc_view[fc_view["FC Cluster"].isin(focus_cluster)]
fc_view = fc_view[fc_view["FC Dispatch"] >= min_dispatch]

FC_DISP_COLS = ["FC Code", "FC Name", "FC Cluster", "MSKU", "Title",
                "FC Stock", "Avg Daily Sale", "FC Days of Cover",
                "FC Health", "FC Dispatch", "FC Priority"]

# ══════════════════════════════════════════════════════════════
# ⑤  CLUSTER SUMMARY
# ══════════════════════════════════════════════════════════════
cluster_sum = (
    fc_plan.groupby("FC Cluster")
    .agg(
        FC_Names       =("FC Name",          lambda x: " | ".join(sorted(set(x)))),
        FC_Codes       =("FC Code",          lambda x: ", ".join(sorted(set(x)))),
        Total_Stock    =("FC Stock",         "sum"),
        Dispatch_Needed=("FC Dispatch",      "sum"),
        Avg_DOC        =("FC Days of Cover", "mean"),
        SKUs           =("MSKU",            "nunique"),
    )
    .reset_index()
    .rename(columns={
        "FC_Names": "FC Names", "FC_Codes": "FC Codes",
        "Total_Stock": "Total Stock", "Dispatch_Needed": "Dispatch Needed",
        "Avg_DOC": "Avg Days of Cover", "SKUs": "Unique SKUs",
    })
)
cluster_sum["Avg Days of Cover"] = cluster_sum["Avg Days of Cover"].round(1)
cluster_sum = cluster_sum.sort_values("Dispatch Needed", ascending=False)
cluster_sum = add_sno(cluster_sum)

# ══════════════════════════════════════════════════════════════
# ⑥  RISK TABLES
# ══════════════════════════════════════════════════════════════
df_critical = plan[(plan["Days of Cover"]  < 14) & (plan["Avg Daily Sale"] > 0)].copy()
df_dead      = plan[(plan["Avg Daily Sale"] == 0) & (plan["Sellable Stock"] > 0)].copy()
df_slow      = plan[(plan["Avg Daily Sale"] > 0)  & (plan["Days of Cover"] > 90)].copy()
df_excess    = plan[plan["Days of Cover"] > planning_days * 2].copy()
df_top20     = plan[plan["Avg Daily Sale"] > 0].nlargest(20, "Avg Daily Sale").copy()

# ══════════════════════════════════════════════════════════════
# ⑦  STATE / TREND ANALYTICS
# ══════════════════════════════════════════════════════════════
state_demand = (
    sales.groupby("Ship To State")["Qty"].sum()
    .reset_index().rename(columns={"Qty": "Total Units"})
    .sort_values("Total Units", ascending=False)
)
state_demand["% Share"] = (
    state_demand["Total Units"] / state_demand["Total Units"].sum() * 100
).round(1)
state_demand = add_sno(state_demand)

weekly_trend = (
    sales.assign(Week=sales["Date"].dt.to_period("W").astype(str))
    .groupby("Week")["Qty"].sum()
    .reset_index().rename(columns={"Qty": "Units Sold"})
    .sort_values("Week")
)
monthly_trend = (
    sales.assign(Month=sales["Date"].dt.to_period("M").astype(str))
    .groupby("Month")["Qty"].sum()
    .reset_index().rename(columns={"Qty": "Units Sold"})
    .sort_values("Month")
)

# ══════════════════════════════════════════════════════════════
# ⑧  EXECUTIVE DASHBOARD
# ══════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("### 📊 Executive Dashboard")

_total_skus     = plan["MSKU"].nunique()
_total_sellable = int(sku_sellable_stock["Sellable Stock"].sum())
_total_damaged  = int(sku_damaged_stock["Damaged Stock"].sum()) if not sku_damaged_stock.empty else 0
_total_dispatch = int(fba_plan["Dispatch Needed"].sum())
_avg_doc        = (plan[plan["Avg Daily Sale"] > 0]["Days of Cover"]
                   .replace(9999, np.nan).mean())

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Total SKUs",         fmt(_total_skus))
m2.metric("Sellable Stock",     fmt(_total_sellable))
m3.metric("Damaged Stock",      fmt(_total_damaged))
m4.metric("Units to Dispatch",  fmt(_total_dispatch))
m5.metric("🔴 Critical SKUs",   fmt(len(df_critical)))
m6.metric("Avg Days of Cover",  f"{_avg_doc:.0f}d" if pd.notna(_avg_doc) else "N/A")

# Inventory snapshot info
st.markdown(
    f'<div class="alert-green">✅ Inventory snapshot: latest date used = <b>{inv_date_label}</b> '
    f'| {len(inv)} rows after daily deduplication (was {len(inv_raw)} raw rows)</div>',
    unsafe_allow_html=True,
)
if len(df_critical):
    st.markdown(
        f'<div class="alert-red">🚨 {len(df_critical)} SKU(s) below 14 days of stock — replenish immediately!</div>',
        unsafe_allow_html=True,
    )
if len(df_dead):
    st.markdown(
        f'<div class="alert-yellow">⚠️ {len(df_dead)} SKU(s) have stock but ZERO sales — fix listing or liquidate.</div>',
        unsafe_allow_html=True,
    )
if len(df_excess):
    st.markdown(
        f'<div class="alert-yellow">📦 {len(df_excess)} SKU(s) overstocked (>{planning_days*2} days) — pause replenishment.</div>',
        unsafe_allow_html=True,
    )
if _total_damaged:
    st.markdown(
        f'<div class="alert-yellow">🔧 {fmt(_total_damaged)} damaged/non-sellable units — '
        f'raise a removal order or reimbursement claim in Seller Central.</div>',
        unsafe_allow_html=True,
    )

# ══════════════════════════════════════════════════════════════
# ⑨  MAIN TABS
# ══════════════════════════════════════════════════════════════
tabs = st.tabs([
    "🏠 Overview",
    "📦 FBA Plan",
    "📮 FBM Plan",
    "🏭 FC-Wise Plan",
    "🗺️ Cluster View",
    "🚨 Risk Alerts",
    "📈 Trends",
    "🌍 State Demand",
    "🔧 Damaged Stock",
    "📄 Amazon Flat File",
])

# ── 0  OVERVIEW ──────────────────────────────────────────────
with tabs[0]:
    st.subheader("🏠 Planning Overview")
    oc1, oc2 = st.columns(2)
    with oc1:
        fcs_found = ", ".join(sorted(inv["FC Code"].unique()))
        st.info(
            f"**Sales data:** {s_min.strftime('%d %b %Y')} → {s_max.strftime('%d %b %Y')} "
            f"({uploaded_days} days)\n\n"
            f"**Sales basis:** {sales_basis} ({hist_days} days used)\n\n"
            f"**Planning horizon:** {planning_days} days | "
            f"**Service level:** {service_level} (Z = {z_val})\n\n"
            f"**Total SKUs:** {_total_skus} | **FCs:** {fcs_found}"
        )
    with oc2:
        st.markdown("**🔝 Top 5 Most Urgent SKUs**")
        top5 = plan[plan["Dispatch Needed"] > 0].nlargest(5, "Priority Score")[
            ["MSKU", "Title", "Avg Daily Sale", "Days of Cover",
             "Dispatch Needed", "Priority Score"]
        ].copy()
        top5["Title"] = top5["Title"].apply(lambda x: trunc(x, 45))
        st.dataframe(top5, use_container_width=True, hide_index=True)

    st.markdown("**📊 Monthly Sales Trend**")
    if not monthly_trend.empty:
        st.bar_chart(monthly_trend.set_index("Month")["Units Sold"])

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**🗺️ Cluster Dispatch Summary**")
        st.dataframe(
            cluster_sum[["FC Cluster", "FC Names", "Total Stock",
                         "Dispatch Needed", "Avg Days of Cover", "Unique SKUs"]],
            use_container_width=True, hide_index=True,
        )
    with col_b:
        st.markdown("**🏭 Inventory by FC**")
        fc_inv_tbl = (
            fc_plan.groupby(["FC Code", "FC Name", "FC Cluster"])["FC Stock"]
            .sum().reset_index()
            .rename(columns={"FC Stock": "Total Sellable Stock"})
            .sort_values("Total Sellable Stock", ascending=False)
        )
        st.dataframe(fc_inv_tbl, use_container_width=True, hide_index=True)

# ── 1  FBA PLAN ──────────────────────────────────────────────
with tabs[1]:
    st.subheader("📦 FBA Planning")
    st.caption(
        f"{len(fba_plan)} SKUs | {planning_days}-day plan @ {service_level} service level"
    )
    fba_v = (
        fba_plan[fba_plan["Dispatch Needed"] >= min_dispatch]
        .sort_values("Priority Score", ascending=False)
        .copy()
    )
    fba_v["Title"] = fba_v["Title"].apply(lambda x: trunc(x, 60))
    st.dataframe(
        add_sno(fba_v[[c for c in DISP_COLS if c in fba_v.columns]]),
        use_container_width=True,
    )

# ── 2  FBM PLAN ──────────────────────────────────────────────
with tabs[2]:
    st.subheader("📮 FBM Planning")
    if fbm_plan.empty:
        st.info("No FBM SKUs found — all inventory appears to be FBA.")
    else:
        fbm_v = (
            fbm_plan[fbm_plan["Dispatch Needed"] >= min_dispatch]
            .sort_values("Priority Score", ascending=False)
            .copy()
        )
        fbm_v["Title"] = fbm_v["Title"].apply(lambda x: trunc(x, 60))
        st.dataframe(add_sno(fbm_v), use_container_width=True)

# ── 3  FC-WISE PLAN ──────────────────────────────────────────
with tabs[3]:
    st.subheader("🏭 FC / Warehouse Wise Dispatch Plan")
    st.caption(
        "Units to dispatch per FC — allocated proportionally by current stock share at each location."
    )
    fv = (
        fc_view.sort_values(["FC Cluster", "FC Dispatch"], ascending=[True, False])
        .copy()
    )
    fv["Title"] = fv["Title"].apply(lambda x: trunc(x, 50))
    st.dataframe(
        add_sno(fv[[c for c in FC_DISP_COLS if c in fv.columns]]),
        use_container_width=True,
    )

# ── 4  CLUSTER VIEW ──────────────────────────────────────────
with tabs[4]:
    st.subheader("🗺️ Cluster-Level Inventory View")
    st.dataframe(cluster_sum, use_container_width=True, hide_index=True)
    st.divider()
    for _, row in cluster_sum.iterrows():
        cl = row["FC Cluster"]
        with st.expander(
            f"📍 {cl}  —  {fmt(int(row['Dispatch Needed']))} units needed  |  {row['FC Names']}"
        ):
            cd = fc_plan[fc_plan["FC Cluster"] == cl].copy()
            cd["Title"] = cd["Title"].apply(lambda x: trunc(x, 55))
            st.dataframe(
                cd[["FC Code", "FC Name", "MSKU", "Title", "FC Stock",
                    "Avg Daily Sale", "FC Days of Cover", "FC Dispatch", "FC Priority"]],
                use_container_width=True, hide_index=True,
            )

# ── 5  RISK ALERTS ───────────────────────────────────────────
with tabs[5]:
    rt1, rt2, rt3, rt4 = st.tabs(
        ["🔴 Critical (<14d)", "⚫ Dead Stock", "🟡 Slow Moving (>90d)", "🔵 Excess"]
    )

    RISK_COLS = ["MSKU", "Title", "Channel", "Avg Daily Sale", "Sellable Stock",
                 "Days of Cover", "Dispatch Needed", "Health", "Top States"]

    def risk_table(df):
        d = df.copy()
        d["Title"] = d["Title"].apply(lambda x: trunc(x, 55))
        return add_sno(d[[c for c in RISK_COLS if c in d.columns]])

    with rt1:
        st.markdown(f"**{len(df_critical)} SKU(s) — order immediately**")
        if not df_critical.empty:
            st.dataframe(
                risk_table(df_critical.sort_values("Days of Cover")),
                use_container_width=True,
            )
        else:
            st.success("No critical SKUs — all stock levels are healthy.")

    with rt2:
        st.markdown(f"**{len(df_dead)} SKU(s) — stock exists but zero sales**")
        if not df_dead.empty:
            d2 = df_dead.copy()
            d2["Title"] = d2["Title"].apply(lambda x: trunc(x, 55))
            st.dataframe(
                add_sno(d2[["MSKU", "Title", "Channel", "Sellable Stock", "All-Time Sales"]]),
                use_container_width=True,
            )
        else:
            st.success("No dead stock found.")

    with rt3:
        st.markdown(f"**{len(df_slow)} SKU(s) — days of cover > 90**")
        if not df_slow.empty:
            st.dataframe(
                risk_table(df_slow.sort_values("Days of Cover", ascending=False)),
                use_container_width=True,
            )
        else:
            st.success("No slow-moving SKUs found.")

    with rt4:
        st.markdown(f"**{len(df_excess)} SKU(s) — cover > {planning_days * 2} days**")
        if not df_excess.empty:
            st.dataframe(
                risk_table(df_excess.sort_values("Days of Cover", ascending=False)),
                use_container_width=True,
            )
        else:
            st.success("No excess stock found.")

# ── 6  TRENDS ────────────────────────────────────────────────
with tabs[6]:
    st.subheader("📈 Sales Trends")
    tr1, tr2 = st.columns(2)
    with tr1:
        st.markdown("**Weekly Sales**")
        if not weekly_trend.empty:
            st.line_chart(weekly_trend.set_index("Week")["Units Sold"])
    with tr2:
        st.markdown("**Monthly Sales**")
        if not monthly_trend.empty:
            st.bar_chart(monthly_trend.set_index("Month")["Units Sold"])

    st.markdown("**🔥 Top 20 SKUs by Avg Daily Sale**")
    t20 = df_top20.copy()
    t20["Title"] = t20["Title"].apply(lambda x: trunc(x, 60))
    st.dataframe(
        add_sno(t20[["MSKU", "Title", "Avg Daily Sale", "Days of Cover",
                     "Sellable Stock", "Velocity", "Top States"]]),
        use_container_width=True,
    )

    st.divider()
    st.markdown("**🔍 SKU Drill-Down — Daily Sales Chart**")
    sku_options = sorted(sales["MSKU"].unique())
    sel_sku = st.selectbox("Select MSKU", sku_options)
    sku_daily = (
        sales[sales["MSKU"] == sel_sku]
        .groupby("Date")["Qty"].sum()
        .reset_index()
        .set_index("Date")
    )
    if not sku_daily.empty:
        st.caption(f"**{trunc(title_map.get(sel_sku, sel_sku), 90)}**")
        st.line_chart(sku_daily["Qty"])

# ── 7  STATE DEMAND ──────────────────────────────────────────
with tabs[7]:
    st.subheader("🌍 State-Level Demand Analysis")
    sd1, sd2 = st.columns([2, 1])
    with sd1:
        st.dataframe(state_demand, use_container_width=True, hide_index=True)
    with sd2:
        st.markdown("**State × SKU Matrix (Top 10 SKUs)**")
        top10_skus = plan.nlargest(10, "Avg Daily Sale")["MSKU"].tolist()
        sm = (
            hist[hist["MSKU"].isin(top10_skus)]
            .groupby(["Ship To State", "MSKU"])["Qty"].sum()
            .unstack(fill_value=0)
        )
        if not sm.empty:
            st.dataframe(sm, use_container_width=True)

# ── 8  DAMAGED STOCK ─────────────────────────────────────────
with tabs[8]:
    st.subheader("🔧 Damaged / Non-Sellable Stock")
    st.caption(
        "Raise a **Removal Order** or **Reimbursement Claim** in Seller Central for these units."
    )
    if damaged_report.empty:
        st.success("No damaged or non-sellable stock found.")
    else:
        dr = damaged_report.copy()
        dr["Title"] = dr["Title"].apply(lambda x: trunc(x, 65))
        st.dataframe(
            add_sno(dr.sort_values("Qty", ascending=False)),
            use_container_width=True,
        )
        st.divider()
        st.markdown("**Full Disposition Breakdown per MSKU**")
        db = disp_breakdown.copy()
        db.insert(1, "Title", db["MSKU"].map(title_map).fillna("").apply(lambda x: trunc(x, 60)))
        st.dataframe(add_sno(db), use_container_width=True)

# ── 9  AMAZON FLAT FILE ──────────────────────────────────────
with tabs[9]:
    st.subheader("📄 Amazon FBA Shipment Flat File Generator")
    st.info(
        "**How to use:**\n"
        "1. Review / edit the quantities in the table below\n"
        "2. Select target FC and click **Download**\n"
        "3. In Seller Central → **Send to Amazon → Create New Shipment → Upload a file**\n"
        "4. Upload the `.txt` file — Amazon creates shipments automatically ✅\n\n"
        "_Each FC requires a separate shipment. Use the per-FC buttons at the bottom._"
    )

    ff1, ff2, ff3 = st.columns(3)
    with ff1:
        shipment_name = st.text_input(
            "Shipment Name", value=f"Replenishment_{datetime.now().strftime('%Y%m%d')}"
        )
        ship_date = st.date_input(
            "Expected Ship Date", value=datetime.now().date() + timedelta(days=2)
        )
    with ff2:
        all_fcs = sorted(inv["FC Code"].unique())
        target_fc = st.selectbox(
            "Target FC", all_fcs, format_func=lambda x: f"{x} — {fc_name(x)}"
        )
    with ff3:
        st.markdown(f"**FC:** {fc_name(target_fc)}")
        st.markdown(f"**Cluster:** {fc_cluster(target_fc)}")
        st.markdown(f"**State:** {fc_state(target_fc)}")

    # Build editable dispatch table (FBA only, with units needed)
    flat_base = fba_plan[fba_plan["Dispatch Needed"] > 0].copy()
    flat_base["Units to Ship"]  = flat_base["Dispatch Needed"].astype(int)
    flat_base["Units Per Case"] = default_case_qty
    flat_base["No. of Cases"]   = np.ceil(
        flat_base["Units to Ship"] / default_case_qty
    ).astype(int)
    flat_base["Title Short"] = flat_base["Title"].apply(lambda x: trunc(x, 80))

    st.markdown(
        f"**{len(flat_base)} SKUs | Total: {fmt(flat_base['Units to Ship'].sum())} units to dispatch**"
    )

    edit_cols = ["MSKU", "FNSKU", "Title Short", "Units to Ship",
                 "Units Per Case", "No. of Cases"]
    edited = st.data_editor(
        flat_base[[c for c in edit_cols if c in flat_base.columns]].reset_index(drop=True),
        num_rows="dynamic",
        use_container_width=True,
    )

    def build_flat_file(df: pd.DataFrame, fc_code: str) -> str:
        """Generate Amazon FBA Shipment Creation flat file (tab-delimited)."""
        parts   = [p.strip() for p in ship_from_addr.split(",")]
        city    = parts[1] if len(parts) > 1 else ""
        st_zip  = parts[2] if len(parts) > 2 else ""
        st_code = (
            st_zip.split("-")[0].strip()[:2].upper()
            if "-" in st_zip else st_zip[:2].upper()
        )
        postal = st_zip.split("-")[1].strip() if "-" in st_zip else ""

        lines = ["TemplateType=FlatFileShipmentCreation\tVersion=2015.0403", ""]

        # Shipment header
        meta_h = ["ShipmentName", "ShipFromName", "ShipFromAddressLine1",
                  "ShipFromCity", "ShipFromStateOrProvinceCode", "ShipFromPostalCode",
                  "ShipFromCountryCode", "ShipmentStatus", "LabelPrepType",
                  "AreCasesRequired", "DestinationFulfillmentCenterId"]
        meta_v = [shipment_name, ship_from_name or "My Warehouse",
                  parts[0] if parts else "", city, st_code, postal, "IN",
                  "WORKING", label_owner,
                  "YES" if case_packed else "NO",
                  fc_code]
        lines += ["\t".join(meta_h), "\t".join(str(v) for v in meta_v), ""]

        # Item rows
        item_h = ["SellerSKU", "FNSKU", "QuantityShipped", "QuantityInCase",
                  "PrepOwner", "LabelOwner", "ItemDescription", "ExpectedDeliveryDate"]
        lines.append("\t".join(item_h))

        for _, row in df.iterrows():
            qty_in_case = (str(int(row.get("Units Per Case", default_case_qty)))
                           if case_packed else "")
            lines.append(
                "\t".join([
                    str(row["MSKU"]),
                    str(row.get("FNSKU", "")),
                    str(int(row["Units to Ship"])),
                    qty_in_case,
                    prep_owner,
                    label_owner,
                    str(row.get("Title Short", ""))[:200],
                    str(ship_date),
                ])
            )
        return "\n".join(lines)

    # Preview + download for selected FC
    flat_content = build_flat_file(edited, target_fc)
    with st.expander("👁️ Preview Flat File"):
        st.code(
            flat_content[:2500] + ("\n...(truncated)" if len(flat_content) > 2500 else "")
        )

    st.download_button(
        f"📥 Download Flat File → {target_fc} ({fc_name(target_fc)})",
        data=flat_content.encode("utf-8"),
        file_name=f"Amazon_FBA_{target_fc}_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
        mime="text/plain",
    )

    # Per-FC download buttons
    if len(all_fcs) > 1:
        st.divider()
        st.markdown("**📦 Download Separate Flat File per FC**")
        btn_cols = st.columns(min(len(all_fcs), 4))
        for i, fc_code in enumerate(all_fcs):
            fc_flat = build_flat_file(edited, fc_code)
            with btn_cols[i % 4]:
                st.download_button(
                    f"📥 {fc_code}\n{fc_name(fc_code)}",
                    data=fc_flat.encode("utf-8"),
                    file_name=f"Amazon_FBA_{fc_code}_{datetime.now().strftime('%Y%m%d')}.txt",
                    mime="text/plain",
                    key=f"fc_dl_{fc_code}",
                )

# ══════════════════════════════════════════════════════════════
# ⑩  EXCEL EXPORT
# ══════════════════════════════════════════════════════════════
st.markdown("---")


def build_excel() -> io.BytesIO:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb  = writer.book
        hdr = wb.add_format({
            "bold": True, "bg_color": "#1a365d", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })

        def write_sheet(df: pd.DataFrame, sheet_name: str):
            if df.empty:
                pd.DataFrame({"Note": ["No data for this sheet"]}).to_excel(
                    writer, sheet_name=sheet_name, index=False
                )
                return
            d = df.copy()
            if "Title" in d.columns:
                d["Title"] = d["Title"].astype(str).str[:80]
            d.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            ws = writer.sheets[sheet_name]
            for ci, cn in enumerate(d.columns):
                ws.write(0, ci, cn, hdr)
                col_w = max(
                    len(str(cn)),
                    d[cn].astype(str).str.len().max() if not d.empty else 10,
                )
                ws.set_column(ci, ci, min(col_w + 2, 45))
            ws.freeze_panes(2, 0)
            ws.autofilter(0, 0, len(d), len(d.columns) - 1)

        write_sheet(add_sno(fba_plan.sort_values("Priority Score", ascending=False)), "FBA Plan")
        write_sheet(add_sno(fbm_plan),                                                "FBM Plan")
        write_sheet(
            add_sno(fc_plan.sort_values(["FC Cluster", "FC Dispatch"],
                                         ascending=[True, False])),                   "FC Dispatch Plan",
        )
        write_sheet(cluster_sum,                                                       "Cluster Summary")
        write_sheet(add_sno(plan.sort_values("Priority Score", ascending=False)),     "All SKUs")
        write_sheet(add_sno(df_critical.sort_values("Days of Cover")),                "CRITICAL Stock")
        write_sheet(add_sno(df_dead),                                                 "Dead Stock")
        write_sheet(add_sno(df_slow.sort_values("Days of Cover", ascending=False)),   "Slow Moving")
        write_sheet(add_sno(df_excess.sort_values("Days of Cover", ascending=False)), "Excess Stock")
        write_sheet(add_sno(damaged_report.sort_values("Qty", ascending=False)),      "Damaged Stock")
        write_sheet(disp_breakdown,                                                    "Disposition Breakdown")
        write_sheet(state_demand,                                                      "State Demand")
        write_sheet(weekly_trend,                                                      "Weekly Trend")
        write_sheet(monthly_trend,                                                     "Monthly Trend")
        write_sheet(add_sno(df_top20),                                                "Top 20 SKUs")

    out.seek(0)
    return out


dl1, dl2 = st.columns(2)
with dl1:
    st.download_button(
        "📥 Download Full Intelligence Report (Excel)",
        data=build_excel(),
        file_name=f"FBA_Supply_Plan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with dl2:
    dispatch_csv = fba_plan[fba_plan["Dispatch Needed"] > 0][
        ["MSKU", "FNSKU", "Title", "Avg Daily Sale", "Sellable Stock",
         "Days of Cover", "Required Stock", "Dispatch Needed",
         "Priority Score", "Velocity", "Top States"]
    ].to_csv(index=False)
    st.download_button(
        "📋 Download Dispatch Plan (CSV)",
        data=dispatch_csv,
        file_name=f"Dispatch_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )
