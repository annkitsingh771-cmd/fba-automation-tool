import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import math
from datetime import datetime, timedelta

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(page_title="FBA Smart Supply Planner", layout="wide", page_icon="📦")

st.markdown("""
<style>
    [data-testid="stMetricValue"] { font-size: 1.6rem; font-weight: 700; }
    .alert-red    { background:#fde8e8; border-left:4px solid #e53e3e; color:#742a2a;
                    padding:.6rem 1rem; border-radius:6px; margin:.4rem 0; font-weight:600; }
    .alert-yellow { background:#fefcbf; border-left:4px solid #d69e2e; color:#744210;
                    padding:.6rem 1rem; border-radius:6px; margin:.4rem 0; font-weight:600; }
    .alert-green  { background:#f0fff4; border-left:4px solid #38a169; color:#1c4532;
                    padding:.6rem 1rem; border-radius:6px; margin:.4rem 0; font-weight:600; }
    .stTabs [data-baseweb="tab"] { font-size:13px; font-weight:600; }
</style>
""", unsafe_allow_html=True)

st.title("📦 FBA Smart Supply Planner")
st.caption("Amazon Inventory Ledger + MTR Sales → Full dispatch plan, FC allocation & Amazon flat file")

# ─────────────────────────────────────────────
# FC MASTER DATA  (Amazon India)
# ─────────────────────────────────────────────
FC_DATA = {
    "DEX3":{"name":"New Delhi FC (DEX3)",       "city":"New Delhi",  "state":"DELHI",         "cluster":"Delhi NCR"},
    "DEX8":{"name":"New Delhi FC (DEX8)",       "city":"New Delhi",  "state":"DELHI",         "cluster":"Delhi NCR"},
    "DEL4":{"name":"Delhi North FC (DEL4)",     "city":"Delhi",      "state":"DELHI",         "cluster":"Delhi NCR"},
    "DEL5":{"name":"Delhi North FC (DEL5)",     "city":"Delhi",      "state":"DELHI",         "cluster":"Delhi NCR"},
    "DEL6":{"name":"Manesar FC (DEL6)",         "city":"Manesar",    "state":"HARYANA",       "cluster":"Delhi NCR"},
    "DEL7":{"name":"Bilaspur FC (DEL7)",        "city":"Bilaspur",   "state":"HARYANA",       "cluster":"Delhi NCR"},
    "XDEL":{"name":"Delhi XL FC (XDEL)",        "city":"New Delhi",  "state":"DELHI",         "cluster":"Delhi NCR"},
    "LDH1":{"name":"Ludhiana FC (LDH1)",        "city":"Ludhiana",   "state":"PUNJAB",        "cluster":"North"},
    "JAI1":{"name":"Jaipur FC (JAI1)",          "city":"Jaipur",     "state":"RAJASTHAN",     "cluster":"North"},
    "LKO1":{"name":"Lucknow FC (LKO1)",         "city":"Lucknow",    "state":"UTTAR PRADESH", "cluster":"North"},
    "AGR1":{"name":"Agra FC (AGR1)",            "city":"Agra",       "state":"UTTAR PRADESH", "cluster":"North"},
    "BOM1":{"name":"Bhiwandi FC (BOM1)",        "city":"Bhiwandi",   "state":"MAHARASHTRA",   "cluster":"Mumbai West"},
    "BOM3":{"name":"Nashik FC (BOM3)",          "city":"Nashik",     "state":"MAHARASHTRA",   "cluster":"Mumbai West"},
    "BOM4":{"name":"Vasai FC (BOM4)",           "city":"Vasai",      "state":"MAHARASHTRA",   "cluster":"Mumbai West"},
    "BOM5":{"name":"Bhiwandi 2 FC (BOM5)",      "city":"Bhiwandi",   "state":"MAHARASHTRA",   "cluster":"Mumbai West"},
    "SAMB":{"name":"Mumbai West FC (SAMB)",     "city":"Mumbai",     "state":"MAHARASHTRA",   "cluster":"Mumbai West"},
    "PNQ1":{"name":"Pune FC (PNQ1)",            "city":"Pune",       "state":"MAHARASHTRA",   "cluster":"Mumbai West"},
    "PNQ2":{"name":"Pune FC (PNQ2)",            "city":"Pune",       "state":"MAHARASHTRA",   "cluster":"Mumbai West"},
    "XBOM":{"name":"Mumbai XL FC (XBOM)",       "city":"Bhiwandi",   "state":"MAHARASHTRA",   "cluster":"Mumbai West"},
    "BLR5":{"name":"Bangalore South FC (BLR5)", "city":"Bangalore",  "state":"KARNATAKA",     "cluster":"Bangalore"},
    "BLR6":{"name":"Bangalore FC (BLR6)",       "city":"Bangalore",  "state":"KARNATAKA",     "cluster":"Bangalore"},
    "SCJA":{"name":"Bangalore FC (SCJA)",       "city":"Bangalore",  "state":"KARNATAKA",     "cluster":"Bangalore"},
    "XSAB":{"name":"Bangalore XS FC (XSAB)",   "city":"Bangalore",  "state":"KARNATAKA",     "cluster":"Bangalore"},
    "SBLA":{"name":"Bangalore SBLA FC (SBLA)",  "city":"Bangalore",  "state":"KARNATAKA",     "cluster":"Bangalore"},
    "MAA4":{"name":"Chennai FC (MAA4)",         "city":"Chennai",    "state":"TAMIL NADU",    "cluster":"Chennai"},
    "MAA5":{"name":"Chennai FC (MAA5)",         "city":"Chennai",    "state":"TAMIL NADU",    "cluster":"Chennai"},
    "SMAB":{"name":"Chennai SMAB FC (SMAB)",    "city":"Chennai",    "state":"TAMIL NADU",    "cluster":"Chennai"},
    "HYD7":{"name":"Hyderabad FC (HYD7)",       "city":"Hyderabad",  "state":"TELANGANA",     "cluster":"Hyderabad"},
    "HYD8":{"name":"Hyderabad FC (HYD8)",       "city":"Hyderabad",  "state":"TELANGANA",     "cluster":"Hyderabad"},
    "HYD9":{"name":"Hyderabad FC (HYD9)",       "city":"Hyderabad",  "state":"TELANGANA",     "cluster":"Hyderabad"},
    "CCU1":{"name":"Kolkata FC (CCU1)",         "city":"Kolkata",    "state":"WEST BENGAL",   "cluster":"Kolkata East"},
    "CCU2":{"name":"Kolkata FC (CCU2)",         "city":"Kolkata",    "state":"WEST BENGAL",   "cluster":"Kolkata East"},
    "PAT1":{"name":"Patna FC (PAT1)",           "city":"Patna",      "state":"BIHAR",         "cluster":"Kolkata East"},
    "AMD1":{"name":"Ahmedabad FC (AMD1)",       "city":"Ahmedabad",  "state":"GUJARAT",       "cluster":"Gujarat West"},
    "AMD2":{"name":"Ahmedabad FC (AMD2)",       "city":"Ahmedabad",  "state":"GUJARAT",       "cluster":"Gujarat West"},
    "SUB1":{"name":"Surat FC (SUB1)",           "city":"Surat",      "state":"GUJARAT",       "cluster":"Gujarat West"},
}

def fc_name(code):    return FC_DATA.get(code, {}).get("name",    f"FC {code}")
def fc_cluster(code): return FC_DATA.get(code, {}).get("cluster", "Other")
def fc_st(code):      return FC_DATA.get(code, {}).get("state",   "Unknown")

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def read_file(f):
    try:
        nm = f.name.lower()
        if nm.endswith(".zip"):
            with zipfile.ZipFile(f) as z:
                for m in z.namelist():
                    if m.lower().endswith(".csv"):
                        return pd.read_csv(z.open(m), low_memory=False)
        elif nm.endswith((".xlsx", ".xls")):
            return pd.read_excel(f)
        else:
            return pd.read_csv(f, low_memory=False)
    except Exception as e:
        st.error(f"File read error: {e}")
    return pd.DataFrame()

def find_col(df, exact_list, fuzzy_list=None):
    norm = {c.strip().lower(): c for c in df.columns}
    for e in exact_list:
        if e.strip().lower() in norm:
            return norm[e.strip().lower()]
    if fuzzy_list:
        for col in df.columns:
            if any(s in col.lower() for s in fuzzy_list):
                return col
    return ""

def serial(df):
    df = df.copy()
    if "S.No" in df.columns:
        df.drop(columns=["S.No"], inplace=True)
    df.insert(0, "S.No", range(1, len(df)+1))
    return df

def fmt(n):
    try: return f"{int(n):,}"
    except: return str(n)

def health_tag(doc, pd_):
    if doc <= 0:         return "⚫ No Stock"
    if doc < 14:         return "🔴 Critical"
    if doc < 30:         return "🟠 Low"
    if doc < pd_:        return "🟢 Healthy"
    if doc < pd_ * 2:   return "🟡 Excess"
    return "🔵 Overstocked"

def velocity_tag(avg):
    if avg <= 0:   return "⚫ Dead"
    if avg < 0.5:  return "🔵 Slow"
    if avg < 2:    return "🟡 Medium"
    if avg < 5:    return "🟢 Fast"
    return "🔥 Top Seller"

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Planning Controls")
    planning_days = st.number_input("Planning Days (Coverage Target)", 7, 180, 60, 1)
    service_level = st.selectbox("Service Level", ["90%", "95%", "98%"])
    Z_MAP = {"90%": 1.28, "95%": 1.65, "98%": 2.05}
    z_val = Z_MAP[service_level]
    sales_basis = st.selectbox("Sales Basis",
        ["Total Sales (full history)", "Last 30 Days", "Last 60 Days", "Last 90 Days"])

    st.divider()
    st.subheader("🚚 Shipment Settings")
    seller_id        = st.text_input("Seller ID", placeholder="A1B2C3D4E5F6G7")
    ship_from_name   = st.text_input("Ship-From Warehouse Name", placeholder="My Warehouse")
    ship_from_addr   = st.text_area("Ship-From Address",
                                    placeholder="123 Street, City, State - 400001")
    default_case_qty = st.number_input("Default Units Per Carton", 1, 500, 12)
    case_packed      = st.checkbox("Case-Packed Shipment", False)
    prep_owner       = st.selectbox("Prep Ownership",  ["AMAZON", "SELLER"])
    label_owner      = st.selectbox("Label Ownership", ["AMAZON", "SELLER"])

    st.divider()
    st.subheader("🎯 Filters")
    min_dispatch  = st.number_input("Min Dispatch Units to Show", 0, 10000, 1)
    show_clusters = st.multiselect("Focus Clusters",
        sorted(set(v["cluster"] for v in FC_DATA.values())))

# ─────────────────────────────────────────────
# FILE UPLOADS
# ─────────────────────────────────────────────
st.markdown("### 📁 Upload Files")
col_u1, col_u2 = st.columns(2)
with col_u1:
    mtr_files = st.file_uploader(
        "📊 MTR / Sales Report (CSV, ZIP, XLSX) — multiple OK",
        type=["csv", "zip", "xlsx"], accept_multiple_files=True)
with col_u2:
    inv_file = st.file_uploader(
        "🏭 Amazon Inventory Ledger (CSV, ZIP, XLSX)",
        type=["csv", "zip", "xlsx"])

with st.expander("📋 Expected File Formats"):
    st.markdown("""
**Inventory Ledger** — Amazon native export (Seller Central → Reports → Fulfillment → Inventory Ledger):

| Column | Example |
|--------|---------|
| FNSKU | X00275KJZP |
| ASIN | B0DNBJJ46R |
| MSKU | BR-899 |
| Title | GLOOYA Bracelet... |
| Disposition | SELLABLE / CUSTOMER_DAMAGED |
| Ending Warehouse Balance | 89 |
| Location | LKO1 |

**MTR / Sales Report** — Seller Central → Reports → Tax → MTR:

| Column | Example |
|--------|---------|
| Sku / MSKU | BR-899 |
| Quantity | 5 |
| Shipment Date | 2026-01-15 |
| Ship To State | UTTAR PRADESH |
| Fulfilment | AFN / MFN |
""")

if not (mtr_files and inv_file):
    st.info("👆 Upload both MTR and Inventory Ledger files to start.")
    with st.expander("🗺️ Supported Amazon India FC Locations"):
        fc_ref = pd.DataFrame([
            {"FC Code": k, "FC Name": v["name"], "City": v["city"],
             "State": v["state"], "Cluster": v["cluster"]}
            for k, v in FC_DATA.items()])
        st.dataframe(fc_ref, use_container_width=True, hide_index=True)
    st.stop()

# ═══════════════════════════════════════════════════════════
# LOAD INVENTORY LEDGER
# ═══════════════════════════════════════════════════════════
inv_raw = read_file(inv_file)
if inv_raw.empty:
    st.error("Could not read inventory file.")
    st.stop()
inv_raw.columns = inv_raw.columns.str.strip()

# Column detection — exact match to Amazon Inventory Ledger format
INV_SKU   = find_col(inv_raw, ["MSKU","Sku","SKU","ASIN"],                        ["msku","sku","asin"])
INV_TITLE = find_col(inv_raw, ["Title","Product Name","Description"],              ["title","product"])
INV_QTY   = find_col(inv_raw, ["Ending Warehouse Balance","Quantity","Qty"],       ["ending","balance","qty"])
INV_LOC   = find_col(inv_raw, ["Location","Warehouse Code","FC Code","FC"],        ["location","warehouse code","fc code"])
INV_DISP  = find_col(inv_raw, ["Disposition"],                                     ["disposition"])
INV_FNSKU = find_col(inv_raw, ["FNSKU"],                                           ["fnsku"])

for col_name, col_val in [("SKU/MSKU", INV_SKU), ("Ending Balance", INV_QTY), ("Location/FC", INV_LOC)]:
    if not col_val:
        st.error(f"Inventory file missing required column: **{col_name}**")
        st.write("Detected columns:", list(inv_raw.columns))
        st.stop()

inv = inv_raw.copy()
inv["MSKU"]        = inv[INV_SKU].astype(str).str.strip()
inv["Stock"]       = pd.to_numeric(inv[INV_QTY], errors="coerce").fillna(0)
inv["FC Code"]     = inv[INV_LOC].astype(str).str.upper().str.strip()
inv["Title"]       = inv[INV_TITLE].astype(str).str.strip() if INV_TITLE else ""
inv["FNSKU"]       = inv[INV_FNSKU].astype(str).str.strip() if INV_FNSKU else ""
inv["Disposition"] = inv[INV_DISP].astype(str).str.upper().str.strip() if INV_DISP else "SELLABLE"

# Enrich with FC master data
inv["FC Name"]    = inv["FC Code"].apply(fc_name)
inv["FC Cluster"] = inv["FC Code"].apply(fc_cluster)
inv["FC State"]   = inv["FC Code"].apply(fc_st)

# Lookup maps
fnsku_map = (inv[inv["FNSKU"].notna() & ~inv["FNSKU"].isin(["nan",""])]
             .drop_duplicates("MSKU").set_index("MSKU")["FNSKU"].to_dict())
title_map = inv.drop_duplicates("MSKU").set_index("MSKU")["Title"].to_dict()

# Split sellable vs damaged
inv_sellable = inv[inv["Disposition"] == "SELLABLE"].copy()
inv_damaged  = inv[inv["Disposition"] != "SELLABLE"].copy()

# Stock aggregates
fc_stock = (inv_sellable
    .groupby(["MSKU","FC Code","FC Name","FC Cluster","FC State"])["Stock"]
    .sum().reset_index().rename(columns={"Stock": "FC Stock"}))

sku_sellable_stock = (inv_sellable.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock": "Sellable Stock"}))

sku_damaged_stock = (inv_damaged.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock": "Damaged Stock"}))

disp_breakdown = (inv.groupby(["MSKU","Disposition"])["Stock"].sum()
    .unstack(fill_value=0).reset_index())

damaged_report = (inv_damaged.groupby(["MSKU","Disposition"])["Stock"].sum()
    .reset_index().rename(columns={"Stock": "Qty"}))
damaged_report["Title"] = damaged_report["MSKU"].map(title_map).fillna("")

# ═══════════════════════════════════════════════════════════
# LOAD SALES / MTR
# ═══════════════════════════════════════════════════════════
sales_list = []
for f in mtr_files:
    df = read_file(f)
    if not df.empty:
        sales_list.append(df)

if not sales_list:
    st.error("Could not read any MTR/sales file.")
    st.stop()

sales_raw = pd.concat(sales_list, ignore_index=True)

SKU_COL   = find_col(sales_raw, ["Sku","SKU","MSKU","Item SKU"],             ["sku","msku"])
QTY_COL   = find_col(sales_raw, ["Quantity","Qty","Units"],                  ["qty","unit","quant"])
DATE_COL  = find_col(sales_raw, ["Shipment Date","Purchase Date","Order Date","Date"], ["date"])
STATE_COL = find_col(sales_raw, ["Ship To State","Shipping State"],          ["ship to","state"])
FT_COL    = find_col(sales_raw, ["Fulfilment","Fulfillment","Fulfilment Channel","Fulfillment Channel"],
                                 ["fulfil","channel"])

for col_name, col_val in [("SKU", SKU_COL), ("Quantity", QTY_COL), ("Date", DATE_COL)]:
    if not col_val:
        st.error(f"Sales file missing required column: **{col_name}**")
        st.write("Detected columns:", list(sales_raw.columns))
        st.stop()

sales = sales_raw.rename(columns={SKU_COL:"MSKU", QTY_COL:"Qty", DATE_COL:"Date"}).copy()
sales["MSKU"] = sales["MSKU"].astype(str).str.strip()
sales["Qty"]  = pd.to_numeric(sales["Qty"], errors="coerce").fillna(0)
sales["Date"] = pd.to_datetime(sales["Date"], errors="coerce")
sales = sales.dropna(subset=["Date"])
sales = sales[sales["Qty"] > 0]

sales["Ship To State"] = (sales[STATE_COL].astype(str).str.upper().str.strip()
                          if STATE_COL else "UNKNOWN")

FT_MAP = {"AFN":"FBA","FBA":"FBA","AMAZON_FULFILLED":"FBA",
          "MFN":"FBM","FBM":"FBM","MERCHANT_FULFILLED":"FBM"}
sales["Channel"] = (sales[FT_COL].astype(str).str.upper().str.strip().map(FT_MAP).fillna("FBA")
                    if FT_COL else "FBA")

# ── Sales window ────────────────────────────────────────────
s_min = sales["Date"].min()
s_max = sales["Date"].max()
uploaded_days = max((s_max - s_min).days + 1, 1)

if   "30" in sales_basis: window = 30
elif "60" in sales_basis: window = 60
elif "90" in sales_basis: window = 90
else:                     window = uploaded_days

cutoff   = s_max - pd.Timedelta(days=window - 1)
hist     = sales[sales["Date"] >= cutoff].copy() if window < uploaded_days else sales.copy()
hist_days = max((hist["Date"].max() - hist["Date"].min()).days + 1, 1)

# ── Demand aggregates ───────────────────────────────────────
ch_hist  = (hist.groupby(["MSKU","Channel"])["Qty"].sum().reset_index()
            .rename(columns={"Qty":"Hist Sales"}))
ch_full  = (sales.groupby(["MSKU","Channel"])["Qty"].sum().reset_index()
            .rename(columns={"Qty":"All-Time Sales"}))
ch_daily = hist.groupby(["MSKU","Channel","Date"])["Qty"].sum().reset_index()
ch_std   = (ch_daily.groupby(["MSKU","Channel"])["Qty"].std().reset_index()
            .rename(columns={"Qty":"Demand StdDev"}))

sku_states = (hist.groupby(["MSKU","Ship To State"])["Qty"].sum()
    .reset_index().sort_values("Qty", ascending=False)
    .groupby("MSKU")["Ship To State"]
    .apply(lambda x: ", ".join(list(x)[:3]))
    .reset_index().rename(columns={"Ship To State":"Top States"}))

# ═══════════════════════════════════════════════════════════
# PLANNING TABLE
# ═══════════════════════════════════════════════════════════
plan_keys = pd.concat([
    ch_hist[["MSKU","Channel"]],
    sku_sellable_stock.assign(Channel="FBA")[["MSKU","Channel"]],
], ignore_index=True).drop_duplicates()

plan = plan_keys.copy()
plan = plan.merge(ch_hist,  on=["MSKU","Channel"], how="left")
plan = plan.merge(ch_full,  on=["MSKU","Channel"], how="left")
plan = plan.merge(ch_std,   on=["MSKU","Channel"], how="left")
plan = plan.merge(sku_sellable_stock, on="MSKU", how="left")
plan = plan.merge(sku_damaged_stock,  on="MSKU", how="left")
plan = plan.merge(sku_states, on="MSKU", how="left")

for c in ["Hist Sales","All-Time Sales","Sellable Stock","Damaged Stock","Demand StdDev"]:
    plan[c] = pd.to_numeric(plan[c], errors="coerce").fillna(0)

plan["Title"]           = plan["MSKU"].map(title_map).fillna("")
plan["FNSKU"]           = plan["MSKU"].map(fnsku_map).fillna("")
plan["Sales Days Used"] = hist_days
plan["Planning Days"]   = planning_days
plan["Avg Daily Sale"]  = (plan["Hist Sales"] / hist_days).round(4)
plan["Safety Stock"]    = (z_val * plan["Demand StdDev"] * math.sqrt(planning_days)).round(0)
plan["Base Requirement"]= (plan["Avg Daily Sale"] * planning_days).round(0)
plan["Required Stock"]  = (plan["Base Requirement"] + plan["Safety Stock"]).round(0)
plan["Dispatch Needed"] = (plan["Required Stock"] - plan["Sellable Stock"]).clip(lower=0).round(0)

plan["Days of Cover"] = np.where(
    plan["Avg Daily Sale"] > 0,
    (plan["Sellable Stock"] / plan["Avg Daily Sale"]).round(1),
    np.where(plan["Sellable Stock"] > 0, 9999, 0))

plan["Health"]   = plan.apply(lambda r: health_tag(r["Days of Cover"], planning_days), axis=1)
plan["Velocity"] = plan["Avg Daily Sale"].apply(velocity_tag)

max_avg = plan["Avg Daily Sale"].max() or 1
plan["Priority Score"] = plan.apply(lambda r: round(
    (max(0, (planning_days - min(r["Days of Cover"], planning_days)) / planning_days) * 0.65 +
     min(r["Avg Daily Sale"] / max_avg, 1) * 0.35) * 100, 1)
    if r["Avg Daily Sale"] > 0 else 0, axis=1)

fba_plan = plan[plan["Channel"] == "FBA"].copy()
fbm_plan = plan[plan["Channel"] == "FBM"].copy()

# ═══════════════════════════════════════════════════════════
# FC-WISE PLAN
# ═══════════════════════════════════════════════════════════
fc_plan = fc_stock.merge(
    plan[["MSKU","Channel","Avg Daily Sale","Dispatch Needed","Sellable Stock",
          "Required Stock","Hist Sales","Title","FNSKU","Priority Score","Days of Cover"]],
    on="MSKU", how="left")

fc_plan["FC Days of Cover"] = np.where(
    fc_plan["Avg Daily Sale"] > 0,
    (fc_plan["FC Stock"] / fc_plan["Avg Daily Sale"]).round(1),
    np.where(fc_plan["FC Stock"] > 0, 9999, 0))
fc_plan["FC Health"]   = fc_plan.apply(lambda r: health_tag(r["FC Days of Cover"], planning_days), axis=1)

fc_total = (fc_plan.groupby(["MSKU","Channel"])["FC Stock"].sum()
            .reset_index().rename(columns={"FC Stock":"Total FC Stock"}))
fc_plan  = fc_plan.merge(fc_total, on=["MSKU","Channel"], how="left")
fc_plan["FC Share"]    = np.where(fc_plan["Total FC Stock"] > 0,
    fc_plan["FC Stock"] / fc_plan["Total FC Stock"].replace(0, 1), 1.0)
fc_plan["FC Dispatch"] = (fc_plan["Dispatch Needed"] * fc_plan["FC Share"]).round(0)
fc_plan["FC Priority"] = fc_plan["FC Days of Cover"].apply(
    lambda d: "🔴 Urgent" if d < 14 else ("🟠 Soon" if d < 30 else "🟢 OK"))

fc_view = fc_plan.copy()
if show_clusters:
    fc_view = fc_view[fc_view["FC Cluster"].isin(show_clusters)]
fc_view = fc_view[fc_view["FC Dispatch"] >= min_dispatch]

# ═══════════════════════════════════════════════════════════
# CLUSTER SUMMARY
# ═══════════════════════════════════════════════════════════
cluster_sum = (fc_plan.groupby("FC Cluster").agg(
    FC_Names       =("FC Name",           lambda x: " | ".join(sorted(set(x)))),
    FC_Codes       =("FC Code",           lambda x: ", ".join(sorted(set(x)))),
    Total_Stock    =("FC Stock",          "sum"),
    Dispatch_Needed=("FC Dispatch",       "sum"),
    Avg_DOC        =("FC Days of Cover",  "mean"),
    SKUs           =("MSKU",             "nunique"),
).reset_index().rename(columns={
    "FC_Names":"FC Names","FC_Codes":"FC Codes",
    "Total_Stock":"Total Stock","Dispatch_Needed":"Dispatch Needed",
    "Avg_DOC":"Avg Days of Cover","SKUs":"Unique SKUs"}))
cluster_sum["Avg Days of Cover"] = cluster_sum["Avg Days of Cover"].round(1)
cluster_sum = cluster_sum.sort_values("Dispatch Needed", ascending=False)
cluster_sum = serial(cluster_sum)

# ═══════════════════════════════════════════════════════════
# RISK TABLES
# ═══════════════════════════════════════════════════════════
critical = plan[(plan["Days of Cover"] < 14) & (plan["Avg Daily Sale"] > 0)].copy()
dead     = plan[(plan["Avg Daily Sale"] == 0) & (plan["Sellable Stock"] > 0)].copy()
slow     = plan[(plan["Avg Daily Sale"] > 0)  & (plan["Days of Cover"] > 90)].copy()
excess   = plan[plan["Days of Cover"] > planning_days * 2].copy()
top20    = plan[plan["Avg Daily Sale"] > 0].nlargest(20, "Avg Daily Sale").copy()

# ═══════════════════════════════════════════════════════════
# STATE / TREND
# ═══════════════════════════════════════════════════════════
state_demand = (sales.groupby("Ship To State")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Total Units"}).sort_values("Total Units", ascending=False))
state_demand["% Share"] = (state_demand["Total Units"] /
                            state_demand["Total Units"].sum() * 100).round(1)
state_demand = serial(state_demand)

weekly = sales.copy()
weekly["Week"] = weekly["Date"].dt.to_period("W").astype(str)
weekly_trend = (weekly.groupby("Week")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units Sold"}).sort_values("Week"))

monthly = sales.copy()
monthly["Month"] = monthly["Date"].dt.to_period("M").astype(str)
monthly_trend = (monthly.groupby("Month")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units Sold"}).sort_values("Month"))

# ═══════════════════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("### 📊 Executive Dashboard")

total_skus    = plan["MSKU"].nunique()
total_stock   = int(sku_sellable_stock["Sellable Stock"].sum())
total_damaged = int(sku_damaged_stock["Damaged Stock"].sum()) if not sku_damaged_stock.empty else 0
total_dispatch= int(fba_plan["Dispatch Needed"].sum())
avg_doc       = (plan[plan["Avg Daily Sale"] > 0]["Days of Cover"]
                 .replace(9999, np.nan).mean())

m1,m2,m3,m4,m5,m6 = st.columns(6)
m1.metric("Total SKUs",        fmt(total_skus))
m2.metric("Sellable Stock",    fmt(total_stock))
m3.metric("Damaged Stock",     fmt(total_damaged))
m4.metric("Units to Dispatch", fmt(total_dispatch))
m5.metric("🔴 Critical SKUs",  fmt(len(critical)))
m6.metric("Avg Days of Cover", f"{avg_doc:.0f}d" if not np.isnan(avg_doc) else "N/A")

if len(critical):
    st.markdown(f'<div class="alert-red">🚨 {len(critical)} SKU(s) below 14 days of stock — replenish immediately!</div>',
                unsafe_allow_html=True)
if len(dead):
    st.markdown(f'<div class="alert-yellow">⚠️ {len(dead)} SKU(s) have stock but ZERO sales — fix listing or liquidate.</div>',
                unsafe_allow_html=True)
if len(excess):
    st.markdown(f'<div class="alert-yellow">📦 {len(excess)} SKU(s) overstocked (>{planning_days*2} days) — pause replenishment.</div>',
                unsafe_allow_html=True)
if total_damaged > 0:
    st.markdown(f'<div class="alert-yellow">🔧 {fmt(total_damaged)} damaged/non-sellable units — raise removal or reimbursement claim.</div>',
                unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════
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

# ── OVERVIEW ──────────────────────────────────────────────
with tabs[0]:
    st.subheader("🏠 Planning Overview")
    oc1, oc2 = st.columns(2)
    with oc1:
        fcs_detected = ", ".join(sorted(inv["FC Code"].unique()))
        st.info(f"""
**Sales data:** {s_min.strftime('%d %b %Y')} → {s_max.strftime('%d %b %Y')} ({uploaded_days} days)  
**Sales basis:** {sales_basis} ({hist_days} days used)  
**Planning horizon:** {planning_days} days  
**Service level:** {service_level} (Z = {z_val})  
**Total SKUs:** {total_skus}  
**FCs detected:** {fcs_detected}
        """)
    with oc2:
        st.markdown("**🔝 Top 5 Most Urgent SKUs**")
        top5 = plan[plan["Dispatch Needed"] > 0].nlargest(5, "Priority Score")[
            ["MSKU","Title","Avg Daily Sale","Days of Cover","Dispatch Needed","Priority Score"]]
        top5["Title"] = top5["Title"].str[:50]
        st.dataframe(top5, use_container_width=True, hide_index=True)

    st.markdown("**📊 Monthly Sales Trend**")
    if not monthly_trend.empty:
        st.bar_chart(monthly_trend.set_index("Month")["Units Sold"])

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**🗺️ Cluster Dispatch Summary**")
        st.dataframe(cluster_sum[["FC Cluster","FC Names","Total Stock",
                                   "Dispatch Needed","Avg Days of Cover","Unique SKUs"]],
                     use_container_width=True, hide_index=True)
    with col_b:
        st.markdown("**🏭 Inventory by FC**")
        fc_inv_sum = (fc_plan.groupby(["FC Code","FC Name","FC Cluster"])["FC Stock"]
                      .sum().reset_index()
                      .rename(columns={"FC Stock":"Total Sellable Stock"})
                      .sort_values("Total Sellable Stock", ascending=False))
        st.dataframe(fc_inv_sum, use_container_width=True, hide_index=True)

# ── FBA PLAN ──────────────────────────────────────────────
with tabs[1]:
    st.subheader("📦 FBA Planning")
    st.caption(f"{len(fba_plan)} SKUs | {planning_days} days @ {service_level} service level")
    disp_cols = ["MSKU","FNSKU","Title","Avg Daily Sale","Sellable Stock","Safety Stock",
                 "Required Stock","Dispatch Needed","Days of Cover","Health",
                 "Velocity","Priority Score","Top States"]
    fba_v = fba_plan[fba_plan["Dispatch Needed"] >= min_dispatch].sort_values("Priority Score", ascending=False).copy()
    fba_v["Title"] = fba_v["Title"].str[:60]
    st.dataframe(serial(fba_v[[c for c in disp_cols if c in fba_v.columns]]),
                 use_container_width=True)

# ── FBM PLAN ──────────────────────────────────────────────
with tabs[2]:
    st.subheader("📮 FBM Planning")
    if fbm_plan.empty:
        st.info("No FBM SKUs found — all inventory is FBA.")
    else:
        fbm_v = fbm_plan[fbm_plan["Dispatch Needed"] >= min_dispatch].sort_values("Priority Score", ascending=False).copy()
        fbm_v["Title"] = fbm_v["Title"].str[:60]
        st.dataframe(serial(fbm_v), use_container_width=True)

# ── FC-WISE PLAN ───────────────────────────────────────────
with tabs[3]:
    st.subheader("🏭 FC / Warehouse Wise Dispatch Plan")
    st.caption("Units to dispatch per FC — allocated by current stock distribution across FCs")
    fc_dc = ["FC Code","FC Name","FC Cluster","MSKU","Title","FC Stock",
             "Avg Daily Sale","FC Days of Cover","FC Health","FC Dispatch","FC Priority"]
    fv = fc_view.sort_values(["FC Cluster","FC Dispatch"], ascending=[True, False]).copy()
    fv["Title"] = fv["Title"].str[:50]
    st.dataframe(serial(fv[[c for c in fc_dc if c in fv.columns]]),
                 use_container_width=True)

# ── CLUSTER VIEW ───────────────────────────────────────────
with tabs[4]:
    st.subheader("🗺️ Cluster-Level Inventory View")
    st.dataframe(cluster_sum, use_container_width=True, hide_index=True)
    st.divider()
    for _, row in cluster_sum.iterrows():
        cl = row["FC Cluster"]
        label = (f"📍 {cl}  —  {fmt(int(row['Dispatch Needed']))} units needed"
                 f"  |  {row['FC Names']}")
        with st.expander(label):
            cd = fc_plan[fc_plan["FC Cluster"] == cl].copy()
            cd["Title"] = cd["Title"].str[:55]
            st.dataframe(
                cd[["FC Code","FC Name","MSKU","Title","FC Stock",
                    "Avg Daily Sale","FC Days of Cover","FC Dispatch","FC Priority"]],
                use_container_width=True, hide_index=True)

# ── RISK ALERTS ────────────────────────────────────────────
with tabs[5]:
    r1, r2, r3, r4 = st.tabs(["🔴 Critical (<14d)","⚫ Dead Stock","🟡 Slow (>90d)","🔵 Excess"])

    def risk_df(df):
        cols = ["MSKU","Title","Channel","Avg Daily Sale","Sellable Stock",
                "Days of Cover","Dispatch Needed","Health","Top States"]
        d = df.copy()
        d["Title"] = d["Title"].str[:55]
        return serial(d[[c for c in cols if c in d.columns]])

    with r1:
        st.markdown(f"**{len(critical)} SKUs — order immediately**")
        st.dataframe(risk_df(critical.sort_values("Days of Cover")), use_container_width=True)
    with r2:
        st.markdown(f"**{len(dead)} SKUs — stock exists, zero sales**")
        d2 = dead.copy(); d2["Title"] = d2["Title"].str[:55]
        st.dataframe(serial(d2[["MSKU","Title","Channel","Sellable Stock","All-Time Sales"]]),
                     use_container_width=True)
    with r3:
        st.markdown(f"**{len(slow)} SKUs — days of cover > 90**")
        st.dataframe(risk_df(slow.sort_values("Days of Cover", ascending=False)),
                     use_container_width=True)
    with r4:
        st.markdown(f"**{len(excess)} SKUs — cover > {planning_days*2} days**")
        st.dataframe(risk_df(excess.sort_values("Days of Cover", ascending=False)),
                     use_container_width=True)

# ── TRENDS ─────────────────────────────────────────────────
with tabs[6]:
    st.subheader("📈 Sales Trends")
    tc1, tc2 = st.columns(2)
    with tc1:
        st.markdown("**Weekly Sales**")
        if not weekly_trend.empty:
            st.line_chart(weekly_trend.set_index("Week")["Units Sold"])
    with tc2:
        st.markdown("**Monthly Sales**")
        if not monthly_trend.empty:
            st.bar_chart(monthly_trend.set_index("Month")["Units Sold"])

    st.markdown("**🔥 Top 20 SKUs by Avg Daily Sale**")
    t20 = top20.copy(); t20["Title"] = t20["Title"].str[:60]
    st.dataframe(serial(t20[["MSKU","Title","Avg Daily Sale","Days of Cover",
                               "Sellable Stock","Velocity","Top States"]]),
                 use_container_width=True)

    st.divider()
    st.markdown("**🔍 SKU Drill-Down: Daily Sales**")
    sku_list = sorted(sales["MSKU"].unique())
    sel_sku = st.selectbox("Select MSKU", sku_list)
    sku_daily = (sales[sales["MSKU"] == sel_sku]
                 .groupby("Date")["Qty"].sum().reset_index().set_index("Date"))
    if not sku_daily.empty:
        st.caption(f"**{title_map.get(sel_sku, sel_sku)[:80]}**")
        st.line_chart(sku_daily["Qty"])

# ── STATE DEMAND ───────────────────────────────────────────
with tabs[7]:
    st.subheader("🌍 State-Level Demand Analysis")
    sc1, sc2 = st.columns([2, 1])
    with sc1:
        st.dataframe(state_demand, use_container_width=True, hide_index=True)
    with sc2:
        st.markdown("**State × SKU Matrix (Top 10 SKUs)**")
        top10 = plan.nlargest(10, "Avg Daily Sale")["MSKU"].tolist()
        sm = (hist[hist["MSKU"].isin(top10)]
              .groupby(["Ship To State","MSKU"])["Qty"].sum()
              .unstack(fill_value=0))
        if not sm.empty:
            st.dataframe(sm, use_container_width=True)

# ── DAMAGED STOCK ──────────────────────────────────────────
with tabs[8]:
    st.subheader("🔧 Damaged / Non-Sellable Stock")
    st.caption("Raise a Removal Order or Reimbursement Claim in Seller Central for these units.")
    dr = damaged_report.copy()
    dr["Title"] = dr["Title"].str[:65]
    st.dataframe(serial(dr.sort_values("Qty", ascending=False)), use_container_width=True)
    st.divider()
    st.markdown("**Full Disposition Breakdown per MSKU**")
    db = disp_breakdown.copy()
    db.insert(1, "Title", db["MSKU"].map(title_map).fillna("").str[:60])
    st.dataframe(serial(db), use_container_width=True)

# ── AMAZON FLAT FILE ────────────────────────────────────────
with tabs[9]:
    st.subheader("📄 Amazon FBA Shipment Flat File Generator")
    st.info("""
**Steps:**
1. Edit quantities in the table below if needed
2. Choose target FC
3. Download the `.txt` flat file
4. In Seller Central → **Send to Amazon → Create New Shipment → Upload a file**
5. Upload and Amazon creates shipments automatically ✅

*Each FC requires a separate shipment. Use the per-FC download buttons below.*
    """)

    ff1, ff2, ff3 = st.columns(3)
    with ff1:
        shipment_name = st.text_input("Shipment Name",
            f"Replenishment_{datetime.now().strftime('%Y%m%d')}")
        ship_date = st.date_input("Expected Ship Date",
            datetime.now().date() + timedelta(days=2))
    with ff2:
        all_fcs = sorted(inv["FC Code"].unique())
        target_fc = st.selectbox("Target FC", all_fcs,
            format_func=lambda x: f"{x} — {fc_name(x)}")
    with ff3:
        st.markdown(f"**FC:** {fc_name(target_fc)}")
        st.markdown(f"**Cluster:** {fc_cluster(target_fc)}")
        st.markdown(f"**State:** {fc_st(target_fc)}")

    flat_base = fba_plan[fba_plan["Dispatch Needed"] > 0].copy()
    flat_base["Units to Ship"]  = flat_base["Dispatch Needed"].astype(int)
    flat_base["Units Per Case"] = default_case_qty
    flat_base["No. of Cases"]   = np.ceil(flat_base["Units to Ship"] / default_case_qty).astype(int)
    flat_base["Title Short"]    = flat_base["Title"].str[:80]

    st.markdown(f"**{len(flat_base)} SKUs | Total: {fmt(flat_base['Units to Ship'].sum())} units to dispatch**")
    edit_cols = ["MSKU","FNSKU","Title Short","Units to Ship","Units Per Case","No. of Cases"]
    edited = st.data_editor(
        flat_base[[c for c in edit_cols if c in flat_base.columns]].reset_index(drop=True),
        num_rows="dynamic",
        use_container_width=True)

    def build_flat_file(df, fc_code):
        parts   = [x.strip() for x in ship_from_addr.split(",")]
        city    = parts[1] if len(parts) > 1 else ""
        st_zip  = parts[2] if len(parts) > 2 else ""
        st_code = (st_zip.split("-")[0].strip()[:2].upper()
                   if "-" in st_zip else st_zip[:2].upper())
        postal  = st_zip.split("-")[1].strip() if "-" in st_zip else ""

        lines = ["TemplateType=FlatFileShipmentCreation\tVersion=2015.0403", ""]
        meta_h = ["ShipmentName","ShipFromName","ShipFromAddressLine1",
                  "ShipFromCity","ShipFromStateOrProvinceCode","ShipFromPostalCode",
                  "ShipFromCountryCode","ShipmentStatus","LabelPrepType",
                  "AreCasesRequired","DestinationFulfillmentCenterId"]
        meta_v = [shipment_name, ship_from_name or "My Warehouse",
                  parts[0] if parts else "",
                  city, st_code, postal, "IN",
                  "WORKING", label_owner,
                  "YES" if case_packed else "NO",
                  fc_code]
        lines += ["\t".join(meta_h), "\t".join(str(v) for v in meta_v), ""]

        item_h = ["SellerSKU","FNSKU","QuantityShipped","QuantityInCase",
                  "PrepOwner","LabelOwner","ItemDescription","ExpectedDeliveryDate"]
        lines.append("\t".join(item_h))
        for _, row in df.iterrows():
            qic = str(int(row.get("Units Per Case", default_case_qty))) if case_packed else ""
            lines.append("\t".join([
                str(row["MSKU"]),
                str(row.get("FNSKU", "")),
                str(int(row["Units to Ship"])),
                qic, prep_owner, label_owner,
                str(row.get("Title Short", ""))[:200],
                str(ship_date),
            ]))
        return "\n".join(lines)

    flat_content = build_flat_file(edited, target_fc)

    with st.expander("👁️ Preview Flat File"):
        st.code(flat_content[:2500] + ("\n...(truncated)" if len(flat_content) > 2500 else ""))

    st.download_button(
        f"📥 Download Flat File → {target_fc} ({fc_name(target_fc)})",
        flat_content.encode("utf-8"),
        f"Amazon_FBA_{target_fc}_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
        "text/plain")

    if len(all_fcs) > 1:
        st.divider()
        st.markdown("**📦 Download Separate Flat File per FC**")
        fc_btn_cols = st.columns(min(len(all_fcs), 4))
        for i, fc_code in enumerate(all_fcs):
            fc_flat = build_flat_file(edited, fc_code)
            with fc_btn_cols[i % 4]:
                st.download_button(
                    f"📥 {fc_code}\n{fc_name(fc_code)}",
                    fc_flat.encode("utf-8"),
                    f"Amazon_FBA_{fc_code}_{datetime.now().strftime('%Y%m%d')}.txt",
                    "text/plain",
                    key=f"fc_dl_{fc_code}")

# ═══════════════════════════════════════════════════════════
# EXCEL EXPORT
# ═══════════════════════════════════════════════════════════
st.markdown("---")

def build_excel():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb  = writer.book
        hdr = wb.add_format({"bold":True,"bg_color":"#1a365d","font_color":"white",
                              "border":1,"align":"center","valign":"vcenter","text_wrap":True})

        def ws(df, name):
            if df.empty:
                pd.DataFrame({"Note":["No data"]}).to_excel(writer, sheet_name=name, index=False)
                return
            d = df.copy()
            if "Title" in d.columns:
                d["Title"] = d["Title"].astype(str).str[:80]
            d.to_excel(writer, sheet_name=name, index=False, startrow=1)
            sheet = writer.sheets[name]
            for ci, cn in enumerate(d.columns):
                sheet.write(0, ci, cn, hdr)
                w = max(len(str(cn)),
                        d[cn].astype(str).str.len().max() if not d.empty else 10)
                sheet.set_column(ci, ci, min(w + 2, 45))
            sheet.freeze_panes(2, 0)
            sheet.autofilter(0, 0, len(d), len(d.columns) - 1)

        ws(serial(fba_plan.sort_values("Priority Score", ascending=False)),  "FBA Plan")
        ws(serial(fbm_plan),                                                  "FBM Plan")
        ws(serial(fc_plan.sort_values(["FC Cluster","FC Dispatch"],
                                       ascending=[True,False])),              "FC Dispatch Plan")
        ws(cluster_sum,                                                        "Cluster Summary")
        ws(serial(plan.sort_values("Priority Score", ascending=False)),       "All SKUs")
        ws(serial(critical.sort_values("Days of Cover")),                     "CRITICAL Stock")
        ws(serial(dead),                                                       "Dead Stock")
        ws(serial(slow.sort_values("Days of Cover", ascending=False)),        "Slow Moving")
        ws(serial(excess.sort_values("Days of Cover", ascending=False)),      "Excess Stock")
        ws(serial(damaged_report.sort_values("Qty", ascending=False)),        "Damaged Stock")
        ws(disp_breakdown,                                                     "Disposition Breakdown")
        ws(state_demand,                                                       "State Demand")
        ws(weekly_trend,                                                       "Weekly Trend")
        ws(serial(top20),                                                      "Top 20 SKUs")
    out.seek(0)
    return out

dl1, dl2 = st.columns(2)
with dl1:
    st.download_button(
        "📥 Download Full Intelligence Report (Excel)",
        build_excel(),
        f"FBA_Supply_Plan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with dl2:
    dispatch_csv = fba_plan[fba_plan["Dispatch Needed"] > 0][[
        "MSKU","FNSKU","Title","Avg Daily Sale","Sellable Stock","Days of Cover",
        "Required Stock","Dispatch Needed","Priority Score","Velocity","Top States"
    ]].to_csv(index=False)
    st.download_button(
        "📋 Download Dispatch Plan (CSV)",
        dispatch_csv,
        f"Dispatch_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
        "text/csv")
