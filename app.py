"""
FBA Smart Supply Planner  ·  Amazon India Edition
==================================================
Upload:
  1. Amazon Inventory Ledger  (CSV / ZIP / XLSX)
  2. MTR Sales Report         (CSV / ZIP / XLSX — multiple files OK)
"""

import math, io, zipfile
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(page_title="FBA Smart Supply Planner", layout="wide", page_icon="📦")
st.markdown("""
<style>
[data-testid="stMetricValue"]{font-size:1.55rem;font-weight:700}
.ab{padding:.6rem 1rem;border-radius:6px;margin:.3rem 0;font-weight:600}
.ab-r{background:#fff5f5;border-left:4px solid #e53e3e;color:#742a2a}
.ab-y{background:#fffff0;border-left:4px solid #d69e2e;color:#744210}
.ab-g{background:#f0fff4;border-left:4px solid #38a169;color:#1c4532}
.stTabs [data-baseweb="tab"]{font-size:13px;font-weight:600}
</style>""", unsafe_allow_html=True)

st.title("📦 FBA Smart Supply Planner — Amazon India")
st.caption("Inventory Ledger + MTR Sales → Dispatch plan · FC-wise allocation · Risk alerts · Amazon flat file")

# ─────────────────────────────────────────────────────────────────
# COMPLETE AMAZON INDIA FC MASTER  (75 codes — verified)
# Tuple: (FC Full Name, City, Area/Locality, State, Cluster)
#
# KEY NOTES:
#   PNQ2        → New Delhi (Mohan Co-op A-33), NOT Pune
#   PNQ1/PNQ3   → Pune Hinjewadi / Chakan
#   DEL8_DED5   → single FC in Sohna, Haryana
#   HYD8_HYD3   → single FC in Shamshabad, Telangana
# ─────────────────────────────────────────────────────────────────
FC_DATA = {
    # ASSAM
    "SGAA":      ("Guwahati FC (SGAA)",                     "Guwahati",        "Omshree Ind Park",           "ASSAM",          "East — Assam"),
    # DELHI
    "DEX3":      ("New Delhi FC A28 (DEX3)",                "New Delhi",       "Mohan Co-op, Mathura Rd",    "DELHI",          "Delhi NCR"),
    "DEX8":      ("New Delhi FC A29 (DEX8)",                "New Delhi",       "Mohan Co-op, Mathura Rd",    "DELHI",          "Delhi NCR"),
    "PNQ2":      ("New Delhi FC A33 (PNQ2)",                "New Delhi",       "Mohan Co-op Industrial",     "DELHI",          "Delhi NCR"),
    "XDEL":      ("Delhi XL FC (XDEL)",                     "New Delhi",       "Delhi",                      "DELHI",          "Delhi NCR"),
    # HARYANA
    "DEL2":      ("Tauru FC (DEL2)",                        "Mewat",           "Village Tauru",              "HARYANA",        "Delhi NCR"),
    "DEL4":      ("Gurgaon FC (DEL4)",                      "Gurgaon",         "Village Jamalpur",           "HARYANA",        "Delhi NCR"),
    "DEL5":      ("Manesar FC (DEL5)",                      "Manesar",         "Binola, NH-8",               "HARYANA",        "Delhi NCR"),
    "DEL6":      ("Bilaspur FC (DEL6)",                     "Bilaspur",        "Bilaspur",                   "HARYANA",        "Delhi NCR"),
    "DEL7":      ("Bilaspur FC 2 (DEL7)",                   "Bilaspur",        "Bilaspur",                   "HARYANA",        "Delhi NCR"),
    "DEL8":      ("Sohna FC (DEL8)",                        "Sohna",           "ESR Sohna Logistics Park",   "HARYANA",        "Delhi NCR"),
    "DEL8_DED5": ("Sohna ESR FC (DEL8/DED5)",              "Sohna",           "ESR Sohna-Ballabgarh Rd",    "HARYANA",        "Delhi NCR"),
    "DED3":      ("Farrukhnagar FC (DED3)",                 "Farrukhnagar",    "Gurgaon-122506",             "HARYANA",        "Delhi NCR"),
    "DED4":      ("Sohna FC 2 (DED4)",                      "Sohna",           "ESR Sohna Logistics Park",   "HARYANA",        "Delhi NCR"),
    "DED5":      ("Sohna FC 3 (DED5)",                      "Sohna",           "ESR Sohna Logistics Park",   "HARYANA",        "Delhi NCR"),
    # GUJARAT
    "AMD1":      ("Ahmedabad Naroda FC (AMD1)",             "Naroda",          "Naroda, Ahmedabad",          "GUJARAT",        "Gujarat West"),
    "AMD2":      ("Changodar FC (AMD2)",                    "Changodar",       "Gallops Ind Park",           "GUJARAT",        "Gujarat West"),
    "SUB1":      ("Surat FC (SUB1)",                        "Surat",           "Surat",                      "GUJARAT",        "Gujarat West"),
    # KARNATAKA
    "BLR4":      ("Devanahalli FC (BLR4)",                  "Devanahalli",     "Hitech Aerospace Park",      "KARNATAKA",      "Bangalore"),
    "BLR5":      ("Bommasandra FC (BLR5)",                  "Bommasandra",     "Hosakote, Bengaluru",        "KARNATAKA",      "Bangalore"),
    "BLR6":      ("Nelamangala FC (BLR6)",                  "Nelamangala",     "Bengaluru",                  "KARNATAKA",      "Bangalore"),
    "BLR7":      ("Hoskote FC (BLR7)",                      "Hoskote",         "Anekal Taluk, Bengaluru",    "KARNATAKA",      "Bangalore"),
    "BLR8":      ("Devanahalli FC 2 (BLR8)",                "Devanahalli",     "Hitech Aerospace Park",      "KARNATAKA",      "Bangalore"),
    "BLR10":     ("Kudlu Gate FC (BLR10)",                  "Kudlu Gate",      "Bengaluru",                  "KARNATAKA",      "Bangalore"),
    "BLR12":     ("Attibele FC (BLR12)",                    "Attibele",        "Bengaluru",                  "KARNATAKA",      "Bangalore"),
    "BLR13":     ("Jigani FC (BLR13)",                      "Jigani",          "Bengaluru",                  "KARNATAKA",      "Bangalore"),
    "BLR14":     ("Anekal FC (BLR14)",                      "Anekal",          "Bengaluru",                  "KARNATAKA",      "Bangalore"),
    "SCJA":      ("Bangalore SCJA FC (SCJA)",               "Bangalore",       "Bengaluru",                  "KARNATAKA",      "Bangalore"),
    "XSAB":      ("Bangalore XS FC (XSAB)",                 "Bangalore",       "Bengaluru",                  "KARNATAKA",      "Bangalore"),
    "SBLA":      ("Bangalore SBLA FC (SBLA)",               "Bangalore",       "Bengaluru",                  "KARNATAKA",      "Bangalore"),
    # MADHYA PRADESH
    "SIDA":      ("Indore FC (SIDA)",                       "Indore",          "Village Pipliya Kumhar",     "MADHYA PRADESH", "Central"),
    "FBHB":      ("Bhopal FC (FBHB)",                       "Bhopal",          "Govindpura Ind Area",        "MADHYA PRADESH", "Central"),
    "FIDA":      ("Bhopal FC 2 (FIDA)",                     "Bhopal",          "Govindpura Ind Area",        "MADHYA PRADESH", "Central"),
    "IND1":      ("Indore FC 2 (IND1)",                     "Indore",          "Indore",                     "MADHYA PRADESH", "Central"),
    # MAHARASHTRA
    "BOM1":      ("Bhiwandi FC (BOM1)",                     "Bhiwandi",        "Bhiwandi, Thane",            "MAHARASHTRA",    "Mumbai West"),
    "BOM3":      ("Nashik FC (BOM3)",                       "Nashik",          "Nashik",                     "MAHARASHTRA",    "Mumbai West"),
    "BOM4":      ("Vashere FC (BOM4)",                      "Bhiwandi",        "Village Vashere",            "MAHARASHTRA",    "Mumbai West"),
    "BOM5":      ("Bhiwandi 2 FC (BOM5)",                   "Bhiwandi",        "Village Vashere",            "MAHARASHTRA",    "Mumbai West"),
    "BOM7":      ("Bhiwandi 3 FC (BOM7)",                   "Bhiwandi",        "Village Vahuli",             "MAHARASHTRA",    "Mumbai West"),
    "ISK3":      ("Bhiwandi ISK3 FC (ISK3)",                "Bhiwandi",        "Village Pise",               "MAHARASHTRA",    "Mumbai West"),
    "PNQ1":      ("Pune Hinjewadi FC (PNQ1)",               "Pune",            "Hinjewadi",                  "MAHARASHTRA",    "Mumbai West"),
    "PNQ3":      ("Pune Chakan FC (PNQ3)",                  "Pune",            "Village Ambethan, Khed",     "MAHARASHTRA",    "Mumbai West"),
    "SAMB":      ("Mumbai SAMB FC (SAMB)",                  "Mumbai",          "Mumbai",                     "MAHARASHTRA",    "Mumbai West"),
    "XBOM":      ("Mumbai XL FC (XBOM)",                    "Bhiwandi",        "Bhiwandi",                   "MAHARASHTRA",    "Mumbai West"),
    # PUNJAB
    "ATX1":      ("Ludhiana FC (ATX1)",                     "Ludhiana",        "Near Katana Sahib Gurdwara", "PUNJAB",         "North Punjab"),
    "LDH1":      ("Ludhiana FC 2 (LDH1)",                   "Ludhiana",        "Ludhiana",                   "PUNJAB",         "North Punjab"),
    "RAJ1":      ("Rajpura FC (RAJ1)",                      "Rajpura",         "Rajpura",                    "PUNJAB",         "North Punjab"),
    # RAJASTHAN
    "JAI1":      ("Jaipur FC (JAI1)",                       "Jaipur",          "Bagru",                      "RAJASTHAN",      "North Rajasthan"),
    "JPX1":      ("Jaipur JPX1 FC (JPX1)",                  "Jaipur",          "Jhotwara Ind Area",          "RAJASTHAN",      "North Rajasthan"),
    "JPX2":      ("Bagru FC (JPX2)",                        "Bagru",           "Bagru, Sanganer",            "RAJASTHAN",      "North Rajasthan"),
    # TAMIL NADU
    "MAA1":      ("Irungattukottai FC (MAA1)",              "Irungattukottai", "Sriperumbudur Tk",           "TAMIL NADU",     "Chennai"),
    "MAA2":      ("Ponneri FC (MAA2)",                      "Ponneri",         "Thiruvallur Dist",           "TAMIL NADU",     "Chennai"),
    "MAA3":      ("Sriperumbudur FC (MAA3)",                "Sriperumbudur",   "Chennai",                    "TAMIL NADU",     "Chennai"),
    "MAA4":      ("Ambattur FC (MAA4)",                     "Ambattur",        "Thiruvallur Dist",           "TAMIL NADU",     "Chennai"),
    "MAA5":      ("Kanchipuram FC (MAA5)",                  "Kanchipuram",     "Kanchipuram Dist",           "TAMIL NADU",     "Chennai"),
    "CJB1":      ("Coimbatore FC (CJB1)",                   "Coimbatore",      "Palladam Main Rd",           "TAMIL NADU",     "Chennai"),
    "SMAB":      ("Chennai SMAB FC (SMAB)",                 "Chennai",         "Chennai",                    "TAMIL NADU",     "Chennai"),
    # KERALA
    "COK1":      ("Kochi FC (COK1)",                        "Kochi",           "Kochi",                      "KERALA",         "South Kerala"),
    # TELANGANA
    "HYD3":      ("Shamshabad FC (HYD3)",                   "Shamshabad",      "Mamidipally Village",        "TELANGANA",      "Hyderabad"),
    "HYD6":      ("Kothur FC (HYD6)",                       "Kothur",          "Mahbubnagar Dist",           "TELANGANA",      "Hyderabad"),
    "HYD7":      ("Medchal FC (HYD7)",                      "Medchal",         "Hyderabad",                  "TELANGANA",      "Hyderabad"),
    "HYD8":      ("Shamshabad FC 2 (HYD8)",                 "Shamshabad",      "Hyderabad",                  "TELANGANA",      "Hyderabad"),
    "HYD8_HYD3": ("Shamshabad Mamidipally FC (HYD8/HYD3)", "Shamshabad",      "Mamidipally Village",        "TELANGANA",      "Hyderabad"),
    "HYD9":      ("Pedda Amberpet FC (HYD9)",               "Pedda Amberpet",  "Hyderabad",                  "TELANGANA",      "Hyderabad"),
    "HYD10":     ("Ghatkesar FC (HYD10)",                   "Ghatkesar",       "Hyderabad",                  "TELANGANA",      "Hyderabad"),
    "HYD11":     ("Kompally FC (HYD11)",                    "Kompally",        "Hyderabad",                  "TELANGANA",      "Hyderabad"),
    # UTTAR PRADESH
    "LKO1":      ("Lucknow FC (LKO1)",                      "Lucknow",         "Village Bhukapur",           "UTTAR PRADESH",  "North UP"),
    "AGR1":      ("Agra FC (AGR1)",                         "Agra",            "Agra",                       "UTTAR PRADESH",  "North UP"),
    "SLDK":      ("Kishanpur FC (SLDK)",                    "Bijnour",         "Kishanpur Kodia",            "UTTAR PRADESH",  "North UP"),
    # WEST BENGAL
    "CCU1":      ("Kolkata FC (CCU1)",                      "Kolkata",         "Rajarhat",                   "WEST BENGAL",    "Kolkata East"),
    "CCU2":      ("Kolkata FC 2 (CCU2)",                    "Kolkata",         "Kolkata",                    "WEST BENGAL",    "Kolkata East"),
    "CCX1":      ("Howrah FC (CCX1)",                       "Howrah",          "Panchla, Raghudevpur",       "WEST BENGAL",    "Kolkata East"),
    "CCX2":      ("Howrah FC 2 (CCX2)",                     "Howrah",          "Panchla",                    "WEST BENGAL",    "Kolkata East"),
    # BIHAR
    "PAT1":      ("Patna FC (PAT1)",                        "Patna",           "Patna",                      "BIHAR",          "Kolkata East"),
    # ODISHA
    "BBS1":      ("Bhubaneswar FC (BBS1)",                  "Bhubaneswar",     "Bhubaneswar",                "ODISHA",         "Kolkata East"),
}
# Tuple index: 0=Full Name  1=City  2=Area  3=State  4=Cluster


def _fc(code, idx, fallback):
    e = FC_DATA.get(str(code).upper().strip())
    return e[idx] if e else fallback

def fc_name(c):    return _fc(c, 0, f"{c} (Unknown FC)")
def fc_city(c):    return _fc(c, 1, "Unknown")
def fc_area(c):    return _fc(c, 2, "Unknown")
def fc_state(c):   return _fc(c, 3, "Unknown")
def fc_cluster(c): return _fc(c, 4, "Other")


# ─────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────
def read_file(up):
    try:
        n = up.name.lower()
        if n.endswith(".zip"):
            with zipfile.ZipFile(up) as z:
                for m in z.namelist():
                    if m.lower().endswith(".csv"):
                        return pd.read_csv(z.open(m), low_memory=False)
            return pd.DataFrame()
        if n.endswith((".xlsx", ".xls")):
            return pd.read_excel(up)
        return pd.read_csv(up, low_memory=False)
    except Exception as e:
        st.error(f"Cannot read **{up.name}**: {e}")
        return pd.DataFrame()


def find_col(df, exact, fuzzy=None):
    low = {c.strip().lower(): c for c in df.columns}
    for e in exact:
        if e.strip().lower() in low:
            return low[e.strip().lower()]
    if fuzzy:
        for col in df.columns:
            if any(s in col.lower() for s in fuzzy):
                return col
    return ""


def sno(df):
    df = df.copy()
    df.drop(columns=["S.No"], errors="ignore", inplace=True)
    df.insert(0, "S.No", range(1, len(df)+1))
    return df


def fmt(n):
    try: return f"{int(n):,}"
    except: return str(n)


def trunc(s, n=60):
    s = str(s)
    return s[:n]+"…" if len(s)>n else s


def health_tag(doc, pd_):
    if doc <= 0:       return "⚫ No Stock"
    if doc < 14:       return "🔴 Critical"
    if doc < 30:       return "🟠 Low"
    if doc < pd_:      return "🟢 Healthy"
    if doc < pd_*2:    return "🟡 Excess"
    return "🔵 Overstocked"


def vel_tag(avg):
    if avg <= 0:   return "⚫ Dead"
    if avg < 0.5:  return "🔵 Slow"
    if avg < 2:    return "🟡 Medium"
    if avg < 5:    return "🟢 Fast"
    return "🔥 Top Seller"


# ─────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Planning Controls")
    planning_days = st.number_input("Planning Days (Coverage Target)", 7, 180, 60, 1)
    svc_lvl       = st.selectbox("Service Level", ["90%","95%","98%"])
    z_val         = {"90%":1.28,"95%":1.65,"98%":2.05}[svc_lvl]
    sales_basis   = st.selectbox("Sales Basis for Planning",
                        ["Total Sales (full history)","Last 30 Days",
                         "Last 60 Days","Last 90 Days"])
    st.divider()
    st.subheader("🚚 Shipment Settings")
    ship_name   = st.text_input("Ship-From Warehouse Name", placeholder="My Warehouse")
    ship_addr   = st.text_area("Ship-From Address",
                               placeholder="123 Street, City, State - 400001")
    case_qty    = st.number_input("Default Units Per Carton", 1, 500, 12)
    case_packed = st.checkbox("Case-Packed Shipment", value=False)
    prep_own    = st.selectbox("Prep Ownership",  ["AMAZON","SELLER"])
    lbl_own     = st.selectbox("Label Ownership", ["AMAZON","SELLER"])
    st.divider()
    st.subheader("🎯 Display Filters")
    min_disp    = st.number_input("Min Dispatch Units to Show", 0, 99999, 1)
    all_clusters = sorted({v[4] for v in FC_DATA.values()})
    foc_cluster = st.multiselect("Focus Clusters", all_clusters)


# ─────────────────────────────────────────────────────────────────
# FILE UPLOADS
# ─────────────────────────────────────────────────────────────────
st.markdown("### 📁 Upload Files")
u1, u2 = st.columns(2)
with u1:
    mtr_files = st.file_uploader(
        "📊 MTR / Sales Report — multiple files OK (CSV, ZIP, XLSX)",
        type=["csv","zip","xlsx"], accept_multiple_files=True)
with u2:
    inv_file = st.file_uploader(
        "🏭 Amazon Inventory Ledger (CSV, ZIP, XLSX)",
        type=["csv","zip","xlsx"])

with st.expander("📋 Expected File Formats"):
    st.markdown("""
**Inventory Ledger** *(Seller Central → Reports → Fulfillment → Inventory Ledger)*

| Column | Example |
|--------|---------|
| `Date` | 02/12/2026 |
| `FNSKU` | X00275KJZP |
| `MSKU` | BR-899 |
| `Title` | GLOOYA Bracelet |
| `Disposition` | SELLABLE / CUSTOMER_DAMAGED / CARRIER_DAMAGED |
| `Ending Warehouse Balance` | 18 |
| `Location` | LKO1 |

**MTR Sales Report** *(Seller Central → Reports → Tax → MTR)*

| Column | Example |
|--------|---------|
| `Sku` | BR-899 |
| `Quantity` | 5 |
| `Shipment Date` | 2026-01-15 |
| `Ship To State` | UTTAR PRADESH |
| `Fulfilment` | AFN / MFN |
""")

if not (mtr_files and inv_file):
    st.info("👆 Upload both files above to start the planner.")
    with st.expander("🗺️ All 75 Verified Amazon India FC Codes"):
        st.dataframe(pd.DataFrame([{
            "FC Code":k, "FC Full Name":v[0], "City":v[1],
            "Area / Locality":v[2], "State":v[3], "Cluster":v[4]}
            for k,v in FC_DATA.items()]),
            use_container_width=True, hide_index=True)
    st.stop()


# ═════════════════════════════════════════════════════════════════
# ① LOAD & PARSE INVENTORY LEDGER
# ═════════════════════════════════════════════════════════════════
inv_raw = read_file(inv_file)
if inv_raw.empty:
    st.error("Inventory file could not be read — check format."); st.stop()
inv_raw.columns = inv_raw.columns.str.strip()

C_SKU   = find_col(inv_raw, ["MSKU","Sku","SKU","ASIN"],                       ["msku","sku","asin"])
C_TITLE = find_col(inv_raw, ["Title","Product Name","Description"],             ["title","product"])
C_QTY   = find_col(inv_raw, ["Ending Warehouse Balance","Quantity","Qty"],      ["ending","balance"])
C_LOC   = find_col(inv_raw, ["Location","Warehouse Code","FC Code","FC"],       ["location","warehouse code","fc code"])
C_DISP  = find_col(inv_raw, ["Disposition"],                                    ["disposition"])
C_FNSKU = find_col(inv_raw, ["FNSKU"],                                          ["fnsku"])
C_DATE  = find_col(inv_raw, ["Date"],                                           ["date"])

for label, col in [("SKU/MSKU",C_SKU),("Ending Balance",C_QTY),("Location/FC",C_LOC)]:
    if not col:
        st.error(f"Inventory file missing required column: **{label}**")
        st.write("Columns found:", list(inv_raw.columns)); st.stop()

inv = inv_raw.copy()
inv["MSKU"]        = inv[C_SKU].astype(str).str.strip()
inv["Stock"]       = pd.to_numeric(inv[C_QTY], errors="coerce").fillna(0)
inv["FC Code"]     = inv[C_LOC].astype(str).str.upper().str.strip()
inv["Title"]       = inv[C_TITLE].astype(str).str.strip() if C_TITLE else ""
inv["FNSKU"]       = inv[C_FNSKU].astype(str).str.strip() if C_FNSKU else ""
inv["Disposition"] = inv[C_DISP].astype(str).str.upper().str.strip() if C_DISP else "SELLABLE"

# ── CRITICAL FIX ──────────────────────────────────────────────────
# Amazon Inventory Ledger = ONE ROW PER DAY per SKU/FC/Disposition.
# Summing all rows multiplies stock by number of days in the report.
# Fix: sort by date → keep ONLY the LATEST row per group.
# ─────────────────────────────────────────────────────────────────
if C_DATE:
    inv["_dt"]   = pd.to_datetime(inv[C_DATE], dayfirst=True, errors="coerce")
    latest_dt    = inv["_dt"].max()
    inv          = (inv.sort_values("_dt")
                       .groupby(["MSKU","Disposition","FC Code"], as_index=False)
                       .last().drop(columns=["_dt"]))
    inv_date_lbl = latest_dt.strftime("%d %b %Y") if pd.notna(latest_dt) else "latest row"
else:
    inv          = inv.groupby(["MSKU","Disposition","FC Code"], as_index=False).last()
    inv_date_lbl = "latest row"

# Enrich with FC master
inv["FC Name"]    = inv["FC Code"].apply(fc_name)
inv["FC City"]    = inv["FC Code"].apply(fc_city)
inv["FC Area"]    = inv["FC Code"].apply(fc_area)
inv["FC State"]   = inv["FC Code"].apply(fc_state)
inv["FC Cluster"] = inv["FC Code"].apply(fc_cluster)

# Lookup maps
fnsku_map = (inv[inv["FNSKU"].notna() & ~inv["FNSKU"].isin(["nan",""])]
             .drop_duplicates("MSKU").set_index("MSKU")["FNSKU"].to_dict())
title_map = inv.drop_duplicates("MSKU").set_index("MSKU")["Title"].to_dict()

# Split sellable / damaged
inv_sell = inv[inv["Disposition"]=="SELLABLE"].copy()
inv_dmg  = inv[inv["Disposition"]!="SELLABLE"].copy()

# Stock aggregates
fc_stock = (inv_sell
    .groupby(["MSKU","FC Code","FC Name","FC City","FC Area","FC Cluster","FC State"])["Stock"]
    .sum().reset_index().rename(columns={"Stock":"FC Stock"}))

sku_sell_stock = (inv_sell.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock":"Sellable Stock"}))

sku_dmg_stock = (inv_dmg.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock":"Damaged Stock"}))

disp_bkdn = (inv.groupby(["MSKU","Disposition"])["Stock"]
    .sum().unstack(fill_value=0).reset_index())

dmg_report = (inv_dmg.groupby(["MSKU","Disposition"])["Stock"]
    .sum().reset_index().rename(columns={"Stock":"Qty"}))
dmg_report["Title"] = dmg_report["MSKU"].map(title_map).fillna("")


# ═════════════════════════════════════════════════════════════════
# ② LOAD & PARSE SALES / MTR
# ═════════════════════════════════════════════════════════════════
raw_sales = [read_file(f) for f in mtr_files]
raw_sales = [d for d in raw_sales if not d.empty]
if not raw_sales:
    st.error("Could not read any MTR / sales file."); st.stop()

sr = pd.concat(raw_sales, ignore_index=True)

CS_SKU  = find_col(sr, ["Sku","SKU","MSKU","Item SKU"],                 ["sku","msku"])
CS_QTY  = find_col(sr, ["Quantity","Qty","Units"],                      ["qty","unit","quant"])
CS_DATE = find_col(sr, ["Shipment Date","Purchase Date","Order Date","Date"], ["date"])
CS_ST   = find_col(sr, ["Ship To State","Shipping State"],               ["ship to","state"])
CS_FT   = find_col(sr, ["Fulfilment","Fulfillment",
                          "Fulfilment Channel","Fulfillment Channel"],   ["fulfil","channel"])

for label, col in [("SKU",CS_SKU),("Quantity",CS_QTY),("Date",CS_DATE)]:
    if not col:
        st.error(f"Sales file missing: **{label}**")
        st.write("Columns found:", list(sr.columns)); st.stop()

sales = sr.rename(columns={CS_SKU:"MSKU",CS_QTY:"Qty",CS_DATE:"Date"}).copy()
sales["MSKU"] = sales["MSKU"].astype(str).str.strip()
sales["Qty"]  = pd.to_numeric(sales["Qty"], errors="coerce").fillna(0)
sales["Date"] = pd.to_datetime(sales["Date"], dayfirst=True, errors="coerce")
sales = sales.dropna(subset=["Date"])
sales = sales[sales["Qty"]>0].copy()
sales["Ship To State"] = (sales[CS_ST].astype(str).str.upper().str.strip()
                          if CS_ST else "UNKNOWN")
FT_MAP = {"AFN":"FBA","FBA":"FBA","AMAZON_FULFILLED":"FBA",
          "MFN":"FBM","FBM":"FBM","MERCHANT_FULFILLED":"FBM"}
sales["Channel"] = (sales[CS_FT].astype(str).str.upper().str.strip().map(FT_MAP).fillna("FBA")
                    if CS_FT else "FBA")

s_min   = sales["Date"].min();  s_max = sales["Date"].max()
up_days = max((s_max-s_min).days+1, 1)
window  = (30 if "30" in sales_basis else 60 if "60" in sales_basis
           else 90 if "90" in sales_basis else up_days)
cutoff  = s_max - pd.Timedelta(days=window-1)
hist    = sales[sales["Date"]>=cutoff].copy() if window<up_days else sales.copy()
h_days  = max((hist["Date"].max()-hist["Date"].min()).days+1, 1)

ch_hist  = (hist.groupby(["MSKU","Channel"])["Qty"].sum()
            .reset_index().rename(columns={"Qty":"Hist Sales"}))
ch_full  = (sales.groupby(["MSKU","Channel"])["Qty"].sum()
            .reset_index().rename(columns={"Qty":"All-Time Sales"}))
ch_daily = hist.groupby(["MSKU","Channel","Date"])["Qty"].sum().reset_index()
ch_std   = (ch_daily.groupby(["MSKU","Channel"])["Qty"].std()
            .reset_index().rename(columns={"Qty":"Demand StdDev"}))
top_states = (hist.groupby(["MSKU","Ship To State"])["Qty"].sum()
    .reset_index().sort_values("Qty",ascending=False)
    .groupby("MSKU")["Ship To State"]
    .apply(lambda x:", ".join(list(x)[:3]))
    .reset_index().rename(columns={"Ship To State":"Top States"}))


# ═════════════════════════════════════════════════════════════════
# ③ PLANNING TABLE
# ═════════════════════════════════════════════════════════════════
keys = pd.concat([
    ch_hist[["MSKU","Channel"]],
    sku_sell_stock.assign(Channel="FBA")[["MSKU","Channel"]]
], ignore_index=True).drop_duplicates()

plan = keys.copy()
for df_, on_ in [(ch_hist,["MSKU","Channel"]),(ch_full,["MSKU","Channel"]),
                 (ch_std,["MSKU","Channel"]),(sku_sell_stock,"MSKU"),
                 (sku_dmg_stock,"MSKU"),(top_states,"MSKU")]:
    plan = plan.merge(df_, on=on_, how="left")

for c in ["Hist Sales","All-Time Sales","Sellable Stock","Damaged Stock","Demand StdDev"]:
    plan[c] = pd.to_numeric(plan[c], errors="coerce").fillna(0)

plan["Title"]           = plan["MSKU"].map(title_map).fillna("")
plan["FNSKU"]           = plan["MSKU"].map(fnsku_map).fillna("")
plan["Sales Days Used"] = h_days
plan["Planning Days"]   = planning_days
plan["Avg Daily Sale"]  = (plan["Hist Sales"]/h_days).round(4)
plan["Safety Stock"]    = (z_val*plan["Demand StdDev"]*math.sqrt(planning_days)).round(0)
plan["Base Req"]        = (plan["Avg Daily Sale"]*planning_days).round(0)
plan["Required Stock"]  = (plan["Base Req"]+plan["Safety Stock"]).round(0)
plan["Dispatch Needed"] = (plan["Required Stock"]-plan["Sellable Stock"]).clip(lower=0).round(0)
plan["Days of Cover"]   = np.where(
    plan["Avg Daily Sale"]>0,
    (plan["Sellable Stock"]/plan["Avg Daily Sale"]).round(1),
    np.where(plan["Sellable Stock"]>0, 9999, 0))
plan["Health"]   = plan.apply(lambda r: health_tag(r["Days of Cover"],planning_days), axis=1)
plan["Velocity"] = plan["Avg Daily Sale"].apply(vel_tag)

_mx = plan["Avg Daily Sale"].max() or 1
plan["Priority"] = plan.apply(lambda r: round(
    (max(0,(planning_days-min(r["Days of Cover"],planning_days))/planning_days)*0.65
     +min(r["Avg Daily Sale"]/_mx,1)*0.35)*100,1)
    if r["Avg Daily Sale"]>0 else 0, axis=1)

fba_plan = plan[plan["Channel"]=="FBA"].copy()
fbm_plan = plan[plan["Channel"]=="FBM"].copy()

PCOLS = ["MSKU","FNSKU","Title","Avg Daily Sale","Sellable Stock","Safety Stock",
         "Required Stock","Dispatch Needed","Days of Cover","Health","Velocity",
         "Priority","Top States"]


# ═════════════════════════════════════════════════════════════════
# ④ FC-WISE PLAN  — Code · Full Name · City · Area · State · Cluster
# ═════════════════════════════════════════════════════════════════
fcp = fc_stock.merge(
    plan[["MSKU","Channel","Avg Daily Sale","Dispatch Needed","Sellable Stock",
          "Required Stock","Hist Sales","Title","FNSKU","Priority","Days of Cover"]],
    on="MSKU", how="left")

fcp["FC DOC"] = np.where(
    fcp["Avg Daily Sale"]>0,
    (fcp["FC Stock"]/fcp["Avg Daily Sale"]).round(1),
    np.where(fcp["FC Stock"]>0, 9999, 0))
fcp["FC Health"]   = fcp.apply(lambda r: health_tag(r["FC DOC"],planning_days), axis=1)
fcp["FC Priority"] = fcp["FC DOC"].apply(
    lambda d:"🔴 Urgent" if d<14 else("🟠 Soon" if d<30 else "🟢 OK"))

_tot = (fcp.groupby(["MSKU","Channel"])["FC Stock"].sum()
        .reset_index().rename(columns={"FC Stock":"_T"}))
fcp = fcp.merge(_tot, on=["MSKU","Channel"], how="left")
fcp["FC Share"]    = (fcp["FC Stock"]/fcp["_T"].replace(0,1)).fillna(1.0)
fcp["FC Dispatch"] = (fcp["Dispatch Needed"]*fcp["FC Share"]).round(0)
fcp.drop(columns=["_T"], inplace=True)

fcv = fcp.copy()
if foc_cluster: fcv = fcv[fcv["FC Cluster"].isin(foc_cluster)]
fcv = fcv[fcv["FC Dispatch"]>=min_disp]

FC_COLS = ["FC Code","FC Name","FC City","FC Area","FC State","FC Cluster",
           "MSKU","Title","FC Stock","Avg Daily Sale",
           "FC DOC","FC Health","FC Dispatch","FC Priority"]


# ═════════════════════════════════════════════════════════════════
# ⑤ CLUSTER SUMMARY
# ═════════════════════════════════════════════════════════════════
cl_sum = (fcp.groupby("FC Cluster").agg(
    FC_Names       =("FC Name",    lambda x:" | ".join(sorted(set(x)))),
    FC_Codes       =("FC Code",    lambda x:", ".join(sorted(set(x)))),
    FC_Cities      =("FC City",    lambda x:", ".join(sorted(set(x)))),
    Total_Stock    =("FC Stock",   "sum"),
    Dispatch_Needed=("FC Dispatch","sum"),
    Avg_DOC        =("FC DOC",     "mean"),
    SKUs           =("MSKU",       "nunique"),
).reset_index().rename(columns={
    "FC_Names":"FC Names","FC_Codes":"FC Codes","FC_Cities":"Cities",
    "Total_Stock":"Total Stock","Dispatch_Needed":"Dispatch Needed",
    "Avg_DOC":"Avg Days of Cover","SKUs":"Unique SKUs"}))
cl_sum["Avg Days of Cover"] = cl_sum["Avg Days of Cover"].round(1)
cl_sum = sno(cl_sum.sort_values("Dispatch Needed",ascending=False))


# ═════════════════════════════════════════════════════════════════
# ⑥ RISK TABLES
# ═════════════════════════════════════════════════════════════════
r_crit   = plan[(plan["Days of Cover"]<14)  &(plan["Avg Daily Sale"]>0)].copy()
r_dead   = plan[(plan["Avg Daily Sale"]==0) &(plan["Sellable Stock"]>0)].copy()
r_slow   = plan[(plan["Avg Daily Sale"]>0)  &(plan["Days of Cover"]>90)].copy()
r_excess = plan[plan["Days of Cover"]>planning_days*2].copy()
r_top20  = plan[plan["Avg Daily Sale"]>0].nlargest(20,"Avg Daily Sale").copy()


# ═════════════════════════════════════════════════════════════════
# ⑦ STATE / TREND
# ═════════════════════════════════════════════════════════════════
st_dem = (sales.groupby("Ship To State")["Qty"].sum()
    .reset_index().rename(columns={"Qty":"Total Units"})
    .sort_values("Total Units",ascending=False))
st_dem["% Share"] = (st_dem["Total Units"]/st_dem["Total Units"].sum()*100).round(1)
st_dem = sno(st_dem)

wk_tr = (sales.assign(Week=sales["Date"].dt.to_period("W").astype(str))
    .groupby("Week")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units Sold"}).sort_values("Week"))
mo_tr = (sales.assign(Month=sales["Date"].dt.to_period("M").astype(str))
    .groupby("Month")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units Sold"}).sort_values("Month"))


# ═════════════════════════════════════════════════════════════════
# ⑧ EXECUTIVE DASHBOARD
# ═════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("### 📊 Executive Dashboard")

tot_skus = plan["MSKU"].nunique()
tot_sell = int(sku_sell_stock["Sellable Stock"].sum())
tot_dmg  = int(sku_dmg_stock["Damaged Stock"].sum()) if not sku_dmg_stock.empty else 0
tot_disp = int(fba_plan["Dispatch Needed"].sum())
avg_doc  = plan[plan["Avg Daily Sale"]>0]["Days of Cover"].replace(9999,np.nan).mean()

m1,m2,m3,m4,m5,m6 = st.columns(6)
m1.metric("Total SKUs",        fmt(tot_skus))
m2.metric("Sellable Stock",    fmt(tot_sell))
m3.metric("Damaged Stock",     fmt(tot_dmg))
m4.metric("Units to Dispatch", fmt(tot_disp))
m5.metric("🔴 Critical SKUs",  fmt(len(r_crit)))
m6.metric("Avg Days of Cover", f"{avg_doc:.0f}d" if pd.notna(avg_doc) else "N/A")

st.markdown(
    f'<div class="ab ab-g">✅ Inventory snapshot: <b>{inv_date_lbl}</b> | '
    f'{len(inv)} rows after daily dedup (raw: {len(inv_raw)} rows)</div>',
    unsafe_allow_html=True)
if r_crit.shape[0]:
    st.markdown(f'<div class="ab ab-r">🚨 {len(r_crit)} SKU(s) below 14 days — replenish immediately!</div>',unsafe_allow_html=True)
if r_dead.shape[0]:
    st.markdown(f'<div class="ab ab-y">⚠️ {len(r_dead)} SKU(s) have stock but ZERO sales — fix listing or liquidate.</div>',unsafe_allow_html=True)
if r_excess.shape[0]:
    st.markdown(f'<div class="ab ab-y">📦 {len(r_excess)} SKU(s) overstocked (>{planning_days*2}d) — pause replenishment.</div>',unsafe_allow_html=True)
if tot_dmg:
    st.markdown(f'<div class="ab ab-y">🔧 {fmt(tot_dmg)} damaged units — raise Removal / Reimbursement in Seller Central.</div>',unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════
# ⑨ MAIN TABS
# ═════════════════════════════════════════════════════════════════
tabs = st.tabs([
    "🏠 Overview","📦 FBA Plan","📮 FBM Plan","🏭 FC-Wise Plan",
    "🗺️ Cluster View","🚨 Risk Alerts","📈 Trends",
    "🌍 State Demand","🔧 Damaged Stock","📄 Amazon Flat File"])

# ── OVERVIEW ─────────────────────────────────────────────────────
with tabs[0]:
    st.subheader("🏠 Planning Overview")
    c1,c2 = st.columns(2)
    with c1:
        fcs_found = ", ".join(sorted(inv["FC Code"].unique()))
        st.info(
            f"**Sales data:** {s_min.strftime('%d %b %Y')} → {s_max.strftime('%d %b %Y')} ({up_days} days)\n\n"
            f"**Sales basis:** {sales_basis} ({h_days} days used)\n\n"
            f"**Planning horizon:** {planning_days} days | **Service level:** {svc_lvl} (Z={z_val})\n\n"
            f"**Total SKUs:** {tot_skus} | **FCs in data:** {fcs_found}")
    with c2:
        st.markdown("**🔝 Top 5 Most Urgent SKUs**")
        top5 = (plan[plan["Dispatch Needed"]>0].nlargest(5,"Priority")
                [["MSKU","Title","Avg Daily Sale","Days of Cover",
                  "Dispatch Needed","Priority"]].copy())
        top5["Title"] = top5["Title"].apply(lambda x:trunc(x,45))
        st.dataframe(top5, use_container_width=True, hide_index=True)

    st.markdown("**📊 Monthly Sales Trend**")
    if not mo_tr.empty: st.bar_chart(mo_tr.set_index("Month")["Units Sold"])

    oa,ob = st.columns(2)
    with oa:
        st.markdown("**🗺️ Cluster Dispatch Summary**")
        st.dataframe(cl_sum[["FC Cluster","FC Names","Cities","Total Stock",
                              "Dispatch Needed","Avg Days of Cover","Unique SKUs"]],
                     use_container_width=True, hide_index=True)
    with ob:
        st.markdown("**🏭 Stock by FC (Code + Name + City + State)**")
        fc_inv = (fcp.groupby(["FC Code","FC Name","FC City","FC State","FC Cluster"])["FC Stock"]
                  .sum().reset_index()
                  .rename(columns={"FC Stock":"Total Sellable Stock"})
                  .sort_values("Total Sellable Stock",ascending=False))
        st.dataframe(fc_inv, use_container_width=True, hide_index=True)

# ── FBA PLAN ─────────────────────────────────────────────────────
with tabs[1]:
    st.subheader("📦 FBA Planning")
    st.caption(f"{len(fba_plan)} SKUs | {planning_days}-day plan @ {svc_lvl}")
    fv = (fba_plan[fba_plan["Dispatch Needed"]>=min_disp]
          .sort_values("Priority",ascending=False).copy())
    fv["Title"] = fv["Title"].apply(lambda x:trunc(x,60))
    st.dataframe(sno(fv[[c for c in PCOLS if c in fv.columns]]), use_container_width=True)

# ── FBM PLAN ─────────────────────────────────────────────────────
with tabs[2]:
    st.subheader("📮 FBM Planning")
    if fbm_plan.empty:
        st.info("No FBM SKUs found — all inventory appears to be FBA.")
    else:
        fv2 = (fbm_plan[fbm_plan["Dispatch Needed"]>=min_disp]
               .sort_values("Priority",ascending=False).copy())
        fv2["Title"] = fv2["Title"].apply(lambda x:trunc(x,60))
        st.dataframe(sno(fv2[[c for c in PCOLS if c in fv2.columns]]), use_container_width=True)

# ── FC-WISE PLAN ──────────────────────────────────────────────────
with tabs[3]:
    st.subheader("🏭 FC / Warehouse-Wise Dispatch Plan")
    st.caption("Every row: FC Code · Full FC Name · City · Area · State · Cluster. "
               "Units allocated proportionally by current stock share at each location.")
    fv3 = fcv.sort_values(["FC Cluster","FC Dispatch"],ascending=[True,False]).copy()
    fv3["Title"] = fv3["Title"].apply(lambda x:trunc(x,50))
    st.dataframe(sno(fv3[[c for c in FC_COLS if c in fv3.columns]]), use_container_width=True)

# ── CLUSTER VIEW ──────────────────────────────────────────────────
with tabs[4]:
    st.subheader("🗺️ Cluster-Level Inventory View")
    st.dataframe(cl_sum, use_container_width=True, hide_index=True)
    st.divider()
    for _,row in cl_sum.iterrows():
        cl = row["FC Cluster"]
        with st.expander(
            f"📍 {cl}  —  {fmt(int(row['Dispatch Needed']))} units needed  "
            f"|  FC Codes: {row['FC Codes']}"):
            cd = fcp[fcp["FC Cluster"]==cl].copy()
            cd["Title"] = cd["Title"].apply(lambda x:trunc(x,55))
            st.dataframe(
                cd[["FC Code","FC Name","FC City","FC Area","FC State",
                    "MSKU","Title","FC Stock","Avg Daily Sale",
                    "FC DOC","FC Dispatch","FC Priority"]],
                use_container_width=True, hide_index=True)

# ── RISK ALERTS ───────────────────────────────────────────────────
with tabs[5]:
    rt1,rt2,rt3,rt4 = st.tabs(
        ["🔴 Critical (<14d)","⚫ Dead Stock","🟡 Slow Moving (>90d)","🔵 Excess"])
    RCOLS = ["MSKU","Title","Channel","Avg Daily Sale","Sellable Stock",
             "Days of Cover","Dispatch Needed","Health","Top States"]
    def rtbl(df):
        d=df.copy(); d["Title"]=d["Title"].apply(lambda x:trunc(x,55))
        return sno(d[[c for c in RCOLS if c in d.columns]])

    with rt1:
        st.markdown(f"**{len(r_crit)} SKU(s) — order immediately**")
        if not r_crit.empty:
            st.dataframe(rtbl(r_crit.sort_values("Days of Cover")), use_container_width=True)
        else: st.success("All clear — no critical SKUs!")

    with rt2:
        st.markdown(f"**{len(r_dead)} SKU(s) — stock exists but zero sales**")
        if not r_dead.empty:
            d2=r_dead.copy(); d2["Title"]=d2["Title"].apply(lambda x:trunc(x,55))
            st.dataframe(sno(d2[["MSKU","Title","Channel","Sellable Stock","All-Time Sales"]]),
                         use_container_width=True)
        else: st.success("No dead stock found.")

    with rt3:
        st.markdown(f"**{len(r_slow)} SKU(s) — days of cover > 90**")
        if not r_slow.empty:
            st.dataframe(rtbl(r_slow.sort_values("Days of Cover",ascending=False)), use_container_width=True)
        else: st.success("No slow-moving SKUs.")

    with rt4:
        st.markdown(f"**{len(r_excess)} SKU(s) — cover > {planning_days*2} days**")
        if not r_excess.empty:
            st.dataframe(rtbl(r_excess.sort_values("Days of Cover",ascending=False)), use_container_width=True)
        else: st.success("No excess stock.")

# ── TRENDS ────────────────────────────────────────────────────────
with tabs[6]:
    st.subheader("📈 Sales Trends")
    tc1,tc2 = st.columns(2)
    with tc1:
        st.markdown("**Weekly Sales**")
        if not wk_tr.empty: st.line_chart(wk_tr.set_index("Week")["Units Sold"])
    with tc2:
        st.markdown("**Monthly Sales**")
        if not mo_tr.empty: st.bar_chart(mo_tr.set_index("Month")["Units Sold"])

    st.markdown("**🔥 Top 20 SKUs by Avg Daily Sale**")
    t20=r_top20.copy(); t20["Title"]=t20["Title"].apply(lambda x:trunc(x,60))
    st.dataframe(sno(t20[["MSKU","Title","Avg Daily Sale","Days of Cover",
                           "Sellable Stock","Velocity","Top States"]]), use_container_width=True)
    st.divider()
    st.markdown("**🔍 SKU Drill-Down — Daily Sales Chart**")
    sel = st.selectbox("Select MSKU", sorted(sales["MSKU"].unique()))
    sku_d = (sales[sales["MSKU"]==sel].groupby("Date")["Qty"]
             .sum().reset_index().set_index("Date"))
    if not sku_d.empty:
        st.caption(f"**{trunc(title_map.get(sel,sel),90)}**")
        st.line_chart(sku_d["Qty"])

# ── STATE DEMAND ──────────────────────────────────────────────────
with tabs[7]:
    st.subheader("🌍 State-Level Demand Analysis")
    sd1,sd2 = st.columns([2,1])
    with sd1:
        st.dataframe(st_dem, use_container_width=True, hide_index=True)
    with sd2:
        st.markdown("**State × SKU Matrix (Top 10 SKUs)**")
        top10 = plan.nlargest(10,"Avg Daily Sale")["MSKU"].tolist()
        sm = (hist[hist["MSKU"].isin(top10)]
              .groupby(["Ship To State","MSKU"])["Qty"].sum().unstack(fill_value=0))
        if not sm.empty: st.dataframe(sm, use_container_width=True)

# ── DAMAGED STOCK ─────────────────────────────────────────────────
with tabs[8]:
    st.subheader("🔧 Damaged / Non-Sellable Stock")
    st.caption("Raise a Removal Order or Reimbursement Claim in Seller Central.")
    if dmg_report.empty:
        st.success("No damaged or non-sellable stock found.")
    else:
        dr=dmg_report.copy(); dr["Title"]=dr["Title"].apply(lambda x:trunc(x,65))
        st.dataframe(sno(dr.sort_values("Qty",ascending=False)), use_container_width=True)
        st.divider()
        st.markdown("**Full Disposition Breakdown per MSKU**")
        db=disp_bkdn.copy()
        db.insert(1,"Title",db["MSKU"].map(title_map).fillna("").apply(lambda x:trunc(x,60)))
        st.dataframe(sno(db), use_container_width=True)

# ── AMAZON FLAT FILE ──────────────────────────────────────────────
with tabs[9]:
    st.subheader("📄 Amazon FBA Shipment Flat File Generator")
    st.info(
        "1. Review / edit quantities below\n"
        "2. Select target FC → Download\n"
        "3. Seller Central → **Send to Amazon → Create New Shipment → Upload a file** ✅\n\n"
        "_Each FC requires a separate flat file._")

    ff1,ff2,ff3 = st.columns(3)
    with ff1:
        shp_label = st.text_input("Shipment Name",
            value=f"Replenishment_{datetime.now().strftime('%Y%m%d')}")
        shp_date = st.date_input("Expected Ship Date",
            value=datetime.now().date()+timedelta(days=2))
    with ff2:
        all_fcs = sorted(inv["FC Code"].unique())
        tgt_fc  = st.selectbox("Target FC", all_fcs,
            format_func=lambda x:f"{x} — {fc_name(x)}")
    with ff3:
        st.markdown(f"**FC Name:** {fc_name(tgt_fc)}")
        st.markdown(f"**City:** {fc_city(tgt_fc)}  |  **Area:** {fc_area(tgt_fc)}")
        st.markdown(f"**State:** {fc_state(tgt_fc)}")
        st.markdown(f"**Cluster:** {fc_cluster(tgt_fc)}")

    flat_base = fba_plan[fba_plan["Dispatch Needed"]>0].copy()
    flat_base["Units to Ship"]  = flat_base["Dispatch Needed"].astype(int)
    flat_base["Units Per Case"] = case_qty
    flat_base["No. of Cases"]   = np.ceil(flat_base["Units to Ship"]/case_qty).astype(int)
    flat_base["Title Short"]    = flat_base["Title"].apply(lambda x:trunc(x,80))

    st.markdown(f"**{len(flat_base)} SKUs | {fmt(flat_base['Units to Ship'].sum())} total units**")
    ecols  = ["MSKU","FNSKU","Title Short","Units to Ship","Units Per Case","No. of Cases"]
    edited = st.data_editor(
        flat_base[[c for c in ecols if c in flat_base.columns]].reset_index(drop=True),
        num_rows="dynamic", use_container_width=True)

    def make_flat_file(df, fc_code):
        parts = [p.strip() for p in ship_addr.split(",")]
        city  = parts[1] if len(parts)>1 else ""
        stz   = parts[2] if len(parts)>2 else ""
        stc   = stz.split("-")[0].strip()[:2].upper() if "-" in stz else stz[:2].upper()
        post  = stz.split("-")[1].strip() if "-" in stz else ""
        lines = ["TemplateType=FlatFileShipmentCreation\tVersion=2015.0403",""]
        mh = ["ShipmentName","ShipFromName","ShipFromAddressLine1","ShipFromCity",
              "ShipFromStateOrProvinceCode","ShipFromPostalCode","ShipFromCountryCode",
              "ShipmentStatus","LabelPrepType","AreCasesRequired",
              "DestinationFulfillmentCenterId"]
        mv = [shp_label, ship_name or "My Warehouse",
              parts[0] if parts else "", city, stc, post,
              "IN","WORKING",lbl_own,"YES" if case_packed else "NO", fc_code]
        lines += ["\t".join(mh), "\t".join(str(v) for v in mv),""]
        lines.append("\t".join(["SellerSKU","FNSKU","QuantityShipped","QuantityInCase",
                                  "PrepOwner","LabelOwner","ItemDescription","ExpectedDeliveryDate"]))
        for _,r in df.iterrows():
            qic = str(int(r.get("Units Per Case",case_qty))) if case_packed else ""
            lines.append("\t".join([
                str(r["MSKU"]), str(r.get("FNSKU","")),
                str(int(r["Units to Ship"])), qic, prep_own, lbl_own,
                str(r.get("Title Short",""))[:200], str(shp_date)]))
        return "\n".join(lines)

    flat_txt = make_flat_file(edited, tgt_fc)
    with st.expander("👁️ Preview Flat File"):
        st.code(flat_txt[:2500]+("\n...(truncated)" if len(flat_txt)>2500 else ""))

    st.download_button(
        f"📥 Download Flat File → {tgt_fc} ({fc_name(tgt_fc)})",
        flat_txt.encode("utf-8"),
        f"Amazon_FBA_{tgt_fc}_{datetime.now().strftime('%Y%m%d_%H%M')}.txt","text/plain")

    if len(all_fcs)>1:
        st.divider()
        st.markdown("**📦 Download Separate Flat File Per FC**")
        bc = st.columns(min(len(all_fcs),4))
        for i,fc in enumerate(all_fcs):
            with bc[i%4]:
                st.download_button(
                    f"📥 {fc}\n{fc_name(fc)}",
                    make_flat_file(edited,fc).encode("utf-8"),
                    f"Amazon_FBA_{fc}_{datetime.now().strftime('%Y%m%d')}.txt",
                    "text/plain", key=f"dl_{fc}")


# ═════════════════════════════════════════════════════════════════
# ⑩ EXCEL EXPORT  — 16 sheets
# ═════════════════════════════════════════════════════════════════
st.markdown("---")

def build_excel():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb  = w.book
        hdr = wb.add_format({"bold":True,"bg_color":"#1a365d","font_color":"white",
                              "border":1,"align":"center","valign":"vcenter","text_wrap":True})
        def ws(df, name):
            if df.empty:
                pd.DataFrame({"Note":["No data"]}).to_excel(w,sheet_name=name,index=False); return
            d=df.copy()
            if "Title" in d.columns: d["Title"]=d["Title"].astype(str).str[:80]
            d.to_excel(w,sheet_name=name,index=False,startrow=1)
            sh=w.sheets[name]
            for ci,cn in enumerate(d.columns):
                sh.write(0,ci,cn,hdr)
                cw=max(len(str(cn)),d[cn].astype(str).str.len().max() if not d.empty else 10)
                sh.set_column(ci,ci,min(cw+2,45))
            sh.freeze_panes(2,0)
            sh.autofilter(0,0,len(d),len(d.columns)-1)

        ws(sno(fba_plan.sort_values("Priority",ascending=False)),       "FBA Plan")
        ws(sno(fbm_plan),                                                "FBM Plan")
        ws(sno(fcp.sort_values(["FC Cluster","FC Dispatch"],
                                ascending=[True,False])),                "FC Dispatch Plan")
        ws(cl_sum,                                                        "Cluster Summary")
        ws(sno(plan.sort_values("Priority",ascending=False)),            "All SKUs")
        ws(sno(r_crit.sort_values("Days of Cover")),                     "CRITICAL Stock")
        ws(sno(r_dead),                                                   "Dead Stock")
        ws(sno(r_slow.sort_values("Days of Cover",ascending=False)),     "Slow Moving")
        ws(sno(r_excess.sort_values("Days of Cover",ascending=False)),   "Excess Stock")
        ws(sno(dmg_report.sort_values("Qty",ascending=False)),           "Damaged Stock")
        ws(disp_bkdn,                                                     "Disposition Breakdown")
        ws(st_dem,                                                        "State Demand")
        ws(wk_tr,                                                         "Weekly Trend")
        ws(mo_tr,                                                         "Monthly Trend")
        ws(sno(r_top20),                                                  "Top 20 SKUs")
        ws(pd.DataFrame([{"FC Code":k,"FC Full Name":v[0],"City":v[1],
                           "Area / Locality":v[2],"State":v[3],"Cluster":v[4]}
                          for k,v in FC_DATA.items()]),                   "FC Reference")
    out.seek(0); return out

dl1,dl2 = st.columns(2)
with dl1:
    st.download_button(
        "📥 Download Full Intelligence Report (Excel — 16 sheets)",
        data=build_excel(),
        file_name=f"FBA_Supply_Plan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with dl2:
    disp_csv = fba_plan[fba_plan["Dispatch Needed"]>0][
        ["MSKU","FNSKU","Title","Avg Daily Sale","Sellable Stock",
         "Days of Cover","Required Stock","Dispatch Needed",
         "Priority","Velocity","Top States"]].to_csv(index=False)
    st.download_button(
        "📋 Download Dispatch Plan (CSV)",
        data=disp_csv,
        file_name=f"Dispatch_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv")
