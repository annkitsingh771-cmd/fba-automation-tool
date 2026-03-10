"""
FBA Smart Supply Planner  ·  Amazon India Edition
==================================================
Rules:
  • NO state-wise sales anywhere — everything is FC-wise
  • Every plan table has one stock column per active FC
  • All reports, Excel sheets, CSVs are FC-wise
  • 75 verified Amazon India FC codes with full names / city / state / cluster
  • Daily-ledger dedup fix — always shows correct stock (not multiplied by days)
"""

import math, io, zipfile
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import streamlit as st

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="FBA Smart Supply Planner", layout="wide", page_icon="📦")
st.markdown("""<style>
[data-testid="stMetricValue"]{font-size:1.6rem;font-weight:700}
.ab{padding:.5rem 1rem;border-radius:6px;margin:.25rem 0;font-weight:600;font-size:.9rem}
.ab-r{background:#fff5f5;border-left:4px solid #e53e3e;color:#742a2a}
.ab-y{background:#fffff0;border-left:4px solid #d69e2e;color:#744210}
.ab-g{background:#f0fff4;border-left:4px solid #38a169;color:#1c4532}
.stTabs [data-baseweb="tab"]{font-size:13px;font-weight:600}
</style>""", unsafe_allow_html=True)

st.title("📦 FBA Smart Supply Planner — Amazon India")
st.caption("Inventory Ledger + MTR  →  FC-wise stock · Dispatch plan · Risk alerts · Amazon flat file")

# ══════════════════════════════════════════════════════════════════════════════
# FC MASTER  (75 verified Amazon India codes)
# Tuple: (Full Name, City, Area, State, Cluster)
# NOTE: PNQ2 = New Delhi (Mohan Co-op A-33), NOT Pune
#       PNQ1 / PNQ3 = Pune Hinjewadi / Chakan
# ══════════════════════════════════════════════════════════════════════════════
FC_DATA = {
    "SGAA":      ("Guwahati FC",               "Guwahati",        "Omshree Ind Park",           "ASSAM",           "East — Assam"),
    "DEX3":      ("New Delhi FC A28",          "New Delhi",       "Mohan Co-op, Mathura Rd",    "DELHI",           "Delhi NCR"),
    "DEX8":      ("New Delhi FC A29",          "New Delhi",       "Mohan Co-op, Mathura Rd",    "DELHI",           "Delhi NCR"),
    "PNQ2":      ("New Delhi FC A33",          "New Delhi",       "Mohan Co-op Industrial Est", "DELHI",           "Delhi NCR"),
    "XDEL":      ("Delhi XL FC",               "New Delhi",       "Delhi",                      "DELHI",           "Delhi NCR"),
    "DEL2":      ("Tauru FC",                  "Mewat",           "Village Tauru",              "HARYANA",         "Delhi NCR"),
    "DEL4":      ("Gurgaon FC",                "Gurgaon",         "Village Jamalpur",           "HARYANA",         "Delhi NCR"),
    "DEL5":      ("Manesar FC",                "Manesar",         "Binola, NH-8",               "HARYANA",         "Delhi NCR"),
    "DEL6":      ("Bilaspur FC",               "Bilaspur",        "Bilaspur",                   "HARYANA",         "Delhi NCR"),
    "DEL7":      ("Bilaspur FC 2",             "Bilaspur",        "Bilaspur",                   "HARYANA",         "Delhi NCR"),
    "DEL8":      ("Sohna FC",                  "Sohna",           "ESR Sohna Logistics Park",   "HARYANA",         "Delhi NCR"),
    "DEL8_DED5": ("Sohna ESR FC",              "Sohna",           "ESR Sohna-Ballabgarh Rd",    "HARYANA",         "Delhi NCR"),
    "DED3":      ("Farrukhnagar FC",           "Farrukhnagar",    "Gurgaon-122506",             "HARYANA",         "Delhi NCR"),
    "DED4":      ("Sohna FC 2",                "Sohna",           "ESR Sohna Logistics Park",   "HARYANA",         "Delhi NCR"),
    "DED5":      ("Sohna FC 3",                "Sohna",           "ESR Sohna Logistics Park",   "HARYANA",         "Delhi NCR"),
    "AMD1":      ("Ahmedabad Naroda FC",       "Naroda",          "Naroda, Ahmedabad",          "GUJARAT",         "Gujarat West"),
    "AMD2":      ("Changodar FC",              "Changodar",       "Gallops Ind Park",           "GUJARAT",         "Gujarat West"),
    "SUB1":      ("Surat FC",                  "Surat",           "Surat",                      "GUJARAT",         "Gujarat West"),
    "BLR4":      ("Devanahalli FC",            "Devanahalli",     "Hitech Aerospace Park",      "KARNATAKA",       "Bangalore"),
    "BLR5":      ("Bommasandra FC",            "Bommasandra",     "Hosakote, Bengaluru",        "KARNATAKA",       "Bangalore"),
    "BLR6":      ("Nelamangala FC",            "Nelamangala",     "Bengaluru",                  "KARNATAKA",       "Bangalore"),
    "BLR7":      ("Hoskote FC",                "Hoskote",         "Anekal Taluk",               "KARNATAKA",       "Bangalore"),
    "BLR8":      ("Devanahalli FC 2",          "Devanahalli",     "Hitech Aerospace Park",      "KARNATAKA",       "Bangalore"),
    "BLR10":     ("Kudlu Gate FC",             "Kudlu Gate",      "Bengaluru",                  "KARNATAKA",       "Bangalore"),
    "BLR12":     ("Attibele FC",               "Attibele",        "Bengaluru",                  "KARNATAKA",       "Bangalore"),
    "BLR13":     ("Jigani FC",                 "Jigani",          "Bengaluru",                  "KARNATAKA",       "Bangalore"),
    "BLR14":     ("Anekal FC",                 "Anekal",          "Bengaluru",                  "KARNATAKA",       "Bangalore"),
    "SCJA":      ("Bangalore SCJA FC",         "Bangalore",       "Bengaluru",                  "KARNATAKA",       "Bangalore"),
    "XSAB":      ("Bangalore XS FC",           "Bangalore",       "Bengaluru",                  "KARNATAKA",       "Bangalore"),
    "SBLA":      ("Bangalore SBLA FC",         "Bangalore",       "Bengaluru",                  "KARNATAKA",       "Bangalore"),
    "SIDA":      ("Indore FC",                 "Indore",          "Village Pipliya Kumhar",     "MADHYA PRADESH",  "Central"),
    "FBHB":      ("Bhopal FC",                 "Bhopal",          "Govindpura Ind Area",        "MADHYA PRADESH",  "Central"),
    "FIDA":      ("Bhopal FC 2",               "Bhopal",          "Govindpura Ind Area",        "MADHYA PRADESH",  "Central"),
    "IND1":      ("Indore FC 2",               "Indore",          "Indore",                     "MADHYA PRADESH",  "Central"),
    "BOM1":      ("Bhiwandi FC",               "Bhiwandi",        "Bhiwandi, Thane",            "MAHARASHTRA",     "Mumbai West"),
    "BOM3":      ("Nashik FC",                 "Nashik",          "Nashik",                     "MAHARASHTRA",     "Mumbai West"),
    "BOM4":      ("Vashere FC",                "Bhiwandi",        "Village Vashere",            "MAHARASHTRA",     "Mumbai West"),
    "BOM5":      ("Bhiwandi 2 FC",             "Bhiwandi",        "Village Vashere",            "MAHARASHTRA",     "Mumbai West"),
    "BOM7":      ("Bhiwandi 3 FC",             "Bhiwandi",        "Village Vahuli",             "MAHARASHTRA",     "Mumbai West"),
    "ISK3":      ("Bhiwandi ISK3 FC",          "Bhiwandi",        "Village Pise",               "MAHARASHTRA",     "Mumbai West"),
    "PNQ1":      ("Pune Hinjewadi FC",         "Pune",            "Hinjewadi",                  "MAHARASHTRA",     "Mumbai West"),
    "PNQ3":      ("Pune Chakan FC",            "Pune",            "Village Ambethan, Khed",     "MAHARASHTRA",     "Mumbai West"),
    "SAMB":      ("Mumbai SAMB FC",            "Mumbai",          "Mumbai",                     "MAHARASHTRA",     "Mumbai West"),
    "XBOM":      ("Mumbai XL FC",              "Bhiwandi",        "Bhiwandi",                   "MAHARASHTRA",     "Mumbai West"),
    "ATX1":      ("Ludhiana FC",               "Ludhiana",        "Near Katana Sahib Gurdwara", "PUNJAB",          "North Punjab"),
    "LDH1":      ("Ludhiana FC 2",             "Ludhiana",        "Ludhiana",                   "PUNJAB",          "North Punjab"),
    "RAJ1":      ("Rajpura FC",                "Rajpura",         "Rajpura",                    "PUNJAB",          "North Punjab"),
    "JAI1":      ("Jaipur FC",                 "Jaipur",          "Bagru",                      "RAJASTHAN",       "North Rajasthan"),
    "JPX1":      ("Jaipur JPX1 FC",            "Jaipur",          "Jhotwara Ind Area",          "RAJASTHAN",       "North Rajasthan"),
    "JPX2":      ("Bagru FC",                  "Bagru",           "Bagru, Sanganer",            "RAJASTHAN",       "North Rajasthan"),
    "MAA1":      ("Irungattukottai FC",        "Irungattukottai", "Sriperumbudur Tk",           "TAMIL NADU",      "Chennai"),
    "MAA2":      ("Ponneri FC",                "Ponneri",         "Thiruvallur Dist",           "TAMIL NADU",      "Chennai"),
    "MAA3":      ("Sriperumbudur FC",          "Sriperumbudur",   "Chennai",                    "TAMIL NADU",      "Chennai"),
    "MAA4":      ("Ambattur FC",               "Ambattur",        "Thiruvallur Dist",           "TAMIL NADU",      "Chennai"),
    "MAA5":      ("Kanchipuram FC",            "Kanchipuram",     "Kanchipuram Dist",           "TAMIL NADU",      "Chennai"),
    "CJB1":      ("Coimbatore FC",             "Coimbatore",      "Palladam Main Rd",           "TAMIL NADU",      "Chennai"),
    "SMAB":      ("Chennai SMAB FC",           "Chennai",         "Chennai",                    "TAMIL NADU",      "Chennai"),
    "COK1":      ("Kochi FC",                  "Kochi",           "Kochi",                      "KERALA",          "South Kerala"),
    "HYD3":      ("Shamshabad FC",             "Shamshabad",      "Mamidipally Village",        "TELANGANA",       "Hyderabad"),
    "HYD6":      ("Kothur FC",                 "Kothur",          "Mahbubnagar Dist",           "TELANGANA",       "Hyderabad"),
    "HYD7":      ("Medchal FC",                "Medchal",         "Hyderabad",                  "TELANGANA",       "Hyderabad"),
    "HYD8":      ("Shamshabad FC 2",           "Shamshabad",      "Hyderabad",                  "TELANGANA",       "Hyderabad"),
    "HYD8_HYD3": ("Shamshabad Mamidipally FC", "Shamshabad",      "Mamidipally Village",        "TELANGANA",       "Hyderabad"),
    "HYD9":      ("Pedda Amberpet FC",         "Pedda Amberpet",  "Hyderabad",                  "TELANGANA",       "Hyderabad"),
    "HYD10":     ("Ghatkesar FC",              "Ghatkesar",       "Hyderabad",                  "TELANGANA",       "Hyderabad"),
    "HYD11":     ("Kompally FC",               "Kompally",        "Hyderabad",                  "TELANGANA",       "Hyderabad"),
    "LKO1":      ("Lucknow FC",                "Lucknow",         "Village Bhukapur",           "UTTAR PRADESH",   "North UP"),
    "AGR1":      ("Agra FC",                   "Agra",            "Agra",                       "UTTAR PRADESH",   "North UP"),
    "SLDK":      ("Kishanpur FC",              "Bijnour",         "Kishanpur Kodia",            "UTTAR PRADESH",   "North UP"),
    "CCU1":      ("Kolkata FC",                "Kolkata",         "Rajarhat",                   "WEST BENGAL",     "Kolkata East"),
    "CCU2":      ("Kolkata FC 2",              "Kolkata",         "Kolkata",                    "WEST BENGAL",     "Kolkata East"),
    "CCX1":      ("Howrah FC",                 "Howrah",          "Panchla, Raghudevpur",       "WEST BENGAL",     "Kolkata East"),
    "CCX2":      ("Howrah FC 2",               "Howrah",          "Panchla",                    "WEST BENGAL",     "Kolkata East"),
    "PAT1":      ("Patna FC",                  "Patna",           "Patna",                      "BIHAR",           "Kolkata East"),
    "BBS1":      ("Bhubaneswar FC",            "Bhubaneswar",     "Bhubaneswar",                "ODISHA",          "Kolkata East"),
}


def _fc(code, idx, fallback):
    e = FC_DATA.get(str(code).upper().strip())
    return e[idx] if e else fallback

def fc_name(c):    return _fc(c, 0, str(c))
def fc_city(c):    return _fc(c, 1, "Unknown")
def fc_state(c):   return _fc(c, 3, "Unknown")
def fc_cluster(c): return _fc(c, 4, "Other")


# ══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ══════════════════════════════════════════════════════════════════════════════
def read_file(up):
    try:
        n = up.name.lower()
        if n.endswith(".zip"):
            with zipfile.ZipFile(up) as z:
                for m in z.namelist():
                    if m.lower().endswith(".csv"):
                        return pd.read_csv(z.open(m), low_memory=False)
            return pd.DataFrame()
        if n.endswith((".xlsx",".xls")):
            return pd.read_excel(up)
        return pd.read_csv(up, low_memory=False)
    except Exception as e:
        st.error(f"Cannot read {up.name}: {e}")
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
    try:    return f"{int(n):,}"
    except: return str(n)


def trunc(s, n=55):
    s = str(s)
    return s[:n]+"…" if len(s)>n else s


def health_tag(doc, pd_):
    if doc <= 0:    return "⚫ No Stock"
    if doc < 14:    return "🔴 Critical"
    if doc < 30:    return "🟠 Low"
    if doc < pd_:   return "🟢 Healthy"
    if doc < pd_*2: return "🟡 Excess"
    return "🔵 Overstocked"


def vel_tag(avg):
    if avg <= 0:  return "⚫ Dead"
    if avg < 0.5: return "🔵 Slow"
    if avg < 2:   return "🟡 Medium"
    if avg < 5:   return "🟢 Fast"
    return "🔥 Top Seller"


def excel_sheet(writer, df, sheet_name):
    """Write df to Excel sheet with styled header, autofilter, freeze."""
    if df is None or df.empty:
        pd.DataFrame({"Note": ["No data"]}).to_excel(
            writer, sheet_name=sheet_name, index=False)
        return
    d = df.copy()
    if "Title" in d.columns:
        d["Title"] = d["Title"].astype(str).str[:80]
    d.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
    wb  = writer.book
    sh  = writer.sheets[sheet_name]
    hdr = wb.add_format({"bold":True,"bg_color":"#1a365d","font_color":"white",
                          "border":1,"align":"center","valign":"vcenter","text_wrap":True})
    for ci, cn in enumerate(d.columns):
        sh.write(0, ci, cn, hdr)
        cw = max(len(str(cn)), d[cn].astype(str).str.len().max() if not d.empty else 8)
        sh.set_column(ci, ci, min(cw+2, 42))
    sh.freeze_panes(2, 0)
    sh.autofilter(0, 0, len(d), len(d.columns)-1)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
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
    case_qty    = st.number_input("Units Per Carton", 1, 500, 12)
    case_packed = st.checkbox("Case-Packed Shipment", value=False)
    prep_own    = st.selectbox("Prep Ownership",  ["AMAZON","SELLER"])
    lbl_own     = st.selectbox("Label Ownership", ["AMAZON","SELLER"])
    st.divider()
    st.subheader("🎯 Filters")
    min_disp     = st.number_input("Min Dispatch Units to Show", 0, 99999, 1)
    all_clusters = sorted({v[4] for v in FC_DATA.values()})
    foc_cluster  = st.multiselect("Focus Clusters", all_clusters)


# ══════════════════════════════════════════════════════════════════════════════
# FILE UPLOADS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("### 📁 Upload Files")
cu1, cu2 = st.columns(2)
with cu1:
    mtr_files = st.file_uploader(
        "📊 MTR / Sales Report — multiple files OK (CSV, ZIP, XLSX)",
        type=["csv","zip","xlsx"], accept_multiple_files=True)
with cu2:
    inv_file = st.file_uploader(
        "🏭 Amazon Inventory Ledger (CSV, ZIP, XLSX)",
        type=["csv","zip","xlsx"])

with st.expander("📋 Expected File Formats"):
    st.markdown("""
**Inventory Ledger** — Seller Central → Reports → Fulfillment → Inventory Ledger

| Column | Example |
|---|---|
| `Date` | 02/12/2026 |
| `FNSKU` | X00275KJZP |
| `MSKU` | BR-899 |
| `Title` | Bracelet |
| `Disposition` | SELLABLE / CUSTOMER_DAMAGED / CARRIER_DAMAGED |
| `Ending Warehouse Balance` | 18 |
| `Location` | LKO1 |

**MTR Sales Report** — Seller Central → Reports → Tax → MTR

| Column | Example |
|---|---|
| `Sku` | BR-899 |
| `Quantity` | 5 |
| `Shipment Date` | 2026-01-15 |
| `Fulfilment` | AFN / MFN |
""")

if not (mtr_files and inv_file):
    st.info("👆 Upload both files to start planning.")
    with st.expander("🗺️ All 75 Verified Amazon India FC Codes"):
        st.dataframe(pd.DataFrame([{
            "FC Code":k,"FC Name":v[0],"City":v[1],
            "Area":v[2],"State":v[3],"Cluster":v[4]}
            for k,v in FC_DATA.items()]),
            use_container_width=True, hide_index=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — PARSE INVENTORY LEDGER
# ══════════════════════════════════════════════════════════════════════════════
inv_raw = read_file(inv_file)
if inv_raw.empty:
    st.error("Inventory file could not be read — check format."); st.stop()
inv_raw.columns = inv_raw.columns.str.strip()

C_SKU   = find_col(inv_raw, ["MSKU","Sku","SKU","ASIN"],                     ["msku","sku","asin"])
C_TITLE = find_col(inv_raw, ["Title","Product Name","Description"],           ["title","product"])
C_QTY   = find_col(inv_raw, ["Ending Warehouse Balance","Quantity","Qty"],    ["ending","balance"])
C_LOC   = find_col(inv_raw, ["Location","Warehouse Code","FC Code","FC"],     ["location","warehouse","fc code"])
C_DISP  = find_col(inv_raw, ["Disposition"],                                  ["disposition"])
C_FNSKU = find_col(inv_raw, ["FNSKU"],                                        ["fnsku"])
C_DATE  = find_col(inv_raw, ["Date"],                                         ["date"])

for lbl, col in [("SKU/MSKU",C_SKU),("Ending Balance",C_QTY),("Location/FC",C_LOC)]:
    if not col:
        st.error(f"Inventory file missing required column: **{lbl}**")
        st.write("Columns found:", list(inv_raw.columns)); st.stop()

inv = inv_raw.copy()
inv["MSKU"]        = inv[C_SKU].astype(str).str.strip()
inv["Stock"]       = pd.to_numeric(inv[C_QTY], errors="coerce").fillna(0)
inv["FC Code"]     = inv[C_LOC].astype(str).str.upper().str.strip()
inv["Title"]       = inv[C_TITLE].astype(str).str.strip() if C_TITLE else ""
inv["FNSKU"]       = inv[C_FNSKU].astype(str).str.strip() if C_FNSKU else ""
inv["Disposition"] = inv[C_DISP].astype(str).str.upper().str.strip() if C_DISP else "SELLABLE"

# ── DAILY DEDUP FIX ────────────────────────────────────────────────────────
# Amazon Inventory Ledger = one row PER DAY per SKU/FC/Disposition.
# Summing all rows gives stock × days in report. Fix: keep LATEST row per group.
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
inv["FC State"]   = inv["FC Code"].apply(fc_state)
inv["FC Cluster"] = inv["FC Code"].apply(fc_cluster)

# Build lookup maps
fnsku_map, title_map = {}, {}
for _, r in inv.drop_duplicates("MSKU").iterrows():
    title_map[r["MSKU"]] = r["Title"]
    if r["FNSKU"] not in ("nan",""):
        fnsku_map[r["MSKU"]] = r["FNSKU"]

# Split sellable / damaged
inv_sell = inv[inv["Disposition"]=="SELLABLE"].copy()
inv_dmg  = inv[inv["Disposition"]!="SELLABLE"].copy()

# ── FC-WISE STOCK TABLES ───────────────────────────────────────────────────
# Long table: 1 row per MSKU × FC
fc_stock_long = (inv_sell
    .groupby(["MSKU","FC Code","FC Name","FC City","FC State","FC Cluster"])["Stock"]
    .sum().reset_index().rename(columns={"Stock":"FC Stock"}))

# Total sellable / damaged per MSKU
sku_sell_total = (inv_sell.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock":"Total Sellable"}))
sku_dmg_total  = (inv_dmg.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock":"Total Damaged"}))

# Disposition breakdown & damaged detail
disp_bkdn  = (inv.groupby(["MSKU","Disposition"])["Stock"]
    .sum().unstack(fill_value=0).reset_index())
dmg_detail = (inv_dmg.groupby(["MSKU","Disposition"])["Stock"]
    .sum().reset_index().rename(columns={"Stock":"Qty"}))
dmg_detail["Title"] = dmg_detail["MSKU"].map(title_map).fillna("")

# FC inventory summary (code + name + city + state)
fc_inv_summary = (fc_stock_long
    .groupby(["FC Code","FC Name","FC City","FC State","FC Cluster"])["FC Stock"]
    .sum().reset_index().rename(columns={"FC Stock":"Sellable Stock"})
    .sort_values("Sellable Stock", ascending=False))

# FC dispatch summary (built after fcp, reused in tab and Excel)
# → computed after fcp is built below

# ── FC-WISE STOCK PIVOT ────────────────────────────────────────────────────
# Creates one column per active FC: "LKO1 | Lucknow FC"
_active_fcs = sorted(inv_sell["FC Code"].unique())

_piv = (inv_sell.groupby(["MSKU","FC Code"])["Stock"].sum().reset_index())
_piv["col"] = _piv["FC Code"].apply(lambda c: f"{c} | {fc_name(c)}")
fc_pivot = (_piv.pivot_table(index="MSKU", columns="col",
                              values="Stock", aggfunc="sum", fill_value=0)
             .reset_index())
fc_pivot.columns.name = None
FC_STOCK_COLS = [c for c in fc_pivot.columns if c != "MSKU"]
FC_CODE_FROM  = {f"{c} | {fc_name(c)}": c for c in _active_fcs}


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — PARSE SALES / MTR
# ══════════════════════════════════════════════════════════════════════════════
raw_sales = [read_file(f) for f in mtr_files]
raw_sales = [d for d in raw_sales if not d.empty]
if not raw_sales:
    st.error("Could not read any MTR / sales file."); st.stop()

sr = pd.concat(raw_sales, ignore_index=True)

CS_SKU  = find_col(sr, ["Sku","SKU","MSKU","Item SKU"],                   ["sku","msku"])
CS_QTY  = find_col(sr, ["Quantity","Qty","Units"],                        ["qty","unit","quant"])
CS_DATE = find_col(sr, ["Shipment Date","Purchase Date","Order Date","Date"], ["date"])
CS_FT   = find_col(sr, ["Fulfilment","Fulfillment",
                          "Fulfilment Channel","Fulfillment Channel"],     ["fulfil","channel"])

for lbl, col in [("SKU",CS_SKU),("Quantity",CS_QTY),("Date",CS_DATE)]:
    if not col:
        st.error(f"Sales file missing required column: **{lbl}**")
        st.write("Columns found:", list(sr.columns)); st.stop()

sales = sr.rename(columns={CS_SKU:"MSKU", CS_QTY:"Qty", CS_DATE:"Date"}).copy()
sales["MSKU"]    = sales["MSKU"].astype(str).str.strip()
sales["Qty"]     = pd.to_numeric(sales["Qty"], errors="coerce").fillna(0)
sales["Date"]    = pd.to_datetime(sales["Date"], dayfirst=True, errors="coerce")
sales            = sales.dropna(subset=["Date"])
sales            = sales[sales["Qty"]>0].copy()
FT_MAP           = {"AFN":"FBA","FBA":"FBA","AMAZON_FULFILLED":"FBA",
                    "MFN":"FBM","FBM":"FBM","MERCHANT_FULFILLED":"FBM"}
sales["Channel"] = (sales[CS_FT].astype(str).str.upper().str.strip().map(FT_MAP).fillna("FBA")
                    if CS_FT else "FBA")

s_min   = sales["Date"].min();  s_max = sales["Date"].max()
up_days = max((s_max-s_min).days+1, 1)
window  = (30 if "30" in sales_basis else
           60 if "60" in sales_basis else
           90 if "90" in sales_basis else up_days)
cutoff  = s_max - pd.Timedelta(days=window-1)
hist    = sales[sales["Date"]>=cutoff].copy() if window<up_days else sales.copy()
h_days  = max((hist["Date"].max()-hist["Date"].min()).days+1, 1)

# Aggregates
ch_hist  = (hist.groupby(["MSKU","Channel"])["Qty"].sum()
            .reset_index().rename(columns={"Qty":"Hist Sales"}))
ch_full  = (sales.groupby(["MSKU","Channel"])["Qty"].sum()
            .reset_index().rename(columns={"Qty":"All-Time Sales"}))
ch_daily = hist.groupby(["MSKU","Channel","Date"])["Qty"].sum().reset_index()
ch_std   = (ch_daily.groupby(["MSKU","Channel"])["Qty"].std()
            .reset_index().rename(columns={"Qty":"Demand StdDev"}))

# Trend tables (no state)
mo_tr = (sales.assign(Month=sales["Date"].dt.to_period("M").astype(str))
    .groupby(["Month","Channel"])["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units Sold"}).sort_values("Month"))
wk_tr = (sales.assign(Week=sales["Date"].dt.to_period("W").astype(str))
    .groupby("Week")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units Sold"}).sort_values("Week"))


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — PLANNING TABLE  (FC-wise stock columns embedded)
# ══════════════════════════════════════════════════════════════════════════════
keys = pd.concat([
    ch_hist[["MSKU","Channel"]],
    sku_sell_total.assign(Channel="FBA")[["MSKU","Channel"]]
], ignore_index=True).drop_duplicates()

plan = keys.copy()
for df_, on_ in [
    (ch_hist,       ["MSKU","Channel"]),
    (ch_full,       ["MSKU","Channel"]),
    (ch_std,        ["MSKU","Channel"]),
    (sku_sell_total, "MSKU"),
    (sku_dmg_total,  "MSKU"),
    (fc_pivot,       "MSKU"),   # ← FC stock columns join here
]:
    plan = plan.merge(df_, on=on_, how="left")

for c in ["Hist Sales","All-Time Sales","Total Sellable","Total Damaged","Demand StdDev"]:
    plan[c] = pd.to_numeric(plan[c], errors="coerce").fillna(0)
for c in FC_STOCK_COLS:
    plan[c] = pd.to_numeric(plan[c], errors="coerce").fillna(0)

plan["Title"]            = plan["MSKU"].map(title_map).fillna("")
plan["FNSKU"]            = plan["MSKU"].map(fnsku_map).fillna("")
plan["Sales Days Used"]  = h_days
plan["Avg Daily Sale"]   = (plan["Hist Sales"] / h_days).round(4)
plan["Safety Stock"]     = (z_val * plan["Demand StdDev"] * math.sqrt(planning_days)).round(0)
plan["Base Req"]         = (plan["Avg Daily Sale"] * planning_days).round(0)
plan["Required Stock"]   = (plan["Base Req"] + plan["Safety Stock"]).round(0)
plan["Dispatch Needed"]  = (plan["Required Stock"] - plan["Total Sellable"]).clip(lower=0).round(0)
plan["Days of Cover"]    = np.where(
    plan["Avg Daily Sale"]>0,
    (plan["Total Sellable"]/plan["Avg Daily Sale"]).round(1),
    np.where(plan["Total Sellable"]>0, 9999, 0))
plan["Health"]    = plan.apply(lambda r: health_tag(r["Days of Cover"], planning_days), axis=1)
plan["Velocity"]  = plan["Avg Daily Sale"].apply(vel_tag)

_mx = plan["Avg Daily Sale"].max() or 1
plan["Priority"]  = plan.apply(lambda r: round(
    (max(0,(planning_days-min(r["Days of Cover"],planning_days))/planning_days)*0.65
     + min(r["Avg Daily Sale"]/_mx,1)*0.35)*100, 1)
    if r["Avg Daily Sale"]>0 else 0, axis=1)

fba_plan = plan[plan["Channel"]=="FBA"].copy()
fbm_plan = plan[plan["Channel"]=="FBM"].copy()

# Column order: SKU info | Total Sellable | one col per FC | Damaged | Plan calcs
PLAN_CORE  = ["MSKU","FNSKU","Title","Avg Daily Sale"]
PLAN_STOCK = ["Total Sellable"] + FC_STOCK_COLS + ["Total Damaged"]
PLAN_CALC  = ["Safety Stock","Base Req","Required Stock","Dispatch Needed",
              "Days of Cover","Health","Velocity","Priority"]
PLAN_COLS  = PLAN_CORE + PLAN_STOCK + PLAN_CALC


def plan_view(df):
    d = df.copy()
    d["Title"] = d["Title"].apply(lambda x: trunc(x, 50))
    return sno(d[[c for c in PLAN_COLS if c in d.columns]])


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — FC-WISE DISPATCH ALLOCATION
# ══════════════════════════════════════════════════════════════════════════════
fcp = fc_stock_long.merge(
    plan[["MSKU","Channel","Avg Daily Sale","Dispatch Needed","Total Sellable",
          "Required Stock","Title","FNSKU","Priority","Days of Cover"]],
    on="MSKU", how="left")

fcp["FC DOC"]      = np.where(
    fcp["Avg Daily Sale"]>0,
    (fcp["FC Stock"]/fcp["Avg Daily Sale"]).round(1),
    np.where(fcp["FC Stock"]>0, 9999, 0))
fcp["FC Health"]   = fcp.apply(lambda r: health_tag(r["FC DOC"], planning_days), axis=1)
fcp["FC Urgency"]  = fcp["FC DOC"].apply(
    lambda d: "🔴 Urgent" if d<14 else ("🟠 Soon" if d<30 else "🟢 OK"))

_tot = (fcp.groupby(["MSKU","Channel"])["FC Stock"].sum()
        .reset_index().rename(columns={"FC Stock":"_T"}))
fcp  = fcp.merge(_tot, on=["MSKU","Channel"], how="left")
fcp["FC Share"]    = (fcp["FC Stock"]/fcp["_T"].replace(0,1)).fillna(1.0)
fcp["FC Dispatch"] = (fcp["Dispatch Needed"]*fcp["FC Share"]).round(0)
fcp.drop(columns=["_T"], inplace=True)

# Filtered view for display
fcv = fcp.copy()
if foc_cluster: fcv = fcv[fcv["FC Cluster"].isin(foc_cluster)]
fcv = fcv[fcv["FC Dispatch"] >= min_disp]

FC_DISP_COLS = ["FC Code","FC Name","FC City","FC State","FC Cluster",
                "MSKU","Title","FC Stock","Avg Daily Sale",
                "FC DOC","FC Health","FC Dispatch","FC Urgency"]

# ── FC dispatch summary: 1 row per FC ─────────────────────────────────────
fc_disp_summary = (fcp
    .groupby(["FC Code","FC Name","FC City","FC State","FC Cluster"])
    .agg(Total_Stock    =("FC Stock",    "sum"),
         Dispatch_Needed=("FC Dispatch", "sum"),
         SKUs           =("MSKU",        "nunique"),
         Avg_DOC        =("FC DOC",      "mean"))
    .reset_index()
    .rename(columns={"Total_Stock":"Total Stock","Dispatch_Needed":"Dispatch Needed",
                     "SKUs":"SKUs","Avg_DOC":"Avg Days of Cover"})
    .sort_values("Dispatch Needed", ascending=False))
fc_disp_summary["Avg Days of Cover"] = fc_disp_summary["Avg Days of Cover"].round(1)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — CLUSTER SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
cl_sum = (fcp.groupby("FC Cluster").agg(
    FC_Codes       =("FC Code",    lambda x:", ".join(sorted(set(x)))),
    FC_Names       =("FC Name",    lambda x:" | ".join(sorted(set(x)))),
    FC_Cities      =("FC City",    lambda x:", ".join(sorted(set(x)))),
    Total_Stock    =("FC Stock",   "sum"),
    Dispatch_Needed=("FC Dispatch","sum"),
    Avg_DOC        =("FC DOC",     "mean"),
    SKUs           =("MSKU",       "nunique"),
).reset_index().rename(columns={
    "FC_Codes":"FC Codes","FC_Names":"FC Names","FC_Cities":"Cities",
    "Total_Stock":"Total Stock","Dispatch_Needed":"Dispatch Needed",
    "Avg_DOC":"Avg Days of Cover","SKUs":"Unique SKUs"}))
cl_sum["Avg Days of Cover"] = cl_sum["Avg Days of Cover"].round(1)
cl_sum = sno(cl_sum.sort_values("Dispatch Needed", ascending=False))


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — RISK TABLES
# ══════════════════════════════════════════════════════════════════════════════
r_crit   = plan[(plan["Days of Cover"]<14)   & (plan["Avg Daily Sale"]>0)].copy()
r_dead   = plan[(plan["Avg Daily Sale"]==0)  & (plan["Total Sellable"]>0)].copy()
r_slow   = plan[(plan["Avg Daily Sale"]>0)   & (plan["Days of Cover"]>90)].copy()
r_excess = plan[plan["Days of Cover"]>planning_days*2].copy()
r_top20  = plan[plan["Avg Daily Sale"]>0].nlargest(20,"Avg Daily Sale").copy()

RISK_COLS = (["MSKU","Title","Channel","Avg Daily Sale","Total Sellable"]
             + FC_STOCK_COLS
             + ["Days of Cover","Dispatch Needed","Health","Velocity"])

def risk_view(df):
    d = df.copy()
    d["Title"] = d["Title"].apply(lambda x: trunc(x, 50))
    return sno(d[[c for c in RISK_COLS if c in d.columns]])


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — FC VELOCITY TABLE
# ══════════════════════════════════════════════════════════════════════════════
fc_velocity = (plan[["MSKU","Title","Avg Daily Sale","Velocity"]]
    .drop_duplicates("MSKU")
    .merge(fc_stock_long[["MSKU","FC Code","FC Name","FC City","FC Cluster","FC Stock"]],
           on="MSKU", how="left")
    .sort_values(["FC Code","Avg Daily Sale"], ascending=[True,False]))


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — EXECUTIVE DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("### 📊 Executive Dashboard")

tot_skus = plan["MSKU"].nunique()
tot_sell = int(sku_sell_total["Total Sellable"].sum())
tot_dmg  = int(sku_dmg_total["Total Damaged"].sum()) if not sku_dmg_total.empty else 0
tot_disp = int(fba_plan["Dispatch Needed"].sum())
n_fcs    = len(_active_fcs)
avg_doc  = plan[plan["Avg Daily Sale"]>0]["Days of Cover"].replace(9999, np.nan).mean()

m1,m2,m3,m4,m5,m6 = st.columns(6)
m1.metric("Total SKUs",        fmt(tot_skus))
m2.metric("Active FCs",        fmt(n_fcs))
m3.metric("Total Sellable",    fmt(tot_sell))
m4.metric("Units to Dispatch", fmt(tot_disp))
m5.metric("🔴 Critical SKUs",  fmt(len(r_crit)))
m6.metric("Avg Days of Cover", f"{avg_doc:.0f}d" if pd.notna(avg_doc) else "N/A")

st.markdown(
    f'<div class="ab ab-g">✅ Inventory snapshot: <b>{inv_date_lbl}</b> | '
    f'{len(inv)} rows after daily dedup (raw: {len(inv_raw)}) | '
    f'FCs: {", ".join(_active_fcs)}</div>', unsafe_allow_html=True)
if r_crit.shape[0]:
    st.markdown(f'<div class="ab ab-r">🚨 {len(r_crit)} SKU(s) below 14 days — replenish immediately!</div>', unsafe_allow_html=True)
if r_dead.shape[0]:
    st.markdown(f'<div class="ab ab-y">⚠️ {len(r_dead)} SKU(s) have stock but ZERO sales — fix listing or liquidate.</div>', unsafe_allow_html=True)
if r_excess.shape[0]:
    st.markdown(f'<div class="ab ab-y">📦 {len(r_excess)} SKU(s) overstocked (>{planning_days*2}d) — pause replenishment.</div>', unsafe_allow_html=True)
if tot_dmg:
    st.markdown(f'<div class="ab ab-y">🔧 {fmt(tot_dmg)} damaged units — raise Removal/Reimbursement in Seller Central.</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — TABS
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs([
    "🏠 Overview",
    "📦 FBA Plan",
    "📮 FBM Plan",
    "🏭 FC-Wise Dispatch",
    "🗺️ Cluster View",
    "🚨 Risk Alerts",
    "📈 Sales Trends",
    "🔧 Damaged Stock",
    "📄 Amazon Flat File",
])

# ── TAB 0: OVERVIEW ───────────────────────────────────────────────────────────
with tabs[0]:
    st.subheader("🏠 Planning Overview")
    st.info(
        f"**Sales data:** {s_min.strftime('%d %b %Y')} → {s_max.strftime('%d %b %Y')} "
        f"({up_days} days)  |  **Sales basis:** {sales_basis} ({h_days} days used)  \n"
        f"**Planning horizon:** {planning_days} days  |  "
        f"**Service level:** {svc_lvl} (Z={z_val})  |  "
        f"**Active FCs:** {', '.join(_active_fcs)}")

    oa, ob = st.columns(2)
    with oa:
        st.markdown("**🔝 Top 5 Most Urgent SKUs**")
        top5 = (plan[plan["Dispatch Needed"]>0].nlargest(5,"Priority")
                [["MSKU","Title","Avg Daily Sale","Days of Cover",
                  "Dispatch Needed","Priority"]].copy())
        top5["Title"] = top5["Title"].apply(lambda x: trunc(x,40))
        st.dataframe(top5, use_container_width=True, hide_index=True)
    with ob:
        st.markdown("**📊 Monthly Sales Trend**")
        if not mo_tr.empty:
            pm = mo_tr.pivot_table(index="Month", columns="Channel",
                                    values="Units Sold", aggfunc="sum", fill_value=0)
            st.bar_chart(pm)

    st.markdown("**🏭 Inventory by FC — Code · Name · City · State · Sellable Stock**")
    st.dataframe(sno(fc_inv_summary), use_container_width=True, hide_index=True)

    st.markdown("**🗺️ Cluster Summary**")
    st.dataframe(
        cl_sum[["FC Cluster","FC Codes","FC Names","Cities",
                "Total Stock","Dispatch Needed","Avg Days of Cover","Unique SKUs"]],
        use_container_width=True, hide_index=True)


# ── TAB 1: FBA PLAN ───────────────────────────────────────────────────────────
with tabs[1]:
    st.subheader("📦 FBA Inventory Plan")
    st.caption(
        f"{len(fba_plan)} SKUs  |  {planning_days}-day plan @ {svc_lvl}  |  "
        f"FC stock columns: {', '.join(_active_fcs)}")

    fv = (fba_plan[fba_plan["Dispatch Needed"]>=min_disp]
          .sort_values("Priority", ascending=False).copy())
    st.dataframe(plan_view(fv), use_container_width=True)

    st.markdown("**📊 Stock & Dispatch by FC**")
    rows = []
    for col in FC_STOCK_COLS:
        code = FC_CODE_FROM.get(col, col.split(" | ")[0])
        rows.append({
            "FC Code":           code,
            "FC Name":           fc_name(code),
            "City":              fc_city(code),
            "Cluster":           fc_cluster(code),
            "Total Stock in FC": int(fba_plan[col].sum()),
            "SKUs with Stock":   int((fba_plan[col]>0).sum()),
        })
    st.dataframe(
        sno(pd.DataFrame(rows).sort_values("Total Stock in FC", ascending=False)),
        use_container_width=True, hide_index=True)


# ── TAB 2: FBM PLAN ───────────────────────────────────────────────────────────
with tabs[2]:
    st.subheader("📮 FBM Inventory Plan")
    if fbm_plan.empty:
        st.info("No FBM SKUs found — all sales appear to be FBA.")
    else:
        fv2 = (fbm_plan[fbm_plan["Dispatch Needed"]>=min_disp]
               .sort_values("Priority", ascending=False).copy())
        st.dataframe(plan_view(fv2), use_container_width=True)


# ── TAB 3: FC-WISE DISPATCH ───────────────────────────────────────────────────
with tabs[3]:
    st.subheader("🏭 FC-Wise Dispatch Plan")
    st.caption("Dispatch units allocated proportionally by each FC's share of total stock.")

    st.markdown("**📋 FC Dispatch Summary — 1 row per FC**")
    st.dataframe(sno(fc_disp_summary), use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("**📦 SKU-Level Detail — every row is 1 SKU at 1 FC**")
    fv3 = fcv.sort_values(["FC Cluster","FC Dispatch"], ascending=[True,False]).copy()
    fv3["Title"] = fv3["Title"].apply(lambda x: trunc(x, 45))
    st.dataframe(
        sno(fv3[[c for c in FC_DISP_COLS if c in fv3.columns]]),
        use_container_width=True)


# ── TAB 4: CLUSTER VIEW ───────────────────────────────────────────────────────
with tabs[4]:
    st.subheader("🗺️ Cluster-Level View")
    st.dataframe(cl_sum, use_container_width=True, hide_index=True)
    st.divider()
    for _, row in cl_sum.iterrows():
        cl = row["FC Cluster"]
        with st.expander(
                f"📍 {cl}  —  Dispatch: {fmt(int(row['Dispatch Needed']))} units  "
                f"|  FCs: {row['FC Codes']}"):
            cd = fcp[fcp["FC Cluster"]==cl].copy()
            cd["Title"] = cd["Title"].apply(lambda x: trunc(x, 45))
            st.dataframe(
                cd[["FC Code","FC Name","FC City","FC State",
                    "MSKU","Title","FC Stock","Avg Daily Sale",
                    "FC DOC","FC Health","FC Dispatch","FC Urgency"]],
                use_container_width=True, hide_index=True)


# ── TAB 5: RISK ALERTS ────────────────────────────────────────────────────────
with tabs[5]:
    st.subheader("🚨 Risk Alerts")
    rt1,rt2,rt3,rt4 = st.tabs(
        ["🔴 Critical (<14d)","⚫ Dead Stock","🟡 Slow Moving (>90d)","🔵 Excess"])

    with rt1:
        st.markdown(f"**{len(r_crit)} SKU(s) — order immediately**")
        if not r_crit.empty:
            st.dataframe(risk_view(r_crit.sort_values("Days of Cover")),
                         use_container_width=True)
        else: st.success("✅ No critical SKUs!")

    with rt2:
        st.markdown(f"**{len(r_dead)} SKU(s) — stock exists, zero sales**")
        if not r_dead.empty:
            d2 = r_dead.copy()
            d2["Title"] = d2["Title"].apply(lambda x: trunc(x,50))
            cols2 = ["MSKU","Title","Channel","Total Sellable"]+FC_STOCK_COLS+["All-Time Sales"]
            st.dataframe(sno(d2[[c for c in cols2 if c in d2.columns]]),
                         use_container_width=True)
        else: st.success("✅ No dead stock.")

    with rt3:
        st.markdown(f"**{len(r_slow)} SKU(s) — days of cover > 90**")
        if not r_slow.empty:
            st.dataframe(risk_view(r_slow.sort_values("Days of Cover", ascending=False)),
                         use_container_width=True)
        else: st.success("✅ No slow-moving SKUs.")

    with rt4:
        st.markdown(f"**{len(r_excess)} SKU(s) — cover > {planning_days*2} days**")
        if not r_excess.empty:
            st.dataframe(risk_view(r_excess.sort_values("Days of Cover", ascending=False)),
                         use_container_width=True)
        else: st.success("✅ No excess stock.")


# ── TAB 6: SALES TRENDS ───────────────────────────────────────────────────────
with tabs[6]:
    st.subheader("📈 Sales Trends")

    tc1, tc2 = st.columns(2)
    with tc1:
        st.markdown("**Weekly Sales**")
        if not wk_tr.empty: st.line_chart(wk_tr.set_index("Week")["Units Sold"])
    with tc2:
        st.markdown("**Monthly Sales by Channel**")
        if not mo_tr.empty:
            pm = mo_tr.pivot_table(index="Month",columns="Channel",
                                    values="Units Sold",aggfunc="sum",fill_value=0)
            st.bar_chart(pm)

    st.markdown("**🔥 Top 20 SKUs by Avg Daily Sale — with FC-wise Stock**")
    t20 = r_top20.copy()
    t20["Title"] = t20["Title"].apply(lambda x: trunc(x,50))
    t20_cols = (["MSKU","Title","Avg Daily Sale","Total Sellable"]
                + FC_STOCK_COLS
                + ["Days of Cover","Velocity","Dispatch Needed"])
    st.dataframe(sno(t20[[c for c in t20_cols if c in t20.columns]]),
                 use_container_width=True)

    st.divider()
    st.markdown("**📊 FC-Wise Stock Velocity — which SKUs move fastest at each FC**")
    fvel = fc_velocity.copy()
    fvel["Title"] = fvel["Title"].apply(lambda x: trunc(x,45))
    st.dataframe(
        sno(fvel[["FC Code","FC Name","FC City","FC Cluster",
                  "MSKU","Title","FC Stock","Avg Daily Sale","Velocity"]]),
        use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("**🔍 SKU Drill-Down**")
    sel = st.selectbox("Select MSKU", sorted(sales["MSKU"].unique()))
    sku_d = (sales[sales["MSKU"]==sel].groupby("Date")["Qty"]
             .sum().reset_index().set_index("Date"))
    if not sku_d.empty:
        st.caption(f"**{trunc(title_map.get(sel,sel),80)}**")
        st.line_chart(sku_d["Qty"])
    sku_fc = fc_stock_long[fc_stock_long["MSKU"]==sel]
    if not sku_fc.empty:
        st.markdown(f"**FC Stock Breakdown for {sel}**")
        st.dataframe(
            sku_fc[["FC Code","FC Name","FC City","FC State","FC Cluster","FC Stock"]],
            use_container_width=True, hide_index=True)


# ── TAB 7: DAMAGED STOCK ──────────────────────────────────────────────────────
with tabs[7]:
    st.subheader("🔧 Damaged / Non-Sellable Stock")
    st.caption("Raise a Removal Order or Reimbursement Claim in Seller Central.")
    if dmg_detail.empty:
        st.success("✅ No damaged or non-sellable stock found.")
    else:
        dr = dmg_detail.copy()
        dr["Title"] = dr["Title"].apply(lambda x: trunc(x,55))
        st.dataframe(sno(dr.sort_values("Qty",ascending=False)), use_container_width=True)

        st.divider()
        st.markdown("**📋 Full Disposition Breakdown per MSKU**")
        db = disp_bkdn.copy()
        db.insert(1, "Title",
                  db["MSKU"].map(title_map).fillna("").apply(lambda x: trunc(x,55)))
        st.dataframe(sno(db), use_container_width=True)

        if not inv_dmg.empty:
            st.divider()
            st.markdown("**🏭 Damaged Stock by FC**")
            dmg_fc = (inv_dmg
                .groupby(["FC Code","FC Name","Disposition"])["Stock"]
                .sum().reset_index().rename(columns={"Stock":"Qty"})
                .sort_values("Qty", ascending=False))
            st.dataframe(sno(dmg_fc), use_container_width=True, hide_index=True)


# ── TAB 8: AMAZON FLAT FILE ───────────────────────────────────────────────────
with tabs[8]:
    st.subheader("📄 Amazon FBA Shipment Flat File")
    st.info(
        "**How to use:**  \n"
        "1. Review / edit quantities below  \n"
        "2. Select target FC  →  Click Download  \n"
        "3. Seller Central → **Send to Amazon → Create New Shipment → Upload a file** ✅  \n"
        "_Each FC needs a separate flat file._")

    ff1, ff2, ff3 = st.columns(3)
    with ff1:
        shp_label = st.text_input("Shipment Name",
            value=f"Replenishment_{datetime.now().strftime('%Y%m%d')}")
        shp_date  = st.date_input("Expected Ship Date",
            value=datetime.now().date()+timedelta(days=2))
    with ff2:
        tgt_fc = st.selectbox("Target FC", _active_fcs,
            format_func=lambda x: f"{x} — {fc_name(x)}")
    with ff3:
        st.markdown(f"**FC Code:** {tgt_fc}")
        st.markdown(f"**Name:** {fc_name(tgt_fc)}")
        st.markdown(f"**City:** {fc_city(tgt_fc)}  |  **State:** {fc_state(tgt_fc)}")
        st.markdown(f"**Cluster:** {fc_cluster(tgt_fc)}")

    flat_base = fba_plan[fba_plan["Dispatch Needed"]>0].copy()
    flat_base["Units to Ship"]  = flat_base["Dispatch Needed"].astype(int)
    flat_base["Units Per Case"] = case_qty
    flat_base["No. of Cases"]   = np.ceil(flat_base["Units to Ship"]/case_qty).astype(int)
    flat_base["Title Short"]    = flat_base["Title"].apply(lambda x: trunc(x,80))

    st.markdown(f"**{len(flat_base)} SKUs  |  {fmt(flat_base['Units to Ship'].sum())} total units**")
    edited = st.data_editor(
        flat_base[["MSKU","FNSKU","Title Short","Units to Ship",
                   "Units Per Case","No. of Cases"]].reset_index(drop=True),
        num_rows="dynamic", use_container_width=True)

    def make_flat_file(df, fc_code):
        parts = [p.strip() for p in ship_addr.split(",")]
        city  = parts[1] if len(parts)>1 else ""
        stz   = parts[2] if len(parts)>2 else ""
        stc   = stz.split("-")[0].strip()[:2].upper() if "-" in stz else stz[:2].upper()
        post  = stz.split("-")[1].strip() if "-" in stz else ""
        lines = ["TemplateType=FlatFileShipmentCreation\tVersion=2015.0403", ""]
        mh    = ["ShipmentName","ShipFromName","ShipFromAddressLine1","ShipFromCity",
                 "ShipFromStateOrProvinceCode","ShipFromPostalCode","ShipFromCountryCode",
                 "ShipmentStatus","LabelPrepType","AreCasesRequired",
                 "DestinationFulfillmentCenterId"]
        mv    = [shp_label, ship_name or "My Warehouse",
                 parts[0] if parts else "", city, stc, post,
                 "IN","WORKING",lbl_own,"YES" if case_packed else "NO",fc_code]
        lines += ["\t".join(mh), "\t".join(str(v) for v in mv), ""]
        lines.append("\t".join(["SellerSKU","FNSKU","QuantityShipped","QuantityInCase",
                                  "PrepOwner","LabelOwner","ItemDescription",
                                  "ExpectedDeliveryDate"]))
        for _, r in df.iterrows():
            qic = str(int(r.get("Units Per Case", case_qty))) if case_packed else ""
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
        f"Amazon_FBA_{tgt_fc}_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
        "text/plain")

    if len(_active_fcs)>1:
        st.divider()
        st.markdown("**📦 Download one file per FC**")
        bc = st.columns(min(len(_active_fcs), 4))
        for i, fc in enumerate(_active_fcs):
            with bc[i%4]:
                st.download_button(
                    f"📥 {fc}\n{fc_name(fc)}",
                    make_flat_file(edited, fc).encode("utf-8"),
                    f"Amazon_FBA_{fc}_{datetime.now().strftime('%Y%m%d')}.txt",
                    "text/plain", key=f"dl_{fc}")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 10 — EXCEL EXPORT  (17 FC-wise sheets)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")


def build_excel():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        # ── Plan sheets (all have FC stock columns) ──────────────────────────
        excel_sheet(w, plan_view(fba_plan.sort_values("Priority",ascending=False)),
                    "FBA Plan")
        excel_sheet(w,
                    plan_view(fbm_plan) if not fbm_plan.empty
                    else pd.DataFrame({"Note":["No FBM SKUs"]}),
                    "FBM Plan")
        excel_sheet(w, plan_view(plan.sort_values("Priority",ascending=False)),
                    "All SKUs — FC Stock")
        # ── Risk sheets (all have FC stock columns) ──────────────────────────
        excel_sheet(w, risk_view(r_crit.sort_values("Days of Cover")),
                    "CRITICAL — FC Stock")
        excel_sheet(w, risk_view(r_dead),
                    "Dead Stock — FC Stock")
        excel_sheet(w, risk_view(r_slow.sort_values("Days of Cover",ascending=False)),
                    "Slow Moving — FC Stock")
        excel_sheet(w, risk_view(r_excess.sort_values("Days of Cover",ascending=False)),
                    "Excess — FC Stock")
        excel_sheet(w,
                    sno(r_top20[[c for c in (["MSKU","Title","Avg Daily Sale","Total Sellable"]
                                              +FC_STOCK_COLS+["Days of Cover","Velocity","Dispatch Needed"])
                                 if c in r_top20.columns]]),
                    "Top 20 SKUs — FC Stock")
        # ── FC-wise sheets ───────────────────────────────────────────────────
        excel_sheet(w, sno(fc_disp_summary),            "FC Dispatch Summary")
        excel_sheet(w,
                    sno(fcp.sort_values(["FC Cluster","FC Dispatch"],
                                         ascending=[True,False])),
                    "FC Dispatch Detail")
        excel_sheet(w, sno(fc_inv_summary),              "FC Inventory Summary")
        excel_sheet(w, cl_sum,                           "Cluster Summary")
        excel_sheet(w,
                    sno(fc_velocity[[c for c in
                                      ["FC Code","FC Name","FC City","FC Cluster",
                                       "MSKU","Title","FC Stock","Avg Daily Sale","Velocity"]
                                      if c in fc_velocity.columns]]),
                    "FC Velocity")
        # ── Damaged / trends ─────────────────────────────────────────────────
        excel_sheet(w, sno(dmg_detail.sort_values("Qty",ascending=False)),
                    "Damaged Stock")
        excel_sheet(w, wk_tr,  "Weekly Trend")
        excel_sheet(w, mo_tr,  "Monthly Trend")
        # ── FC Master Reference ──────────────────────────────────────────────
        excel_sheet(w,
                    pd.DataFrame([{"FC Code":k,"FC Name":v[0],"City":v[1],
                                    "Area":v[2],"State":v[3],"Cluster":v[4]}
                                   for k,v in FC_DATA.items()]),
                    "FC Master Reference")
    buf.seek(0)
    return buf


dl1, dl2 = st.columns(2)
with dl1:
    st.download_button(
        "📥 Download Full Report — Excel (17 FC-wise sheets)",
        data=build_excel(),
        file_name=f"FBA_Supply_Plan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with dl2:
    csv_df = fba_plan[fba_plan["Dispatch Needed"]>0].copy()
    csv_df["Title"] = csv_df["Title"].apply(lambda x: trunc(x,80))
    csv_cols = [c for c in PLAN_COLS if c in csv_df.columns]
    st.download_button(
        "📋 Download Dispatch Plan — CSV (FC-wise)",
        data=csv_df[csv_cols].to_csv(index=False),
        file_name=f"Dispatch_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv")
