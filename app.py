"""
╔══════════════════════════════════════════════════════════╗
║   FBA Smart Supply Planner  ·  Amazon India Edition      ║
║   Professional Inventory & Dispatch Planning Tool        ║
╚══════════════════════════════════════════════════════════╝

Upload:  Inventory Ledger (required)  +  MTR (optional)
Sales  : Customer Shipments column — the only correct source
Stock  : Ending Warehouse Balance — latest row per SKU/FC (daily dedup)
Output : Full plan · FC dispatch · Bulk shipment file · Excel/CSV reports
"""

import io, math, zipfile
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE CONFIG & GLOBAL CSS
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="FBA Supply Planner — Amazon India",
    layout="wide",
    page_icon="📦",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
/* Metrics */
[data-testid="stMetricValue"]  {font-size:1.7rem;font-weight:800;color:#1a202c}
[data-testid="stMetricLabel"]  {font-size:.78rem;font-weight:600;color:#718096;text-transform:uppercase;letter-spacing:.05em}
[data-testid="stMetricDelta"]  {font-size:.82rem}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {gap:4px;background:#f7fafc;padding:4px;border-radius:10px}
.stTabs [data-baseweb="tab"]      {font-size:13px;font-weight:600;border-radius:8px;padding:6px 16px}
.stTabs [aria-selected="true"]    {background:#2b6cb0!important;color:white!important}

/* Alert banners */
.alert{padding:10px 16px;border-radius:8px;font-weight:600;font-size:.9rem;
       margin:4px 0;line-height:1.5}
.a-red   {background:#fff5f5;border-left:5px solid #e53e3e;color:#742a2a}
.a-amber {background:#fffaf0;border-left:5px solid #ed8936;color:#7b341e}
.a-green {background:#f0fff4;border-left:5px solid #38a169;color:#1c4532}
.a-blue  {background:#ebf8ff;border-left:5px solid #3182ce;color:#1a365d}
.a-grey  {background:#f7fafc;border-left:5px solid #718096;color:#2d3748}

/* Section headers */
.sec-hdr{font-size:1.05rem;font-weight:700;color:#2d3748;padding:4px 0 2px 0;
         border-bottom:2px solid #e2e8f0;margin-bottom:8px}

/* Status pills inline */
.pill{display:inline-block;padding:1px 8px;border-radius:10px;font-size:.78rem;font-weight:700}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  FC MASTER  —  75 verified Amazon India fulfillment centres
#  Tuple: (FC Name, City, State, Cluster)
#  NOTE: PNQ2 = New Delhi (Mohan Co-op A-33), NOT Pune
# ══════════════════════════════════════════════════════════════════════════════
FC_MASTER = {
    "SGAA":      ("Guwahati FC",               "Guwahati",        "ASSAM",           "East Assam"),
    "DEX3":      ("New Delhi FC A28",           "New Delhi",       "DELHI",           "Delhi NCR"),
    "DEX8":      ("New Delhi FC A29",           "New Delhi",       "DELHI",           "Delhi NCR"),
    "PNQ2":      ("New Delhi FC A33",           "New Delhi",       "DELHI",           "Delhi NCR"),
    "XDEL":      ("Delhi XL FC",                "New Delhi",       "DELHI",           "Delhi NCR"),
    "DEL2":      ("Tauru FC",                   "Mewat",           "HARYANA",         "Delhi NCR"),
    "DEL4":      ("Gurgaon FC",                 "Gurgaon",         "HARYANA",         "Delhi NCR"),
    "DEL5":      ("Manesar FC",                 "Manesar",         "HARYANA",         "Delhi NCR"),
    "DEL6":      ("Bilaspur FC",                "Bilaspur",        "HARYANA",         "Delhi NCR"),
    "DEL7":      ("Bilaspur FC 2",              "Bilaspur",        "HARYANA",         "Delhi NCR"),
    "DEL8":      ("Sohna FC",                   "Sohna",           "HARYANA",         "Delhi NCR"),
    "DEL8_DED5": ("Sohna ESR FC",               "Sohna",           "HARYANA",         "Delhi NCR"),
    "DED3":      ("Farrukhnagar FC",            "Farrukhnagar",    "HARYANA",         "Delhi NCR"),
    "DED4":      ("Sohna FC 2",                 "Sohna",           "HARYANA",         "Delhi NCR"),
    "DED5":      ("Sohna FC 3",                 "Sohna",           "HARYANA",         "Delhi NCR"),
    "AMD1":      ("Ahmedabad Naroda FC",        "Naroda",          "GUJARAT",         "Gujarat West"),
    "AMD2":      ("Changodar FC",               "Changodar",       "GUJARAT",         "Gujarat West"),
    "SUB1":      ("Surat FC",                   "Surat",           "GUJARAT",         "Gujarat West"),
    "BLR4":      ("Devanahalli FC",             "Devanahalli",     "KARNATAKA",       "Bangalore"),
    "BLR5":      ("Bommasandra FC",             "Bommasandra",     "KARNATAKA",       "Bangalore"),
    "BLR6":      ("Nelamangala FC",             "Nelamangala",     "KARNATAKA",       "Bangalore"),
    "BLR7":      ("Hoskote FC",                 "Hoskote",         "KARNATAKA",       "Bangalore"),
    "BLR8":      ("Devanahalli FC 2",           "Devanahalli",     "KARNATAKA",       "Bangalore"),
    "BLR10":     ("Kudlu Gate FC",              "Bengaluru",       "KARNATAKA",       "Bangalore"),
    "BLR12":     ("Attibele FC",                "Attibele",        "KARNATAKA",       "Bangalore"),
    "BLR13":     ("Jigani FC",                  "Jigani",          "KARNATAKA",       "Bangalore"),
    "BLR14":     ("Anekal FC",                  "Anekal",          "KARNATAKA",       "Bangalore"),
    "SCJA":      ("Bangalore SCJA FC",          "Bengaluru",       "KARNATAKA",       "Bangalore"),
    "XSAB":      ("Bangalore XS FC",            "Bengaluru",       "KARNATAKA",       "Bangalore"),
    "SBLA":      ("Bangalore SBLA FC",          "Bengaluru",       "KARNATAKA",       "Bangalore"),
    "SIDA":      ("Indore FC",                  "Indore",          "MADHYA PRADESH",  "Central"),
    "FBHB":      ("Bhopal FC",                  "Bhopal",          "MADHYA PRADESH",  "Central"),
    "FIDA":      ("Bhopal FC 2",                "Bhopal",          "MADHYA PRADESH",  "Central"),
    "IND1":      ("Indore FC 2",                "Indore",          "MADHYA PRADESH",  "Central"),
    "BOM1":      ("Bhiwandi FC",                "Bhiwandi",        "MAHARASHTRA",     "Mumbai West"),
    "BOM3":      ("Nashik FC",                  "Nashik",          "MAHARASHTRA",     "Mumbai West"),
    "BOM4":      ("Vashere FC",                 "Bhiwandi",        "MAHARASHTRA",     "Mumbai West"),
    "BOM5":      ("Bhiwandi 2 FC",              "Bhiwandi",        "MAHARASHTRA",     "Mumbai West"),
    "BOM7":      ("Bhiwandi 3 FC",              "Bhiwandi",        "MAHARASHTRA",     "Mumbai West"),
    "ISK3":      ("Bhiwandi ISK3 FC",           "Bhiwandi",        "MAHARASHTRA",     "Mumbai West"),
    "PNQ1":      ("Pune Hinjewadi FC",          "Pune",            "MAHARASHTRA",     "Mumbai West"),
    "PNQ3":      ("Pune Chakan FC",             "Pune",            "MAHARASHTRA",     "Mumbai West"),
    "SAMB":      ("Mumbai SAMB FC",             "Mumbai",          "MAHARASHTRA",     "Mumbai West"),
    "XBOM":      ("Mumbai XL FC",               "Bhiwandi",        "MAHARASHTRA",     "Mumbai West"),
    "ATX1":      ("Ludhiana FC",                "Ludhiana",        "PUNJAB",          "North Punjab"),
    "LDH1":      ("Ludhiana FC 2",              "Ludhiana",        "PUNJAB",          "North Punjab"),
    "RAJ1":      ("Rajpura FC",                 "Rajpura",         "PUNJAB",          "North Punjab"),
    "JAI1":      ("Jaipur FC",                  "Jaipur",          "RAJASTHAN",       "North Rajasthan"),
    "JPX1":      ("Jaipur JPX1 FC",             "Jaipur",          "RAJASTHAN",       "North Rajasthan"),
    "JPX2":      ("Bagru FC",                   "Bagru",           "RAJASTHAN",       "North Rajasthan"),
    "MAA1":      ("Irungattukottai FC",         "Chennai",         "TAMIL NADU",      "Chennai"),
    "MAA2":      ("Ponneri FC",                 "Ponneri",         "TAMIL NADU",      "Chennai"),
    "MAA3":      ("Sriperumbudur FC",           "Sriperumbudur",   "TAMIL NADU",      "Chennai"),
    "MAA4":      ("Ambattur FC",                "Ambattur",        "TAMIL NADU",      "Chennai"),
    "MAA5":      ("Kanchipuram FC",             "Kanchipuram",     "TAMIL NADU",      "Chennai"),
    "CJB1":      ("Coimbatore FC",              "Coimbatore",      "TAMIL NADU",      "Chennai"),
    "SMAB":      ("Chennai SMAB FC",            "Chennai",         "TAMIL NADU",      "Chennai"),
    "COK1":      ("Kochi FC",                   "Kochi",           "KERALA",          "South Kerala"),
    "HYD3":      ("Shamshabad FC",              "Hyderabad",       "TELANGANA",       "Hyderabad"),
    "HYD6":      ("Kothur FC",                  "Kothur",          "TELANGANA",       "Hyderabad"),
    "HYD7":      ("Medchal FC",                 "Medchal",         "TELANGANA",       "Hyderabad"),
    "HYD8":      ("Shamshabad FC 2",            "Hyderabad",       "TELANGANA",       "Hyderabad"),
    "HYD8_HYD3": ("Shamshabad Mamidipally FC",  "Hyderabad",       "TELANGANA",       "Hyderabad"),
    "HYD9":      ("Pedda Amberpet FC",          "Hyderabad",       "TELANGANA",       "Hyderabad"),
    "HYD10":     ("Ghatkesar FC",               "Ghatkesar",       "TELANGANA",       "Hyderabad"),
    "HYD11":     ("Kompally FC",                "Kompally",        "TELANGANA",       "Hyderabad"),
    "LKO1":      ("Lucknow FC",                 "Lucknow",         "UTTAR PRADESH",   "North UP"),
    "AGR1":      ("Agra FC",                    "Agra",            "UTTAR PRADESH",   "North UP"),
    "SLDK":      ("Kishanpur FC",               "Bijnour",         "UTTAR PRADESH",   "North UP"),
    "CCU1":      ("Kolkata FC",                 "Kolkata",         "WEST BENGAL",     "Kolkata East"),
    "CCU2":      ("Kolkata FC 2",               "Kolkata",         "WEST BENGAL",     "Kolkata East"),
    "CCX1":      ("Howrah FC",                  "Howrah",          "WEST BENGAL",     "Kolkata East"),
    "CCX2":      ("Howrah FC 2",                "Howrah",          "WEST BENGAL",     "Kolkata East"),
    "PAT1":      ("Patna FC",                   "Patna",           "BIHAR",           "Kolkata East"),
    "BBS1":      ("Bhubaneswar FC",             "Bhubaneswar",     "ODISHA",          "Kolkata East"),
}

def fc_name(c):    return FC_MASTER.get(str(c).upper(), (str(c),))[0]
def fc_city(c):    return FC_MASTER.get(str(c).upper(), ("","Unknown"))[1]
def fc_state(c):   return FC_MASTER.get(str(c).upper(), ("","","Unknown"))[2]
def fc_cluster(c): return FC_MASTER.get(str(c).upper(), ("","","","Other"))[3]


# ══════════════════════════════════════════════════════════════════════════════
#  UTILITY FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════
def read_file(f):
    """Read CSV / ZIP / XLSX → DataFrame. Returns empty DF on failure."""
    try:
        n = f.name.lower()
        if n.endswith(".zip"):
            with zipfile.ZipFile(f) as z:
                for m in z.namelist():
                    if m.lower().endswith(".csv"):
                        return pd.read_csv(z.open(m), low_memory=False)
            return pd.DataFrame()
        if n.endswith((".xlsx", ".xls")):
            return pd.read_excel(f)
        return pd.read_csv(f, low_memory=False)
    except Exception as e:
        st.warning(f"Could not read {f.name}: {e}")
        return pd.DataFrame()


def get_col(df, candidates):
    """Return first matching column name (case-insensitive). '' if none."""
    low = {c.strip().lower(): c for c in df.columns}
    for name in candidates:
        if name.strip().lower() in low:
            return low[name.strip().lower()]
    return ""


def parse_dates(series):
    """Try MM/DD/YYYY first (Amazon default), then DD/MM/YYYY as fallback."""
    out = pd.to_datetime(series, errors="coerce", dayfirst=False)
    mask = out.isna()
    if mask.any():
        out.loc[mask] = pd.to_datetime(series.loc[mask], errors="coerce", dayfirst=True)
    return out


def trunc(s, n=60):
    s = str(s)
    return s[:n] + "…" if len(s) > n else s


def fmtn(x):
    try:    return f"{int(x):,}"
    except: return "—"


def health_label(doc, plan_days):
    if doc <= 0:           return "⚫ No Stock"
    if doc < 14:           return "🔴 Critical"
    if doc < 30:           return "🟠 Low"
    if doc < plan_days:    return "🟢 Healthy"
    if doc < plan_days*2:  return "🟡 Excess"
    return "🔵 Overstocked"


def velocity_label(avg):
    if avg <= 0:   return "⚫ Dead"
    if avg < 0.5:  return "🔵 Slow"
    if avg < 2:    return "🟡 Medium"
    if avg < 5:    return "🟢 Fast"
    return "🔥 Hot Seller"


def alert(msg, kind="green"):
    cls = {"green":"a-green","red":"a-red","amber":"a-amber","blue":"a-blue","grey":"a-grey"}[kind]
    st.markdown(f'<div class="alert {cls}">{msg}</div>', unsafe_allow_html=True)


def section(title):
    st.markdown(f'<div class="sec-hdr">{title}</div>', unsafe_allow_html=True)


def show_df(df, height=420):
    st.dataframe(df.reset_index(drop=True), use_container_width=True,
                 hide_index=True, height=height)


def xl_write(writer, df, sheet_name, freeze_col=0):
    """Write a styled Excel sheet with header formatting."""
    if df is None or df.empty:
        pd.DataFrame({"(No data)": []}).to_excel(writer, sheet_name=sheet_name, index=False)
        return
    d = df.copy()
    # Truncate long text for Excel
    for c in d.select_dtypes("object").columns:
        d[c] = d[c].astype(str).str[:120]
    d.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
    wb  = writer.book
    ws  = writer.sheets[sheet_name]
    hdr = wb.add_format({
        "bold": True, "bg_color": "#1a365d", "font_color": "white",
        "border": 1, "align": "center", "valign": "vcenter",
        "text_wrap": True, "font_size": 10,
    })
    num = wb.add_format({"num_format": "#,##0", "align": "right"})
    dec = wb.add_format({"num_format": "0.000",  "align": "right"})
    for ci, cn in enumerate(d.columns):
        ws.write(0, ci, cn, hdr)
        col_w = max(len(str(cn)), d[cn].astype(str).str.len().max() if not d.empty else 8)
        ws.set_column(ci, ci, min(col_w + 2, 48))
    ws.freeze_panes(2, freeze_col)
    ws.autofilter(0, 0, len(d), len(d.columns) - 1)
    ws.set_row(0, 22)


# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR — PLANNING SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Planning Settings")

    plan_days = st.number_input("Target Coverage Days", 7, 180, 60, step=1,
                                help="How many days of stock you want to maintain")
    svc_level = st.selectbox("Service Level (Safety Stock)", ["95%", "90%", "98%"],
                             help="Higher = more safety buffer")
    Z_SCORE   = {"90%": 1.28, "95%": 1.65, "98%": 2.05}[svc_level]
    sales_win = st.selectbox("Sales Window for Avg Daily Sale",
                             ["Full History", "Last 30 Days", "Last 60 Days", "Last 90 Days"])
    min_units = st.number_input("Min Dispatch Units to Show", 0, 9999, 0)

    st.markdown("---")
    st.markdown("## 🏭 Your Warehouse")
    wh_name  = st.text_input("Company / Warehouse Name", placeholder="GLOOYA Warehouse")
    wh_addr  = st.text_input("Address Line 1",           placeholder="Plot 12, Industrial Area")
    wh_city  = st.text_input("City",                     placeholder="Lucknow")
    wh_state = st.text_input("State Code (2 letters)",   placeholder="UP", max_chars=2).upper()
    wh_pin   = st.text_input("PIN Code",                 placeholder="226001")

    st.markdown("---")
    st.markdown("## 📦 Shipment Defaults")
    case_qty    = st.number_input("Units Per Carton", 1, 500, 12)
    case_packed = st.checkbox("Case-Packed Shipment")
    prep_own    = st.selectbox("Prep Owner",  ["AMAZON", "SELLER"])
    lbl_own     = st.selectbox("Label Owner", ["AMAZON", "SELLER"])


# ══════════════════════════════════════════════════════════════════════════════
#  HEADER
# ══════════════════════════════════════════════════════════════════════════════
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown("# 📦 FBA Supply Planner")
    st.markdown("**Amazon India** · Inventory Ledger + MTR → Full dispatch plan, dead stock view & bulk shipment file")
with col_h2:
    st.markdown(f"<div style='text-align:right;color:#718096;font-size:.85rem;padding-top:20px'>"
                f"Run date: {datetime.now().strftime('%d %b %Y %H:%M')}</div>", unsafe_allow_html=True)

st.markdown("---")


# ══════════════════════════════════════════════════════════════════════════════
#  FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("### 📂 Step 1 — Upload Your Reports")
st.caption("You can select **multiple files** at once (hold Ctrl / Cmd while clicking, or drag-drop several files)")

u1, u2 = st.columns(2)
with u1:
    inv_files = st.file_uploader(
        "📊 **Inventory Ledger** — Required",
        type=["csv", "zip", "xlsx"],
        accept_multiple_files=True,
        key="inv_up",
        help="Download from: Seller Central → Reports → Fulfillment → Inventory Ledger")
with u2:
    mtr_files = st.file_uploader(
        "📋 **MTR Sales Report** — Optional",
        type=["csv", "zip", "xlsx"],
        accept_multiple_files=True,
        key="mtr_up",
        help="Download from: Seller Central → Reports → Tax → MTR (adds FBM data)")

with st.expander("📖 File format guide"):
    st.markdown("""
**Inventory Ledger** *(Seller Central → Reports → Fulfillment → Inventory Ledger)*

| Column | Example | Used for |
|---|---|---|
| `Date` | 02/12/2026 | Daily dedup — keep latest snapshot per SKU/FC |
| `MSKU` | GL-CHOKER | SKU identifier |
| `FNSKU` | X00283UABJ | Amazon barcode for flat file |
| `Title` | GLOOYA Choker... | Product name |
| `Disposition` | SELLABLE | Split sellable vs damaged |
| `Ending Warehouse Balance` | 23 | ✅ Correct stock number |
| `Customer Shipments` | -5 | ✅ Correct sales (negative = sold) |
| `Customer Returns` | 0 | Subtracted from shipments |
| `Location` | LKO1 | FC code |

> **Why Customer Shipments is the correct sales source:** The Ending Warehouse Balance is calculated
> from all these movement columns. Customer Shipments captures exactly what was dispatched each day.
""")

if not inv_files:
    alert("👆 Upload your <b>Inventory Ledger</b> file above to begin.", "blue")
    with st.expander("🗺️ All 75 Amazon India FC Codes"):
        fc_ref = pd.DataFrame([{"FC Code": k, "FC Name": v[0], "City": v[1],
                                 "State": v[2], "Cluster": v[3]}
                                for k, v in FC_MASTER.items()])
        show_df(fc_ref, height=500)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 1 — PARSE & MERGE INVENTORY LEDGER
# ══════════════════════════════════════════════════════════════════════════════
parts = [read_file(f) for f in inv_files]
parts = [d for d in parts if not d.empty]
if not parts:
    st.error("❌ Could not read any Inventory Ledger file. Check format."); st.stop()

for i in range(len(parts)):
    parts[i].columns = parts[i].columns.str.strip()
raw = pd.concat(parts, ignore_index=True)

# Detect columns
CD    = get_col(raw, ["Date"])
CMSKU = get_col(raw, ["MSKU", "Sku", "SKU"])
CFNSK = get_col(raw, ["FNSKU"])
CASIN = get_col(raw, ["ASIN"])
CTITL = get_col(raw, ["Title", "Product Name", "Description"])
CDISP = get_col(raw, ["Disposition"])
CEND  = get_col(raw, ["Ending Warehouse Balance", "Quantity", "Qty"])
CLOC  = get_col(raw, ["Location", "FC Code", "Warehouse Code"])
CSHIP = get_col(raw, ["Customer Shipments"])
CRET  = get_col(raw, ["Customer Returns"])

for lbl, c in [("MSKU", CMSKU), ("Ending Warehouse Balance", CEND), ("Location", CLOC)]:
    if not c:
        st.error(f"❌ Required column not found: **{lbl}**")
        st.write("Columns in your file:", list(raw.columns))
        st.stop()

# Build clean frame — use consistent internal column names from here on
inv = pd.DataFrame()
inv["MSKU"]   = raw[CMSKU].astype(str).str.strip()
inv["FNSKU"]  = raw[CFNSK].astype(str).str.strip()  if CFNSK else ""
inv["ASIN"]   = raw[CASIN].astype(str).str.strip()  if CASIN else ""
inv["Title"]  = raw[CTITL].astype(str).str.strip()  if CTITL else ""
inv["Disp"]   = raw[CDISP].astype(str).str.upper().str.strip() if CDISP else "SELLABLE"
inv["Stock"]  = pd.to_numeric(raw[CEND],  errors="coerce").fillna(0)
inv["FC Code"]= raw[CLOC].astype(str).str.upper().str.strip()

n_raw = len(inv)

# ── DAILY DEDUP ───────────────────────────────────────────────────────────────
# The ledger has 1 row per SKU/FC/Disposition per day.
# Stock = Ending Warehouse Balance of the LATEST date only.
# Summing all rows = Stock × number of days → WRONG.
if CD:
    inv["_dt"] = parse_dates(raw[CD].astype(str))
    last_date  = inv["_dt"].max()
    inv = (inv.sort_values("_dt")
              .groupby(["MSKU", "Disp", "FC Code"], as_index=False)
              .last()
              .drop(columns=["_dt"]))
    date_label = last_date.strftime("%d %b %Y") if pd.notna(last_date) else "Latest"
else:
    inv = inv.groupby(["MSKU", "Disp", "FC Code"], as_index=False).last()
    date_label = "Latest row"

n_dedup = len(inv)

# Enrich with FC metadata
inv["FC Name"]    = inv["FC Code"].apply(fc_name)
inv["FC City"]    = inv["FC Code"].apply(fc_city)
inv["FC State"]   = inv["FC Code"].apply(fc_state)
inv["FC Cluster"] = inv["FC Code"].apply(fc_cluster)

# ── SKU MASTER (MSKU → FNSKU, ASIN, Title) ───────────────────────────────────
sku_df = (inv[["MSKU","FNSKU","ASIN","Title"]]
    .sort_values("MSKU")
    .drop_duplicates("MSKU", keep="last")
    .set_index("MSKU"))

# ── SPLIT SELLABLE / DAMAGED ──────────────────────────────────────────────────
sell  = inv[inv["Disp"] == "SELLABLE"].copy()
dmg   = inv[inv["Disp"] != "SELLABLE"].copy()

active_fcs = sorted(sell["FC Code"].unique())

# Per-SKU totals
sku_sell = (sell.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock": "Sellable Stock"}))
sku_dmg  = (dmg.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock": "Damaged Stock"}))

# FC-level stock (long format) — SINGLE SOURCE OF TRUTH, column always "FC Code"
fc_long = (sell
    .groupby(["FC Code","FC Name","FC City","FC State","FC Cluster","MSKU"])["Stock"]
    .sum().reset_index()
    .rename(columns={"Stock": "FC Stock"}))

# FC inventory summary (1 row per FC)
fc_inv_summary = (fc_long
    .groupby(["FC Code","FC Name","FC City","FC State","FC Cluster"])["FC Stock"]
    .sum().reset_index().rename(columns={"FC Stock": "Total Sellable"})
    .sort_values("Total Sellable", ascending=False))

# Damaged by MSKU + FC
dmg_by_msku = (dmg.groupby(["MSKU","Disp"])["Stock"]
    .sum().reset_index().rename(columns={"Disp": "Disposition","Stock": "Qty"}))
dmg_by_msku["Product Name"] = (dmg_by_msku["MSKU"]
    .map(sku_df["Title"]).fillna("").apply(lambda x: trunc(x, 70)))

dmg_by_fc = (dmg.groupby(["FC Code","FC Name","FC City","MSKU","Disp"])["Stock"]
    .sum().reset_index().rename(columns={"Disp": "Disposition","Stock": "Qty"}))
dmg_by_fc["Product Name"] = (dmg_by_fc["MSKU"]
    .map(sku_df["Title"]).fillna("").apply(lambda x: trunc(x, 60)))


# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 2 — PARSE SALES DATA
#  PRIMARY  : Customer Shipments column (Inventory Ledger) — always correct
#  SECONDARY: MTR file — adds FBM SKUs or fills gaps
# ══════════════════════════════════════════════════════════════════════════════
ledger_sales = pd.DataFrame()
if CSHIP and CD:
    ls = pd.DataFrame()
    ls["MSKU"]    = raw[CMSKU].astype(str).str.strip()
    ls["Date"]    = parse_dates(raw[CD].astype(str))
    ls["Shipped"] = pd.to_numeric(raw[CSHIP], errors="coerce").fillna(0).abs()
    ls["Returns"] = (pd.to_numeric(raw[CRET],  errors="coerce").fillna(0).clip(lower=0)
                     if CRET else pd.Series(0, index=raw.index))
    ls["Qty"]     = (ls["Shipped"] - ls["Returns"]).clip(lower=0)
    ls["Channel"] = "FBA"
    ledger_sales  = (ls[ls["Qty"] > 0][["MSKU","Date","Qty","Channel"]]
                     .dropna(subset=["Date"]).copy())

mtr_sales = pd.DataFrame()
if mtr_files:
    mtr_parts = [read_file(f) for f in mtr_files]
    mtr_parts = [d for d in mtr_parts if not d.empty]
    if mtr_parts:
        for i in range(len(mtr_parts)):
            mtr_parts[i].columns = mtr_parts[i].columns.str.strip()
        mr = pd.concat(mtr_parts, ignore_index=True)
        CS  = get_col(mr, ["Sku","SKU","MSKU","Item SKU"])
        CQ  = get_col(mr, ["Quantity","Qty","Units"])
        CDT = get_col(mr, ["Shipment Date","Purchase Date","Order Date","Date"])
        CF  = get_col(mr, ["Fulfilment","Fulfillment","Fulfilment Channel","Fulfillment Channel"])
        if CS and CQ and CDT:
            FTM = {"AFN":"FBA","FBA":"FBA","AMAZON_FULFILLED":"FBA",
                   "MFN":"FBM","FBM":"FBM","MERCHANT_FULFILLED":"FBM"}
            ms = pd.DataFrame()
            ms["MSKU"]    = mr[CS].astype(str).str.strip()
            ms["Qty"]     = pd.to_numeric(mr[CQ], errors="coerce").fillna(0)
            ms["Date"]    = parse_dates(mr[CDT].astype(str))
            ms["Channel"] = (mr[CF].astype(str).str.upper().map(FTM).fillna("FBA")
                             if CF else "FBA")
            mtr_sales = ms[(ms["Qty"]>0)&ms["Date"].notna()][["MSKU","Date","Qty","Channel"]].copy()

# Merge sales sources
if not ledger_sales.empty and not mtr_sales.empty:
    extra = mtr_sales[~mtr_sales["MSKU"].isin(set(ledger_sales["MSKU"]))]
    sales = pd.concat([ledger_sales, extra], ignore_index=True)
    sales_src = "Inventory Ledger (Customer Shipments) + MTR"
elif not ledger_sales.empty:
    sales = ledger_sales.copy()
    sales_src = "Inventory Ledger — Customer Shipments"
elif not mtr_sales.empty:
    sales = mtr_sales.copy()
    sales_src = "MTR only"
else:
    st.error("❌ No sales data found. Ensure your Inventory Ledger has a **Customer Shipments** column.")
    st.stop()

sales = sales.sort_values("Date").reset_index(drop=True)

# Date window
s_min    = sales["Date"].min()
s_max    = sales["Date"].max()
all_days = max((s_max - s_min).days + 1, 1)
win_days = (30 if "30" in sales_win else 60 if "60" in sales_win else
            90 if "90" in sales_win else all_days)
cutoff   = s_max - pd.Timedelta(days=win_days - 1)
hist     = sales[sales["Date"] >= cutoff].copy() if win_days < all_days else sales.copy()
h_days   = max((hist["Date"].max() - hist["Date"].min()).days + 1, 1)

# Aggregates
hist_agg  = (hist.groupby(["MSKU","Channel"])["Qty"].sum()
             .reset_index().rename(columns={"Qty": "Sales"}))
full_agg  = (sales.groupby(["MSKU","Channel"])["Qty"].sum()
             .reset_index().rename(columns={"Qty": "All Time Sales"}))
daily_std = (hist.groupby(["MSKU","Channel","Date"])["Qty"].sum()
             .reset_index()
             .groupby(["MSKU","Channel"])["Qty"].std()
             .reset_index().rename(columns={"Qty": "Std Dev"}))

# Trend tables
mo_trend = (sales.assign(Month=sales["Date"].dt.to_period("M").astype(str))
    .groupby(["Month","Channel"])["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units"}).sort_values("Month"))
wk_trend = (sales.assign(Week=sales["Date"].dt.to_period("W").astype(str))
    .groupby("Week")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units"}).sort_values("Week"))


# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 3 — BUILD MASTER PLAN TABLE
#  Every SKU that has stock OR sales gets one row.
#  Columns: MSKU | FNSKU | Product Name | Sellable Stock | Damaged Stock |
#           Sales | All Time Sales | Avg/Day | Safety Stock | Need |
#           Dispatch | DOC | Status | Velocity | Priority
# ══════════════════════════════════════════════════════════════════════════════
keys = pd.concat([
    hist_agg[["MSKU","Channel"]],
    sku_sell.assign(Channel="FBA")[["MSKU","Channel"]],
], ignore_index=True).drop_duplicates()

plan = keys.copy()
for df_, on_ in [
    (hist_agg,   ["MSKU","Channel"]),
    (full_agg,   ["MSKU","Channel"]),
    (daily_std,  ["MSKU","Channel"]),
    (sku_sell,    "MSKU"),
    (sku_dmg,     "MSKU"),
]:
    plan = plan.merge(df_, on=on_, how="left")

for c in ["Sales","All Time Sales","Sellable Stock","Damaged Stock","Std Dev"]:
    plan[c] = pd.to_numeric(plan.get(c), errors="coerce").fillna(0)

# Add product info
plan["Product Name"] = plan["MSKU"].map(sku_df["Title"]).fillna("—")
plan["FNSKU"]        = plan["MSKU"].map(sku_df["FNSKU"]).fillna("")
plan["ASIN"]         = plan["MSKU"].map(sku_df["ASIN"]).fillna("")

# Planning calculations
plan["Avg/Day"]      = (plan["Sales"] / h_days).round(4)
plan["Safety Stock"] = (Z_SCORE * plan["Std Dev"] * math.sqrt(plan_days)).round(0)
plan["Need"]         = (plan["Avg/Day"] * plan_days + plan["Safety Stock"]).round(0)
plan["Dispatch"]     = (plan["Need"] - plan["Sellable Stock"]).clip(lower=0).round(0)
plan["DOC"]          = np.where(
    plan["Avg/Day"] > 0,
    (plan["Sellable Stock"] / plan["Avg/Day"]).round(1),
    np.where(plan["Sellable Stock"] > 0, 9999, 0)
)
plan["Status"]   = plan.apply(lambda r: health_label(r["DOC"], plan_days), axis=1)
plan["Velocity"] = plan["Avg/Day"].apply(velocity_label)

_mx = plan["Avg/Day"].max() or 1
plan["Priority"] = plan.apply(lambda r: round(
    (max(0, (plan_days - min(r["DOC"], plan_days)) / plan_days) * 0.65
     + min(r["Avg/Day"] / _mx, 1) * 0.35) * 100, 1)
    if r["Avg/Day"] > 0 else 0, axis=1)

fba_plan = plan[plan["Channel"] == "FBA"].copy()
fbm_plan = plan[plan["Channel"] == "FBM"].copy()

# Display column order — MSKU and Product Name always first
DISP_COLS = ["MSKU","FNSKU","Product Name","Sellable Stock","Damaged Stock",
             "Avg/Day","DOC","Status","Velocity","Sales","All Time Sales",
             "Need","Dispatch","Priority"]


def show_plan_table(df, min_disp=0, height=480):
    """Render a plan DataFrame with consistent column set."""
    d = df.copy()
    if min_disp > 0:
        d = d[d["Dispatch"] >= min_disp]
    d = d.sort_values("Priority", ascending=False)
    d["Product Name"] = d["Product Name"].apply(lambda x: trunc(x, 58))
    cols = [c for c in DISP_COLS if c in d.columns]
    show_df(d[cols], height=height)


# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 4 — FC-WISE DISPATCH ALLOCATION
#  Each SKU's total Dispatch is split across FCs proportional to current stock.
#  Internal column name is always "FC Code" — no renaming mid-flight.
# ══════════════════════════════════════════════════════════════════════════════
# Join fc_long (has "FC Code") with plan metrics
fcp = fc_long.merge(
    plan[["MSKU","Channel","Avg/Day","Dispatch","Sellable Stock",
          "Need","Product Name","FNSKU","ASIN","Priority","DOC"]],
    on="MSKU", how="left"
)

# Fill any unmapped plan columns
for c in ["Avg/Day","Dispatch","Sellable Stock","Need","Priority","DOC"]:
    fcp[c] = pd.to_numeric(fcp.get(c), errors="coerce").fillna(0)

fcp["Channel"] = fcp["Channel"].fillna("FBA")

# FC-level days of cover
fcp["FC DOC"] = np.where(
    fcp["Avg/Day"] > 0,
    (fcp["FC Stock"] / fcp["Avg/Day"]).round(1),
    np.where(fcp["FC Stock"] > 0, 9999, 0)
)
fcp["FC Status"] = fcp.apply(lambda r: health_label(r["FC DOC"], plan_days), axis=1)

# Proportional dispatch allocation
_tot = (fcp.groupby(["MSKU","Channel"])["FC Stock"]
        .sum().reset_index().rename(columns={"FC Stock": "_T"}))
fcp  = fcp.merge(_tot, on=["MSKU","Channel"], how="left")
fcp["FC Share"]    = (fcp["FC Stock"] / fcp["_T"].replace(0, 1)).fillna(1.0)
fcp["FC Dispatch"] = (fcp["Dispatch"] * fcp["FC Share"]).round(0).astype(int)
fcp.drop(columns=["_T"], inplace=True)

# FC dispatch summary (1 row per FC, all channels)
fc_disp_sum = (fcp
    .groupby(["FC Code","FC Name","FC City","FC State","FC Cluster"])
    .agg(
        Total_Stock    =("FC Stock",    "sum"),
        Units_Dispatch =("FC Dispatch", "sum"),
        SKUs           =("MSKU",        "nunique"),
        Avg_DOC        =("FC DOC",      "mean"),
    )
    .reset_index()
    .rename(columns={
        "Total_Stock":   "Total Stock",
        "Units_Dispatch":"Units to Dispatch",
        "Avg_DOC":       "Avg DOC",
    })
    .sort_values("Units to Dispatch", ascending=False)
)
fc_disp_sum["Avg DOC"] = fc_disp_sum["Avg DOC"].round(1)


# ══════════════════════════════════════════════════════════════════════════════
#  SECTION 5 — RISK TABLES
# ══════════════════════════════════════════════════════════════════════════════
r_crit   = plan[(plan["DOC"] < 14)     & (plan["Avg/Day"] > 0)].copy()
r_dead   = plan[(plan["Avg/Day"] == 0) & (plan["Sellable Stock"] > 0)].copy()
r_slow   = plan[(plan["Avg/Day"] > 0)  & (plan["DOC"] > 90)].copy()
r_excess = plan[plan["DOC"] > plan_days * 2].copy()
r_top20  = plan[plan["Avg/Day"] > 0].nlargest(20, "Avg/Day").copy()


# ══════════════════════════════════════════════════════════════════════════════
#  EXECUTIVE DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("### 📊 Dashboard")

tot_sell  = int(sku_sell["Sellable Stock"].sum())
tot_dmg   = int(sku_dmg["Damaged Stock"].sum()) if not sku_dmg.empty else 0
tot_disp  = int(fba_plan["Dispatch"].sum())
tot_sold  = int(sales["Qty"].sum())
n_skus    = plan["MSKU"].nunique()
n_fcs     = len(active_fcs)
avg_doc   = plan[plan["Avg/Day"] > 0]["DOC"].replace(9999, np.nan).mean()

m = st.columns(7)
m[0].metric("Total SKUs",         fmtn(n_skus))
m[1].metric("Active FCs",         fmtn(n_fcs))
m[2].metric("Sellable Units",     fmtn(tot_sell))
m[3].metric("Units Sold (period)", fmtn(tot_sold))
m[4].metric("Units to Dispatch",  fmtn(tot_disp))
m[5].metric("🔴 Critical SKUs",   fmtn(len(r_crit)))
m[6].metric("Avg Days of Cover",  f"{avg_doc:.0f}d" if pd.notna(avg_doc) else "N/A")

alert(f"✅ Stock as of <b>{date_label}</b>  |  Sales source: <b>{sales_src}</b>  |  "
      f"Sales window: <b>{sales_win}</b> ({h_days} days)  |  "
      f"Dedup: {n_dedup} rows from {n_raw} raw  |  "
      f"FCs: {', '.join(active_fcs)}", "green")

if len(r_crit):
    alert(f"🚨 <b>{len(r_crit)} SKU(s)</b> are CRITICAL — less than 14 days of stock! Dispatch immediately.", "red")
if len(r_dead):
    alert(f"⚠️ <b>{len(r_dead)} SKU(s)</b> have stock but <b>zero sales</b> — fix listing or plan removal.", "amber")
if tot_dmg:
    alert(f"🔧 <b>{fmtn(tot_dmg)} damaged / non-sellable units</b> — raise Removal Order or Reimbursement Claim.", "amber")
if len(r_excess):
    alert(f"📦 <b>{len(r_excess)} SKU(s)</b> are overstocked (>{plan_days*2} days cover) — pause replenishment.", "grey")

st.markdown("---")


# ══════════════════════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════════════════════
TABS = st.tabs([
    "📋 Full Plan",
    "📦 Dispatch Needed",
    "🏭 FC-Wise View",
    "🚨 Critical & Dead Stock",
    "📈 Sales Trends",
    "📄 Bulk Shipment File",
    "📥 Download Reports",
])


# ─────────────────────────────────────────────────────────────────────────────
# TAB 0 — FULL PLAN
# ─────────────────────────────────────────────────────────────────────────────
with TABS[0]:
    st.subheader("📋 Full Inventory Plan — All SKUs")
    st.caption(f"Every SKU with stock or sales history | {n_skus} SKUs | "
               f"{plan_days}-day plan | {svc_level} service level")

    fa, fb, fc_, fd = st.columns(4)
    with fa:
        f_status = st.multiselect("Filter Status",
            ["🔴 Critical","🟠 Low","🟢 Healthy","🟡 Excess","🔵 Overstocked","⚫ No Stock"],
            key="f0_st")
    with fb:
        f_vel = st.multiselect("Filter Velocity",
            ["🔥 Hot Seller","🟢 Fast","🟡 Medium","🔵 Slow","⚫ Dead"], key="f0_v")
    with fc_:
        f_ch = st.multiselect("Channel", ["FBA","FBM"], key="f0_ch")
    with fd:
        f_disp_only = st.checkbox("Only show dispatch-needed SKUs", key="f0_d")

    view = plan.copy()
    if f_status:    view = view[view["Status"].isin(f_status)]
    if f_vel:       view = view[view["Velocity"].isin(f_vel)]
    if f_ch:        view = view[view["Channel"].isin(f_ch)]
    if f_disp_only: view = view[view["Dispatch"] > 0]

    st.caption(f"Showing **{len(view)}** of {n_skus} SKUs")
    show_plan_table(view, height=520)

    # SKU detail drill-down
    st.markdown("---")
    st.subheader("🔍 SKU Detail")
    sel_sku = st.selectbox("Select MSKU", sorted(plan["MSKU"].unique()), key="drill0")
    if sel_sku:
        row = plan[plan["MSKU"] == sel_sku].iloc[0]
        pn  = row.get("Product Name","")

        st.markdown(f"**{sel_sku}** — {trunc(pn, 90)}")
        st.caption(f"FNSKU: `{row['FNSKU']}`  |  ASIN: `{row['ASIN']}`")

        d1,d2,d3,d4,d5,d6 = st.columns(6)
        d1.metric("Sellable Stock",  fmtn(row["Sellable Stock"]))
        d2.metric("Damaged Stock",   fmtn(row["Damaged Stock"]))
        d3.metric("Avg/Day",         f"{row['Avg/Day']:.3f}")
        d4.metric("Days of Cover",   f"{row['DOC']:.0f}d" if row["DOC"] < 9999 else "∞")
        d5.metric("Dispatch Needed", fmtn(row["Dispatch"]))
        d6.metric("Status",          row["Status"])

        # FC breakdown for this SKU
        sku_fc_view = fc_long[fc_long["MSKU"] == sel_sku].copy()
        if not sku_fc_view.empty:
            sku_fc_view["FC DOC"] = np.where(row["Avg/Day"] > 0,
                (sku_fc_view["FC Stock"] / row["Avg/Day"]).round(1), 9999)
            sku_fc_view["FC Status"] = sku_fc_view["FC DOC"].apply(
                lambda d: health_label(d, plan_days))
            show_df(sku_fc_view[["FC Code","FC Name","FC City","FC Cluster",
                                  "FC Stock","FC DOC","FC Status"]])
        else:
            st.info("No stock at any FC for this SKU.")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 — DISPATCH NEEDED
# ─────────────────────────────────────────────────────────────────────────────
with TABS[1]:
    st.subheader("📦 SKUs That Need Dispatching")

    need = fba_plan[fba_plan["Dispatch"] > 0].sort_values("Priority", ascending=False).copy()
    if need.empty:
        alert("✅ All FBA SKUs are well-stocked. No dispatch required right now!", "green")
    else:
        st.caption(f"{len(need)} SKUs need restocking — "
                   f"{fmtn(int(need['Dispatch'].sum()))} total units | "
                   f"Sorted by priority score")
        show_plan_table(need, height=440)

        st.markdown("---")
        st.subheader("🏭 FC-Wise Dispatch Breakdown")
        st.caption("How many units to send to each FC (allocated proportionally by current stock share)")

        fc_need = fcp[(fcp["Channel"] == "FBA") & (fcp["FC Dispatch"] > 0)].copy()
        fc_need = fc_need.sort_values(["FC Code","Priority"], ascending=[True, False])
        fc_need["Product Name"] = fc_need["Product Name"].apply(lambda x: trunc(x, 50))
        show_df(fc_need[["FC Code","FC Name","FC City","FC Cluster",
                          "MSKU","FNSKU","Product Name",
                          "FC Stock","Avg/Day","FC DOC","FC Dispatch"]],
                height=440)

    if not fbm_plan.empty and fbm_plan["Dispatch"].sum() > 0:
        st.markdown("---")
        st.subheader("📮 FBM Dispatch")
        show_plan_table(fbm_plan[fbm_plan["Dispatch"] > 0], height=300)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 — FC-WISE VIEW
# ─────────────────────────────────────────────────────────────────────────────
with TABS[2]:
    st.subheader("🏭 Fulfillment Center Summary")

    section("FC Inventory & Dispatch Overview")
    show_df(fc_disp_sum, height=320)

    st.markdown("---")
    section("Drill Into a Specific FC")

    if active_fcs:
        fc_opts = [f"{fc} — {fc_name(fc)}" for fc in active_fcs]
        sel_fc_str = st.selectbox("Choose FC", fc_opts, key="fc_drill")
        sel_fc_code = sel_fc_str.split(" — ")[0].strip()

        fc_skus = fcp[fcp["FC Code"] == sel_fc_code].copy()
        fc_skus = fc_skus.sort_values("Priority", ascending=False)
        fc_skus["Product Name"] = fc_skus["Product Name"].apply(lambda x: trunc(x, 50))

        tot_fc_stock    = int(fc_skus["FC Stock"].sum())
        tot_fc_dispatch = int(fc_skus["FC Dispatch"].sum())
        avg_fc_doc      = fc_skus[fc_skus["Avg/Day"]>0]["FC DOC"].replace(9999, np.nan).mean()

        ci1,ci2,ci3,ci4 = st.columns(4)
        ci1.metric("FC", sel_fc_code)
        ci2.metric("Total Sellable", fmtn(tot_fc_stock))
        ci3.metric("Units to Dispatch", fmtn(tot_fc_dispatch))
        ci4.metric("Avg FC DOC", f"{avg_fc_doc:.0f}d" if pd.notna(avg_fc_doc) else "N/A")
        st.caption(f"{fc_name(sel_fc_code)} | {fc_city(sel_fc_code)} | "
                   f"{fc_state(sel_fc_code)} | {fc_cluster(sel_fc_code)}")

        show_df(fc_skus[["FC Code","MSKU","FNSKU","Product Name",
                          "FC Stock","Avg/Day","FC DOC","FC Status","FC Dispatch"]],
                height=420)

    st.markdown("---")
    section("All FC Inventory Details")
    fc_inv_show = fc_inv_summary.copy()
    show_df(fc_inv_show, height=350)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 — CRITICAL & DEAD STOCK
# ─────────────────────────────────────────────────────────────────────────────
with TABS[3]:
    st.subheader("🚨 Inventory Health Alerts")

    t0, t1, t2, t3, t4 = st.tabs([
        f"🔴 Critical  ({len(r_crit)})",
        f"⚫ Dead Stock  ({len(r_dead)})",
        f"🟡 Slow Moving  ({len(r_slow)})",
        f"🔵 Excess  ({len(r_excess)})",
        f"🔧 Damaged",
    ])

    def risk_cols(df):
        d = df.copy()
        d["Product Name"] = d["Product Name"].apply(lambda x: trunc(x, 58))
        rc = ["MSKU","FNSKU","Product Name","Sellable Stock","Damaged Stock",
              "Avg/Day","DOC","Status","All Time Sales","Dispatch"]
        return d[[c for c in rc if c in d.columns]].reset_index(drop=True)

    with t0:
        st.caption("Stock below 14 days — dispatch as soon as possible")
        if r_crit.empty:
            alert("✅ No critical SKUs right now!", "green")
        else:
            show_df(risk_cols(r_crit.sort_values("DOC")))

    with t1:
        st.caption("These SKUs have sellable stock but zero sales — "
                   "check listing, pricing or plan removal")
        if r_dead.empty:
            alert("✅ No dead stock found!", "green")
        else:
            d2 = r_dead.copy()
            d2["Product Name"] = d2["Product Name"].apply(lambda x: trunc(x, 58))
            show_df(d2[["MSKU","FNSKU","ASIN","Product Name","Sellable Stock",
                        "Damaged Stock","All Time Sales"]].reset_index(drop=True))

            # Where is the dead stock physically sitting?
            st.markdown("---")
            st.markdown("**📍 Dead Stock Location — which FC is holding it?**")
            dead_fc_view = fc_long[fc_long["MSKU"].isin(r_dead["MSKU"])].copy()
            dead_fc_view["Product Name"] = (dead_fc_view["MSKU"]
                .map(sku_df["Title"]).fillna("").apply(lambda x: trunc(str(x), 50)))
            show_df(dead_fc_view[["FC Code","FC Name","FC City","FC Cluster",
                                   "MSKU","Product Name","FC Stock"]]
                    .sort_values("FC Stock", ascending=False), height=300)

    with t2:
        st.caption("Days of cover > 90 — consider pausing replenishment")
        if r_slow.empty:
            alert("✅ No slow-moving SKUs.", "green")
        else:
            show_df(risk_cols(r_slow.sort_values("DOC", ascending=False)))

    with t3:
        st.caption(f"Days of cover > {plan_days*2} days — significantly overstocked")
        if r_excess.empty:
            alert("✅ No overstocked SKUs.", "green")
        else:
            show_df(risk_cols(r_excess.sort_values("DOC", ascending=False)))

    with t4:
        st.caption("Raise a Removal Order or Reimbursement Claim in Seller Central")
        if dmg_by_msku.empty:
            alert("✅ No damaged or non-sellable stock found.", "green")
        else:
            show_df(dmg_by_msku.sort_values("Qty", ascending=False), height=280)
            if not dmg_by_fc.empty:
                st.markdown("**📍 Damaged Stock by FC:**")
                show_df(dmg_by_fc.sort_values("Qty", ascending=False), height=250)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 — SALES TRENDS
# ─────────────────────────────────────────────────────────────────────────────
with TABS[4]:
    st.subheader("📈 Sales Trends & Analytics")

    st.caption(f"Sales range: {s_min.strftime('%d %b %Y')} → {s_max.strftime('%d %b %Y')} "
               f"({all_days} days total) | Window used: {sales_win} ({h_days} days)")

    sa, sb = st.columns(2)
    with sa:
        section("Weekly Sales")
        if not wk_trend.empty:
            st.line_chart(wk_trend.set_index("Week")["Units"], height=220)
        else:
            st.info("No weekly data.")
    with sb:
        section("Monthly Sales by Channel")
        if not mo_trend.empty:
            pm = mo_trend.pivot_table(index="Month", columns="Channel",
                                       values="Units", aggfunc="sum", fill_value=0)
            st.bar_chart(pm, height=220)

    st.markdown("---")
    section("Top 20 Fastest-Moving SKUs")
    top20 = r_top20.copy()
    top20["Product Name"] = top20["Product Name"].apply(lambda x: trunc(x, 55))
    show_df(top20[["MSKU","FNSKU","Product Name","Avg/Day","Velocity",
                   "Sellable Stock","DOC","Status","All Time Sales"]], height=400)

    st.markdown("---")
    section("SKU Sales Drill-Down")
    spick = st.selectbox("Select MSKU", sorted(plan["MSKU"].unique()), key="s_pick")
    if spick:
        spn = trunc(sku_df["Title"].get(spick, spick), 80)
        st.caption(f"**{spick}** — {spn}")
        sd = (sales[sales["MSKU"] == spick]
              .groupby("Date")["Qty"].sum().reset_index().set_index("Date"))
        if not sd.empty:
            st.line_chart(sd["Qty"], height=200)
        else:
            st.info("No sales data for this SKU in the selected window.")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 5 — BULK SHIPMENT FILE
# Amazon FlatFile format — one upload creates all FC shipments automatically
# ─────────────────────────────────────────────────────────────────────────────
with TABS[5]:
    st.subheader("📄 Amazon FBA Bulk Shipment Upload File")

    alert(
        "📋 <b>How to use this file:</b><br>"
        "1. Edit quantities in the table if needed<br>"
        "2. Click <b>Download Bulk Shipment File</b><br>"
        "3. In Seller Central: <b>Send to Amazon → Create New Shipment → Upload a file</b><br>"
        "4. Amazon reads the file and <b>automatically creates one shipment per FC</b> ✅<br>"
        "5. No manual FC selection needed — all FCs covered in one file",
        "blue"
    )

    bf1, bf2 = st.columns(2)
    with bf1:
        shp_ref  = st.text_input("Shipment Reference / Name",
            value=f"Dispatch_{datetime.now().strftime('%d%b%Y')}", key="shp_ref")
        shp_date = st.date_input("Expected Ship Date",
            value=datetime.now().date() + timedelta(days=2), key="shp_dt")
    with bf2:
        sel_fcs  = st.multiselect("Include FCs (blank = all active FCs)",
            active_fcs, default=[], key="sel_fcs")
        split_by_fc = st.checkbox("Generate separate file per FC instead", key="split_fc")

    # Build dispatch rows
    dispatch_rows = fcp[
        (fcp["Channel"] == "FBA") & (fcp["FC Dispatch"] > 0)
    ].copy()

    if sel_fcs:
        dispatch_rows = dispatch_rows[dispatch_rows["FC Code"].isin(sel_fcs)]

    if dispatch_rows.empty:
        alert("No FBA dispatch required right now. All stock levels are adequate.", "green")
    else:
        dispatch_rows = dispatch_rows.copy()
        dispatch_rows["Units to Ship"] = dispatch_rows["FC Dispatch"].astype(int)
        dispatch_rows["Cases"]         = np.ceil(
            dispatch_rows["Units to Ship"] / case_qty).astype(int)
        dispatch_rows["Product Name"]  = dispatch_rows["Product Name"].apply(
            lambda x: trunc(x, 60))

        edit_cols = ["FC Code","FC Name","MSKU","FNSKU","Product Name",
                     "Units to Ship","Cases"]
        dispatch_rows = dispatch_rows.sort_values(["FC Code","Priority"],
                                                   ascending=[True, False])

        st.caption(
            f"**{len(dispatch_rows)} lines**  |  "
            f"**{fmtn(dispatch_rows['Units to Ship'].sum())} total units**  |  "
            f"**{dispatch_rows['FC Code'].nunique()} FCs**  |  "
            "Edit quantities below if needed"
        )

        edited = st.data_editor(
            dispatch_rows[edit_cols].reset_index(drop=True),
            num_rows="dynamic",
            use_container_width=True,
            key="edit_dispatch",
        )

        # ── FLAT FILE GENERATOR ───────────────────────────────────────────
        def make_flat_file(rows_df, fc_code, shipment_name):
            """Generate Amazon FlatFileShipmentCreation format for one FC."""
            lines = [
                "TemplateType=FlatFileShipmentCreation\tVersion=2015.0403",
                "",
            ]
            # Shipment header row
            hdr_fields = [
                "ShipmentName","ShipFromName","ShipFromAddressLine1",
                "ShipFromCity","ShipFromStateOrProvinceCode",
                "ShipFromPostalCode","ShipFromCountryCode",
                "ShipmentStatus","LabelPrepType",
                "AreCasesRequired","DestinationFulfillmentCenterId",
            ]
            hdr_values = [
                shipment_name,
                wh_name  or "My Warehouse",
                wh_addr  or "Warehouse Address",
                wh_city  or "City",
                (wh_state or "UP").upper()[:2],
                wh_pin   or "000000",
                "IN",
                "WORKING",
                lbl_own,
                "YES" if case_packed else "NO",
                fc_code,
            ]
            lines.append("\t".join(hdr_fields))
            lines.append("\t".join(str(v) for v in hdr_values))
            lines.append("")
            # Item header
            lines.append("\t".join([
                "SellerSKU","FNSKU","QuantityShipped",
                "QuantityInCase","PrepOwner","LabelOwner",
                "ItemDescription","ExpectedDeliveryDate",
            ]))
            # Item rows
            for _, r in rows_df.iterrows():
                qty = int(r.get("Units to Ship", 0))
                if qty <= 0:
                    continue
                qic  = str(case_qty) if case_packed else ""
                desc = str(r.get("Product Name", ""))[:200]
                msku = str(r.get("MSKU",""))
                fnsk = str(r.get("FNSKU",""))
                lines.append("\t".join([
                    msku, fnsk, str(qty), qic,
                    prep_own, lbl_own,
                    desc, str(shp_date),
                ]))
            return "\n".join(lines)

        st.markdown("---")

        if not split_by_fc:
            # ── COMBINED FILE (one file, Amazon auto-splits per FC) ───────
            fcs_in  = sorted(edited["FC Code"].unique())
            blocks  = []
            for fc in fcs_in:
                rows = edited[edited["FC Code"] == fc]
                if rows.empty:
                    continue
                blocks.append(make_flat_file(
                    rows,
                    fc,
                    f"{shp_ref}_{fc}",
                ))
            combined_txt = "\n\n".join(blocks)

            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label=f"📥 Download Bulk Shipment File ({len(fcs_in)} FCs, {fmtn(edited['Units to Ship'].sum())} units)",
                    data=combined_txt.encode("utf-8"),
                    file_name=f"FBA_Bulk_Shipment_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                    mime="text/plain",
                    use_container_width=True,
                )
            with col_dl2:
                st.download_button(
                    label="📋 Download Dispatch List (CSV)",
                    data=edited.to_csv(index=False),
                    file_name=f"Dispatch_List_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            with st.expander("👁️ Preview first 3,000 characters of flat file"):
                st.code(combined_txt[:3000] +
                        ("\n...(truncated)" if len(combined_txt) > 3000 else ""))
        else:
            # ── SEPARATE FILE PER FC ──────────────────────────────────────
            fcs_in = sorted(edited["FC Code"].unique())
            st.markdown(f"**Generating {len(fcs_in)} separate files:**")
            dl_cols = st.columns(min(len(fcs_in), 4))
            for i, fc in enumerate(fcs_in):
                rows = edited[edited["FC Code"] == fc]
                if rows.empty:
                    continue
                txt = make_flat_file(rows, fc, f"{shp_ref}_{fc}")
                with dl_cols[i % 4]:
                    st.download_button(
                        label=f"📥 {fc}\n{fc_name(fc)}\n({int(rows['Units to Ship'].sum())} units)",
                        data=txt.encode("utf-8"),
                        file_name=f"FBA_{fc}_{datetime.now().strftime('%Y%m%d')}.txt",
                        mime="text/plain",
                        key=f"dl_fc_{fc}",
                    )


# ─────────────────────────────────────────────────────────────────────────────
# TAB 6 — DOWNLOAD REPORTS
# ─────────────────────────────────────────────────────────────────────────────
with TABS[6]:
    st.subheader("📥 Download Full Reports")

    def build_excel_report():
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:

            # 1. Full Plan
            fp = plan.copy()
            fp["Product Name"] = fp["Product Name"].apply(lambda x: trunc(x,90))
            fp = fp.sort_values("Priority", ascending=False)
            xl_write(w, fp[[c for c in DISP_COLS if c in fp.columns]], "Full Plan", freeze_col=1)

            # 2. FBA Dispatch Needed
            fd_ = fba_plan[fba_plan["Dispatch"] > 0].copy()
            fd_["Product Name"] = fd_["Product Name"].apply(lambda x: trunc(x,90))
            xl_write(w, fd_[[c for c in DISP_COLS if c in fd_.columns]]
                     .sort_values("Priority",ascending=False), "FBA Dispatch Needed", freeze_col=1)

            # 3. FC-Wise Dispatch
            fco = fcp[(fcp["Channel"]=="FBA")].copy()
            fco["Product Name"] = fco["Product Name"].apply(lambda x: trunc(x,90))
            fco = fco.sort_values(["FC Code","Priority"], ascending=[True,False])
            xl_write(w, fco[["FC Code","FC Name","FC City","FC State","FC Cluster",
                              "MSKU","FNSKU","Product Name","FC Stock",
                              "Avg/Day","FC DOC","FC Status","FC Dispatch"]],
                     "FC-Wise Dispatch", freeze_col=2)

            # 4. FC Summary
            xl_write(w, fc_disp_sum, "FC Summary")

            # 5. Critical SKUs
            cr = r_crit.copy()
            cr["Product Name"] = cr["Product Name"].apply(lambda x: trunc(x,90))
            xl_write(w, cr[[c for c in DISP_COLS if c in cr.columns]]
                     .sort_values("DOC"), "Critical SKUs", freeze_col=1)

            # 6. Dead Stock
            dd = r_dead.copy()
            dd["Product Name"] = dd["Product Name"].apply(lambda x: trunc(x,90))
            xl_write(w, dd[["MSKU","FNSKU","ASIN","Product Name","Sellable Stock",
                             "Damaged Stock","All Time Sales"]], "Dead Stock", freeze_col=1)

            # 7. Dead Stock by FC location
            dl = fc_long[fc_long["MSKU"].isin(r_dead["MSKU"])].copy()
            dl["Product Name"] = dl["MSKU"].map(sku_df["Title"]).apply(lambda x: trunc(str(x),90))
            xl_write(w, dl[["FC Code","FC Name","FC City","FC Cluster",
                             "MSKU","Product Name","FC Stock"]]
                     .sort_values("FC Stock",ascending=False),
                     "Dead Stock by FC", freeze_col=2)

            # 8. Slow Moving
            sl = r_slow.copy()
            sl["Product Name"] = sl["Product Name"].apply(lambda x: trunc(x,90))
            xl_write(w, sl[[c for c in DISP_COLS if c in sl.columns]]
                     .sort_values("DOC",ascending=False), "Slow Moving", freeze_col=1)

            # 9. Excess Stock
            ex = r_excess.copy()
            ex["Product Name"] = ex["Product Name"].apply(lambda x: trunc(x,90))
            xl_write(w, ex[[c for c in DISP_COLS if c in ex.columns]]
                     .sort_values("DOC",ascending=False), "Excess Stock", freeze_col=1)

            # 10. Damaged Stock
            if not dmg_by_msku.empty:
                xl_write(w, dmg_by_msku.sort_values("Qty",ascending=False),
                         "Damaged Stock")

            # 11. Monthly Trend
            xl_write(w, mo_trend, "Monthly Trend")

            # 12. Weekly Trend
            xl_write(w, wk_trend, "Weekly Trend")

            # 13. SKU Master
            sm = sku_df.reset_index().copy()
            sm["Title"] = sm["Title"].apply(lambda x: trunc(x,120))
            xl_write(w, sm[["MSKU","FNSKU","ASIN","Title"]], "SKU Master")

            # 14. FC Master
            xl_write(w, pd.DataFrame([
                {"FC Code":k,"FC Name":v[0],"City":v[1],"State":v[2],"Cluster":v[3]}
                for k,v in FC_MASTER.items()
            ]), "FC Master Reference")

        buf.seek(0)
        return buf

    st.markdown("**Choose your download:**")
    dc1, dc2, dc3 = st.columns(3)

    with dc1:
        st.download_button(
            label="📗 Full Excel Report\n(14 sheets — complete data)",
            data=build_excel_report(),
            file_name=f"FBA_Supply_Plan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with dc2:
        _csv1 = fba_plan[fba_plan["Dispatch"] > 0].copy()
        _csv1["Product Name"] = _csv1["Product Name"].apply(lambda x: trunc(x,90))
        st.download_button(
            label="📋 Dispatch Plan CSV\n(FBA SKUs needing restock)",
            data=_csv1[[c for c in DISP_COLS if c in _csv1.columns]]
                 .sort_values("Priority",ascending=False).to_csv(index=False),
            file_name=f"Dispatch_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with dc3:
        _fcc = fcp[(fcp["Channel"]=="FBA") & (fcp["FC Dispatch"]>0)].copy()
        _fcc["Product Name"] = _fcc["Product Name"].apply(lambda x: trunc(x,90))
        st.download_button(
            label="🏭 FC-Wise Dispatch CSV\n(1 row per SKU per FC)",
            data=_fcc[["FC Code","FC Name","FC City","MSKU","FNSKU","Product Name",
                        "FC Stock","Avg/Day","FC DOC","FC Dispatch"]]
                  .sort_values(["FC Code","Priority"],ascending=[True,False])
                  .to_csv(index=False),
            file_name=f"FC_Dispatch_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("**📑 Report contents:**")
    st.markdown("""
| Sheet / File | What's inside |
|---|---|
| **Full Plan** | All SKUs — MSKU, Product Name, stock, avg daily sale, dispatch, status |
| **FBA Dispatch Needed** | Only SKUs that need restocking (sorted by priority) |
| **FC-Wise Dispatch** | 1 row per SKU per FC — how many units to each location |
| **FC Summary** | 1 row per FC — total stock, units to dispatch, avg days of cover |
| **Critical SKUs** | Stock < 14 days — urgent replenishment needed |
| **Dead Stock** | Stock exists but zero sales — needs action |
| **Dead Stock by FC** | Dead stock with exact FC location |
| **Slow Moving** | Sales exist but > 90 days of cover |
| **Excess Stock** | Significantly overstocked |
| **Damaged Stock** | Non-sellable units by MSKU and FC |
| **Monthly / Weekly Trend** | Sales trend charts data |
| **SKU Master** | MSKU → FNSKU → ASIN → Product Name mapping |
| **FC Master Reference** | All 75 Amazon India FC codes |
""")
