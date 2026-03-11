"""
FBA Smart Supply Planner  ·  Amazon India
==========================================
Upload Inventory Ledger (+ optional MTR) to get:
  • Full inventory plan with Product Name & SKU
  • FC-wise stock and dispatch allocation
  • Dead stock, critical, slow-moving alerts
  • Amazon FBA Bulk Shipment Upload File
  • Excel & CSV reports (14 sheets)

KEY FIX: sort BEFORE column-slice everywhere — no KeyError on Priority / FC Code.
"""

import io, math, zipfile
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FBA Supply Planner — Amazon India",
    layout="wide",
    page_icon="📦",
    initial_sidebar_state="expanded",
)
st.markdown("""<style>
[data-testid="stMetricValue"]{font-size:1.65rem;font-weight:800;color:#1a202c}
[data-testid="stMetricLabel"]{font-size:.75rem;font-weight:600;color:#718096;
    text-transform:uppercase;letter-spacing:.06em}
.stTabs [data-baseweb="tab-list"]{gap:4px;background:#f0f4f8;padding:4px 6px;
    border-radius:10px;margin-bottom:4px}
.stTabs [data-baseweb="tab"]{font-size:13px;font-weight:600;padding:5px 14px;
    border-radius:7px}
.stTabs [aria-selected="true"]{background:#2b6cb0!important;color:#fff!important}
.al{padding:9px 15px;border-radius:7px;font-weight:600;font-size:.88rem;
    margin:3px 0;line-height:1.5}
.al-g{background:#f0fff4;border-left:5px solid #38a169;color:#1c4532}
.al-r{background:#fff5f5;border-left:5px solid #e53e3e;color:#742a2a}
.al-a{background:#fffaf0;border-left:5px solid #ed8936;color:#7b341e}
.al-b{background:#ebf8ff;border-left:5px solid #3182ce;color:#1a365d}
.al-y{background:#fffff0;border-left:5px solid #d69e2e;color:#744210}
.sh{font-size:1rem;font-weight:700;color:#2d3748;padding:3px 0 2px;
    border-bottom:2px solid #e2e8f0;margin-bottom:6px}
</style>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
#  FC MASTER  (75 Amazon India FCs)
#  Tuple: (Name, City, State, Cluster)
#  PNQ2 = New Delhi (Mohan Co-op A-33) — NOT Pune
# ─────────────────────────────────────────────────────────────────────────────
FC_MASTER = {
    "SGAA":("Guwahati FC","Guwahati","ASSAM","East Assam"),
    "DEX3":("New Delhi FC A28","New Delhi","DELHI","Delhi NCR"),
    "DEX8":("New Delhi FC A29","New Delhi","DELHI","Delhi NCR"),
    "PNQ2":("New Delhi FC A33","New Delhi","DELHI","Delhi NCR"),
    "XDEL":("Delhi XL FC","New Delhi","DELHI","Delhi NCR"),
    "DEL2":("Tauru FC","Mewat","HARYANA","Delhi NCR"),
    "DEL4":("Gurgaon FC","Gurgaon","HARYANA","Delhi NCR"),
    "DEL5":("Manesar FC","Manesar","HARYANA","Delhi NCR"),
    "DEL6":("Bilaspur FC","Bilaspur","HARYANA","Delhi NCR"),
    "DEL7":("Bilaspur FC 2","Bilaspur","HARYANA","Delhi NCR"),
    "DEL8":("Sohna FC","Sohna","HARYANA","Delhi NCR"),
    "DEL8_DED5":("Sohna ESR FC","Sohna","HARYANA","Delhi NCR"),
    "DED3":("Farrukhnagar FC","Farrukhnagar","HARYANA","Delhi NCR"),
    "DED4":("Sohna FC 2","Sohna","HARYANA","Delhi NCR"),
    "DED5":("Sohna FC 3","Sohna","HARYANA","Delhi NCR"),
    "AMD1":("Ahmedabad Naroda FC","Naroda","GUJARAT","Gujarat West"),
    "AMD2":("Changodar FC","Changodar","GUJARAT","Gujarat West"),
    "SUB1":("Surat FC","Surat","GUJARAT","Gujarat West"),
    "BLR4":("Devanahalli FC","Devanahalli","KARNATAKA","Bangalore"),
    "BLR5":("Bommasandra FC","Bommasandra","KARNATAKA","Bangalore"),
    "BLR6":("Nelamangala FC","Nelamangala","KARNATAKA","Bangalore"),
    "BLR7":("Hoskote FC","Hoskote","KARNATAKA","Bangalore"),
    "BLR8":("Devanahalli FC 2","Devanahalli","KARNATAKA","Bangalore"),
    "BLR10":("Kudlu Gate FC","Bengaluru","KARNATAKA","Bangalore"),
    "BLR12":("Attibele FC","Attibele","KARNATAKA","Bangalore"),
    "BLR13":("Jigani FC","Jigani","KARNATAKA","Bangalore"),
    "BLR14":("Anekal FC","Anekal","KARNATAKA","Bangalore"),
    "SCJA":("Bangalore SCJA FC","Bengaluru","KARNATAKA","Bangalore"),
    "XSAB":("Bangalore XS FC","Bengaluru","KARNATAKA","Bangalore"),
    "SBLA":("Bangalore SBLA FC","Bengaluru","KARNATAKA","Bangalore"),
    "SIDA":("Indore FC","Indore","MADHYA PRADESH","Central"),
    "FBHB":("Bhopal FC","Bhopal","MADHYA PRADESH","Central"),
    "FIDA":("Bhopal FC 2","Bhopal","MADHYA PRADESH","Central"),
    "IND1":("Indore FC 2","Indore","MADHYA PRADESH","Central"),
    "BOM1":("Bhiwandi FC","Bhiwandi","MAHARASHTRA","Mumbai West"),
    "BOM3":("Nashik FC","Nashik","MAHARASHTRA","Mumbai West"),
    "BOM4":("Vashere FC","Bhiwandi","MAHARASHTRA","Mumbai West"),
    "BOM5":("Bhiwandi 2 FC","Bhiwandi","MAHARASHTRA","Mumbai West"),
    "BOM7":("Bhiwandi 3 FC","Bhiwandi","MAHARASHTRA","Mumbai West"),
    "ISK3":("Bhiwandi ISK3 FC","Bhiwandi","MAHARASHTRA","Mumbai West"),
    "PNQ1":("Pune Hinjewadi FC","Pune","MAHARASHTRA","Mumbai West"),
    "PNQ3":("Pune Chakan FC","Pune","MAHARASHTRA","Mumbai West"),
    "SAMB":("Mumbai SAMB FC","Mumbai","MAHARASHTRA","Mumbai West"),
    "XBOM":("Mumbai XL FC","Bhiwandi","MAHARASHTRA","Mumbai West"),
    "ATX1":("Ludhiana FC","Ludhiana","PUNJAB","North Punjab"),
    "LDH1":("Ludhiana FC 2","Ludhiana","PUNJAB","North Punjab"),
    "RAJ1":("Rajpura FC","Rajpura","PUNJAB","North Punjab"),
    "JAI1":("Jaipur FC","Jaipur","RAJASTHAN","North Rajasthan"),
    "JPX1":("Jaipur JPX1 FC","Jaipur","RAJASTHAN","North Rajasthan"),
    "JPX2":("Bagru FC","Bagru","RAJASTHAN","North Rajasthan"),
    "MAA1":("Irungattukottai FC","Chennai","TAMIL NADU","Chennai"),
    "MAA2":("Ponneri FC","Ponneri","TAMIL NADU","Chennai"),
    "MAA3":("Sriperumbudur FC","Sriperumbudur","TAMIL NADU","Chennai"),
    "MAA4":("Ambattur FC","Ambattur","TAMIL NADU","Chennai"),
    "MAA5":("Kanchipuram FC","Kanchipuram","TAMIL NADU","Chennai"),
    "CJB1":("Coimbatore FC","Coimbatore","TAMIL NADU","Chennai"),
    "SMAB":("Chennai SMAB FC","Chennai","TAMIL NADU","Chennai"),
    "COK1":("Kochi FC","Kochi","KERALA","South Kerala"),
    "HYD3":("Shamshabad FC","Hyderabad","TELANGANA","Hyderabad"),
    "HYD6":("Kothur FC","Kothur","TELANGANA","Hyderabad"),
    "HYD7":("Medchal FC","Medchal","TELANGANA","Hyderabad"),
    "HYD8":("Shamshabad FC 2","Hyderabad","TELANGANA","Hyderabad"),
    "HYD8_HYD3":("Shamshabad Mamidipally FC","Hyderabad","TELANGANA","Hyderabad"),
    "HYD9":("Pedda Amberpet FC","Hyderabad","TELANGANA","Hyderabad"),
    "HYD10":("Ghatkesar FC","Ghatkesar","TELANGANA","Hyderabad"),
    "HYD11":("Kompally FC","Kompally","TELANGANA","Hyderabad"),
    "LKO1":("Lucknow FC","Lucknow","UTTAR PRADESH","North UP"),
    "AGR1":("Agra FC","Agra","UTTAR PRADESH","North UP"),
    "SLDK":("Kishanpur FC","Bijnour","UTTAR PRADESH","North UP"),
    "CCU1":("Kolkata FC","Kolkata","WEST BENGAL","Kolkata East"),
    "CCU2":("Kolkata FC 2","Kolkata","WEST BENGAL","Kolkata East"),
    "CCX1":("Howrah FC","Howrah","WEST BENGAL","Kolkata East"),
    "CCX2":("Howrah FC 2","Howrah","WEST BENGAL","Kolkata East"),
    "PAT1":("Patna FC","Patna","BIHAR","Kolkata East"),
    "BBS1":("Bhubaneswar FC","Bhubaneswar","ODISHA","Kolkata East"),
}
def fc_name(c):    return FC_MASTER.get(str(c).upper(),(str(c),))[0]
def fc_city(c):    return FC_MASTER.get(str(c).upper(),("","?"))[1]
def fc_state(c):   return FC_MASTER.get(str(c).upper(),("","","?"))[2]
def fc_cluster(c): return FC_MASTER.get(str(c).upper(),("","","","Other"))[3]


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def read_file(f):
    try:
        n = f.name.lower()
        if n.endswith(".zip"):
            with zipfile.ZipFile(f) as z:
                for m in z.namelist():
                    if m.lower().endswith(".csv"):
                        return pd.read_csv(z.open(m), low_memory=False)
            return pd.DataFrame()
        if n.endswith((".xlsx",".xls")):
            return pd.read_excel(f)
        return pd.read_csv(f, low_memory=False)
    except Exception as e:
        st.warning(f"Cannot read {f.name}: {e}")
        return pd.DataFrame()

def gcol(df, names):
    """First matching column name (case-insensitive). Returns '' if none."""
    low = {c.strip().lower(): c for c in df.columns}
    for n in names:
        if n.strip().lower() in low:
            return low[n.strip().lower()]
    return ""

def parse_dates(s):
    out = pd.to_datetime(s, errors="coerce", dayfirst=False)
    bad = out.isna()
    if bad.any():
        out.loc[bad] = pd.to_datetime(s.loc[bad], errors="coerce", dayfirst=True)
    return out

def trunc(s, n=60):
    s = str(s)
    return s[:n]+"…" if len(s)>n else s

def fmt(x):
    try:    return f"{int(x):,}"
    except: return "—"

def health(doc, pd_):
    if doc<=0:      return "⚫ No Stock"
    if doc<14:      return "🔴 Critical"
    if doc<30:      return "🟠 Low"
    if doc<pd_:     return "🟢 Healthy"
    if doc<pd_*2:   return "🟡 Excess"
    return "🔵 Overstocked"

def vel(avg):
    if avg<=0:   return "⚫ Dead"
    if avg<0.5:  return "🔵 Slow"
    if avg<2:    return "🟡 Medium"
    if avg<5:    return "🟢 Fast"
    return "🔥 Hot"

def al(msg, k="g"):
    cls={"g":"al-g","r":"al-r","a":"al-a","b":"al-b","y":"al-y"}[k]
    st.markdown(f'<div class="al {cls}">{msg}</div>', unsafe_allow_html=True)

def sh(t):
    st.markdown(f'<div class="sh">{t}</div>', unsafe_allow_html=True)

def sdf(df, h=420):
    st.dataframe(df.reset_index(drop=True), use_container_width=True,
                 hide_index=True, height=h)

def safe_cols(df, wanted):
    """Return only columns that actually exist in df."""
    return [c for c in wanted if c in df.columns]

def safe_sort(df, by, ascending=None):
    """sort_values only on columns that exist — no KeyError."""
    if isinstance(by, str):
        by = [by]
    if ascending is None:
        ascending = [True]*len(by)
    elif isinstance(ascending, bool):
        ascending = [ascending]*len(by)
    valid_by  = [b for b in by if b in df.columns]
    valid_asc = [ascending[i] for i,b in enumerate(by) if b in df.columns]
    if not valid_by:
        return df
    return df.sort_values(valid_by, ascending=valid_asc)

def xl(writer, df, sheet, freeze=0):
    if df is None or df.empty:
        pd.DataFrame({"(No data)":[]}).to_excel(writer, sheet_name=sheet, index=False)
        return
    d = df.copy()
    for c in d.select_dtypes("object").columns:
        d[c] = d[c].astype(str).str[:120]
    d.to_excel(writer, sheet_name=sheet, index=False, startrow=1)
    wb = writer.book; ws = writer.sheets[sheet]
    hf = wb.add_format({"bold":True,"bg_color":"#1a365d","font_color":"white",
                         "border":1,"align":"center","valign":"vcenter",
                         "text_wrap":True,"font_size":10})
    for ci,cn in enumerate(d.columns):
        ws.write(0, ci, cn, hf)
        w_ = max(len(str(cn)),
                 d[cn].astype(str).str.len().max() if not d.empty else 8)
        ws.set_column(ci, ci, min(w_+2, 48))
    ws.freeze_panes(2, freeze)
    ws.autofilter(0, 0, len(d), len(d.columns)-1)
    ws.set_row(0, 22)


# ─────────────────────────────────────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Planning Settings")
    plan_days = st.number_input("Target Coverage Days", 7, 180, 60)
    svc       = st.selectbox("Service Level", ["95%","90%","98%"])
    Z         = {"90%":1.28,"95%":1.65,"98%":2.05}[svc]
    s_win     = st.selectbox("Sales Window",
                    ["Full History","Last 30 Days","Last 60 Days","Last 90 Days"])
    min_disp  = st.number_input("Min Units to Show in Dispatch", 0, 9999, 0)

    st.divider()
    st.markdown("## 🏭 Your Warehouse")
    wh_name  = st.text_input("Company Name",  placeholder="GLOOYA Warehouse")
    wh_addr  = st.text_input("Address Line 1",placeholder="Plot 12, Indl Area")
    wh_city  = st.text_input("City",          placeholder="Lucknow")
    wh_state = st.text_input("State Code",    placeholder="UP", max_chars=2).upper()
    wh_pin   = st.text_input("PIN Code",      placeholder="226001")

    st.divider()
    st.markdown("## 📦 Shipment Defaults")
    case_qty    = st.number_input("Units Per Carton", 1, 500, 12)
    case_packed = st.checkbox("Case-Packed")
    prep_own    = st.selectbox("Prep Owner",  ["AMAZON","SELLER"])
    lbl_own     = st.selectbox("Label Owner", ["AMAZON","SELLER"])


# ─────────────────────────────────────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────────────────────────────────────
h1, h2 = st.columns([4,1])
with h1:
    st.markdown("# 📦 FBA Smart Supply Planner")
    st.caption("Amazon India · Inventory Ledger + MTR → Full dispatch plan, "
               "dead stock alerts & bulk shipment file")
with h2:
    st.caption(f"Run: {datetime.now().strftime('%d %b %Y %H:%M')}")
st.divider()


# ─────────────────────────────────────────────────────────────────────────────
#  FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Step 1 — Upload Reports")
st.caption("Hold **Ctrl / Cmd** while clicking to select multiple files, "
           "or drag-drop several at once")

u1, u2 = st.columns(2)
with u1:
    inv_files = st.file_uploader(
        "📊 **Inventory Ledger** — Required",
        type=["csv","zip","xlsx"], accept_multiple_files=True, key="inv_up",
        help="Seller Central → Reports → Fulfillment → Inventory Ledger")
with u2:
    mtr_files = st.file_uploader(
        "📋 **MTR Sales Report** — Optional",
        type=["csv","zip","xlsx"], accept_multiple_files=True, key="mtr_up",
        help="Seller Central → Reports → Tax → MTR")

with st.expander("📖 File format guide & FC reference"):
    st.markdown("""
**Required Inventory Ledger columns:**

| Column | Used for |
|---|---|
| `Date` | Daily dedup — keep latest snapshot per SKU/FC |
| `MSKU` | SKU identifier |
| `FNSKU` | Amazon barcode (for flat file) |
| `Title` | Product name |
| `Disposition` | Sellable vs damaged |
| `Ending Warehouse Balance` | Current stock units |
| `Customer Shipments` | Sales — negative = sold |
| `Customer Returns` | Returns subtracted from sales |
| `Location` | FC code (e.g. LKO1) |

> **Why Customer Shipments is correct:** It captures exactly what was dispatched each day.
> Ending Warehouse Balance is the snapshot — do NOT sum it across rows.
""")
    fc_ref = pd.DataFrame([{"FC Code":k,"FC Name":v[0],"City":v[1],
                             "State":v[2],"Cluster":v[3]}
                            for k,v in FC_MASTER.items()])
    sdf(fc_ref, h=300)

if not inv_files:
    al("👆 Upload your <b>Inventory Ledger</b> file above to begin.", "b")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
#  PARSE INVENTORY LEDGER
# ══════════════════════════════════════════════════════════════════════════════
parts = [read_file(f) for f in inv_files]
parts = [d for d in parts if not d.empty]
if not parts:
    st.error("❌ Could not read any Inventory Ledger file."); st.stop()
for i in range(len(parts)):
    parts[i].columns = parts[i].columns.str.strip()
raw = pd.concat(parts, ignore_index=True)

# Column detection
CD    = gcol(raw, ["Date"])
CMSKU = gcol(raw, ["MSKU","Sku","SKU"])
CFNSK = gcol(raw, ["FNSKU"])
CASIN = gcol(raw, ["ASIN"])
CTITL = gcol(raw, ["Title","Product Name","Description"])
CDISP = gcol(raw, ["Disposition"])
CEND  = gcol(raw, ["Ending Warehouse Balance","Quantity","Qty"])
CLOC  = gcol(raw, ["Location","FC Code","Warehouse Code"])
CSHIP = gcol(raw, ["Customer Shipments"])
CRET  = gcol(raw, ["Customer Returns"])

for lbl, c in [("MSKU",CMSKU),("Ending Warehouse Balance",CEND),("Location",CLOC)]:
    if not c:
        st.error(f"❌ Column not found: **{lbl}**  |  Columns found: {list(raw.columns)}")
        st.stop()

# Build clean frame with FIXED column names used everywhere
inv = pd.DataFrame({
    "MSKU":    raw[CMSKU].astype(str).str.strip(),
    "FNSKU":   raw[CFNSK].astype(str).str.strip() if CFNSK else "",
    "ASIN":    raw[CASIN].astype(str).str.strip()  if CASIN else "",
    "Title":   raw[CTITL].astype(str).str.strip()  if CTITL else "",
    "Disp":    (raw[CDISP].astype(str).str.upper().str.strip() if CDISP else "SELLABLE"),
    "Stock":   pd.to_numeric(raw[CEND],  errors="coerce").fillna(0),
    "FC Code": raw[CLOC].astype(str).str.upper().str.strip(),
})
n_raw = len(inv)

# DAILY DEDUP — ledger has 1 row per SKU/FC/Disp per day; keep LATEST only
if CD:
    inv["_dt"] = parse_dates(raw[CD].astype(str))
    last_date  = inv["_dt"].max()
    inv = (inv.sort_values("_dt")
              .groupby(["MSKU","Disp","FC Code"], as_index=False).last()
              .drop(columns=["_dt"]))
    date_lbl = last_date.strftime("%d %b %Y") if pd.notna(last_date) else "Latest"
else:
    inv = inv.groupby(["MSKU","Disp","FC Code"], as_index=False).last()
    date_lbl = "Latest row"
n_dedup = len(inv)

# Enrich FC metadata — all use "FC Code" column
inv["FC Name"]    = inv["FC Code"].apply(fc_name)
inv["FC City"]    = inv["FC Code"].apply(fc_city)
inv["FC State"]   = inv["FC Code"].apply(fc_state)
inv["FC Cluster"] = inv["FC Code"].apply(fc_cluster)

# SKU master
sku_master = (inv[["MSKU","FNSKU","ASIN","Title"]]
    .drop_duplicates("MSKU", keep="last").set_index("MSKU"))

# Sellable / damaged split
sell = inv[inv["Disp"]=="SELLABLE"].copy()
dmg  = inv[inv["Disp"]!="SELLABLE"].copy()
active_fcs = sorted(sell["FC Code"].unique())

# Per-SKU totals
sku_sell = (sell.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock":"Sellable Stock"}))
sku_dmg  = (dmg.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock":"Damaged Stock"}))

# FC-level stock — ALWAYS called fc_long, column ALWAYS "FC Code"
fc_long = (sell
    .groupby(["FC Code","FC Name","FC City","FC State","FC Cluster","MSKU"])["Stock"]
    .sum().reset_index().rename(columns={"Stock":"FC Stock"}))

# FC inventory summary
fc_inv_sum = (fc_long
    .groupby(["FC Code","FC Name","FC City","FC State","FC Cluster"])["FC Stock"]
    .sum().reset_index().rename(columns={"FC Stock":"Total Sellable"})
    .sort_values("Total Sellable", ascending=False))

# Damaged breakdowns
dmg_msku = (dmg.groupby(["MSKU","Disp"])["Stock"].sum().reset_index()
    .rename(columns={"Disp":"Disposition","Stock":"Qty"}))
dmg_msku["Product Name"] = (dmg_msku["MSKU"]
    .map(sku_master["Title"]).fillna("").apply(lambda x: trunc(x,70)))

dmg_fc = (dmg.groupby(["FC Code","FC Name","FC City","MSKU","Disp"])["Stock"].sum().reset_index()
    .rename(columns={"Disp":"Disposition","Stock":"Qty"}))
dmg_fc["Product Name"] = (dmg_fc["MSKU"]
    .map(sku_master["Title"]).fillna("").apply(lambda x: trunc(x,60)))


# ══════════════════════════════════════════════════════════════════════════════
#  PARSE SALES
#  Primary  : Customer Shipments (Inventory Ledger) — always correct
#  Secondary: MTR — adds FBM SKUs or fills gaps
# ══════════════════════════════════════════════════════════════════════════════
ledger_sales = pd.DataFrame()
if CSHIP and CD:
    ls = pd.DataFrame({
        "MSKU":    raw[CMSKU].astype(str).str.strip(),
        "Date":    parse_dates(raw[CD].astype(str)),
        "Shipped": pd.to_numeric(raw[CSHIP], errors="coerce").fillna(0).abs(),
        "Returns": (pd.to_numeric(raw[CRET], errors="coerce").fillna(0).clip(lower=0)
                    if CRET else pd.Series(0, index=raw.index)),
    })
    ls["Qty"]     = (ls["Shipped"] - ls["Returns"]).clip(lower=0)
    ls["Channel"] = "FBA"
    ledger_sales  = ls[ls["Qty"]>0][["MSKU","Date","Qty","Channel"]].dropna(subset=["Date"])

mtr_sales = pd.DataFrame()
if mtr_files:
    mp = [read_file(f) for f in mtr_files]
    mp = [d for d in mp if not d.empty]
    if mp:
        for i in range(len(mp)): mp[i].columns = mp[i].columns.str.strip()
        mr = pd.concat(mp, ignore_index=True)
        CS  = gcol(mr,["Sku","SKU","MSKU","Item SKU"])
        CQ  = gcol(mr,["Quantity","Qty","Units"])
        CDT = gcol(mr,["Shipment Date","Purchase Date","Order Date","Date"])
        CF  = gcol(mr,["Fulfilment","Fulfillment","Fulfilment Channel","Fulfillment Channel"])
        if CS and CQ and CDT:
            FTM = {"AFN":"FBA","FBA":"FBA","AMAZON_FULFILLED":"FBA",
                   "MFN":"FBM","FBM":"FBM","MERCHANT_FULFILLED":"FBM"}
            ms = pd.DataFrame({
                "MSKU":    mr[CS].astype(str).str.strip(),
                "Qty":     pd.to_numeric(mr[CQ], errors="coerce").fillna(0),
                "Date":    parse_dates(mr[CDT].astype(str)),
                "Channel": (mr[CF].astype(str).str.upper().map(FTM).fillna("FBA")
                            if CF else "FBA"),
            })
            mtr_sales = ms[(ms["Qty"]>0)&ms["Date"].notna()][
                ["MSKU","Date","Qty","Channel"]].copy()

if not ledger_sales.empty and not mtr_sales.empty:
    extra = mtr_sales[~mtr_sales["MSKU"].isin(set(ledger_sales["MSKU"]))]
    sales = pd.concat([ledger_sales, extra], ignore_index=True)
    sales_src = "Inventory Ledger + MTR"
elif not ledger_sales.empty:
    sales = ledger_sales.copy()
    sales_src = "Inventory Ledger (Customer Shipments)"
elif not mtr_sales.empty:
    sales = mtr_sales.copy()
    sales_src = "MTR only"
else:
    st.error("❌ No sales data found. Ensure Inventory Ledger has a 'Customer Shipments' column.")
    st.stop()

sales = sales.sort_values("Date").reset_index(drop=True)

s_min    = sales["Date"].min()
s_max    = sales["Date"].max()
all_days = max((s_max - s_min).days+1, 1)
win_d    = 30 if "30" in s_win else 60 if "60" in s_win else 90 if "90" in s_win else all_days
cutoff   = s_max - pd.Timedelta(days=win_d-1)
hist     = sales[sales["Date"]>=cutoff].copy() if win_d<all_days else sales.copy()
h_days   = max((hist["Date"].max()-hist["Date"].min()).days+1, 1)

hist_agg = (hist.groupby(["MSKU","Channel"])["Qty"].sum()
    .reset_index().rename(columns={"Qty":"Sales"}))
full_agg = (sales.groupby(["MSKU","Channel"])["Qty"].sum()
    .reset_index().rename(columns={"Qty":"All Time Sales"}))
daily_std = (hist.groupby(["MSKU","Channel","Date"])["Qty"].sum()
    .reset_index().groupby(["MSKU","Channel"])["Qty"].std()
    .reset_index().rename(columns={"Qty":"Std"}))

mo_trend = (sales.assign(Month=sales["Date"].dt.to_period("M").astype(str))
    .groupby(["Month","Channel"])["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units"}).sort_values("Month"))
wk_trend = (sales.assign(Week=sales["Date"].dt.to_period("W").astype(str))
    .groupby("Week")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units"}).sort_values("Week"))


# ══════════════════════════════════════════════════════════════════════════════
#  BUILD MASTER PLAN  — one row per MSKU / Channel
# ══════════════════════════════════════════════════════════════════════════════
keys = pd.concat([
    hist_agg[["MSKU","Channel"]],
    sku_sell.assign(Channel="FBA")[["MSKU","Channel"]],
], ignore_index=True).drop_duplicates()

plan = keys.copy()
for df_, on_ in [(hist_agg,["MSKU","Channel"]),(full_agg,["MSKU","Channel"]),
                 (daily_std,["MSKU","Channel"]),(sku_sell,"MSKU"),(sku_dmg,"MSKU")]:
    plan = plan.merge(df_, on=on_, how="left")

for c in ["Sales","All Time Sales","Sellable Stock","Damaged Stock","Std"]:
    plan[c] = pd.to_numeric(plan.get(c), errors="coerce").fillna(0)

plan["Product Name"] = plan["MSKU"].map(sku_master["Title"]).fillna("—")
plan["FNSKU"]        = plan["MSKU"].map(sku_master["FNSKU"]).fillna("")
plan["ASIN"]         = plan["MSKU"].map(sku_master["ASIN"]).fillna("")

plan["Avg/Day"]      = (plan["Sales"]/h_days).round(4)
plan["Safety Stock"] = (Z * plan["Std"] * math.sqrt(plan_days)).round(0)
plan["Need"]         = (plan["Avg/Day"]*plan_days + plan["Safety Stock"]).round(0)
plan["Dispatch"]     = (plan["Need"] - plan["Sellable Stock"]).clip(lower=0).round(0)
plan["DOC"]          = np.where(plan["Avg/Day"]>0,
    (plan["Sellable Stock"]/plan["Avg/Day"]).round(1),
    np.where(plan["Sellable Stock"]>0, 9999, 0))
plan["Status"]   = plan.apply(lambda r: health(r["DOC"], plan_days), axis=1)
plan["Velocity"] = plan["Avg/Day"].apply(vel)
_mx = plan["Avg/Day"].max() or 1
plan["Priority"] = plan.apply(lambda r: round(
    (max(0,(plan_days-min(r["DOC"],plan_days))/plan_days)*0.65
     + min(r["Avg/Day"]/_mx,1)*0.35)*100, 1)
    if r["Avg/Day"]>0 else 0, axis=1)

fba_plan = plan[plan["Channel"]=="FBA"].copy()
fbm_plan = plan[plan["Channel"]=="FBM"].copy()

# Display column order
DCOLS = ["MSKU","FNSKU","Product Name","Sellable Stock","Damaged Stock",
         "Avg/Day","DOC","Status","Velocity","Sales","All Time Sales",
         "Need","Dispatch","Priority"]

def plan_table(df, min_d=0, h=480):
    """Render plan table — sort before slicing so Priority never goes missing."""
    d = df.copy()
    if min_d > 0: d = d[d["Dispatch"]>=min_d]
    d = safe_sort(d, ["Priority"], [False])          # sort FIRST
    d["Product Name"] = d["Product Name"].apply(lambda x: trunc(x, 58))
    sdf(d[safe_cols(d, DCOLS)], h=h)                 # THEN slice


# ══════════════════════════════════════════════════════════════════════════════
#  FC-WISE DISPATCH  — column always "FC Code", sort before every slice
# ══════════════════════════════════════════════════════════════════════════════
fcp = fc_long.merge(
    plan[["MSKU","Channel","Avg/Day","Dispatch","Sellable Stock",
          "Need","Product Name","FNSKU","ASIN","Priority","DOC"]],
    on="MSKU", how="left")
for c in ["Avg/Day","Dispatch","Sellable Stock","Need","Priority","DOC"]:
    fcp[c] = pd.to_numeric(fcp.get(c), errors="coerce").fillna(0)
fcp["Channel"] = fcp["Channel"].fillna("FBA")
fcp["FC DOC"]  = np.where(fcp["Avg/Day"]>0,
    (fcp["FC Stock"]/fcp["Avg/Day"]).round(1),
    np.where(fcp["FC Stock"]>0, 9999, 0))
fcp["FC Status"] = fcp.apply(lambda r: health(r["FC DOC"], plan_days), axis=1)

_tot = (fcp.groupby(["MSKU","Channel"])["FC Stock"]
        .sum().reset_index().rename(columns={"FC Stock":"_T"}))
fcp  = fcp.merge(_tot, on=["MSKU","Channel"], how="left")
fcp["FC Share"]    = (fcp["FC Stock"]/fcp["_T"].replace(0,1)).fillna(1.0)
fcp["FC Dispatch"] = (fcp["Dispatch"]*fcp["FC Share"]).round(0).astype(int)
fcp.drop(columns=["_T"], inplace=True)

# FC dispatch summary — 1 row per FC
fc_disp_sum = (fcp
    .groupby(["FC Code","FC Name","FC City","FC State","FC Cluster"])
    .agg(Total_Stock=("FC Stock","sum"), Dispatch=("FC Dispatch","sum"),
         SKUs=("MSKU","nunique"), Avg_DOC=("FC DOC","mean"))
    .reset_index()
    .rename(columns={"Total_Stock":"Total Stock","Dispatch":"Units to Dispatch","Avg_DOC":"Avg DOC"})
    .sort_values("Units to Dispatch", ascending=False))
fc_disp_sum["Avg DOC"] = fc_disp_sum["Avg DOC"].round(1)

# Risk tables
r_crit   = plan[(plan["DOC"]<14)    & (plan["Avg/Day"]>0)].copy()
r_dead   = plan[(plan["Avg/Day"]==0)& (plan["Sellable Stock"]>0)].copy()
r_slow   = plan[(plan["Avg/Day"]>0) & (plan["DOC"]>90)].copy()
r_excess = plan[plan["DOC"]>plan_days*2].copy()


# ══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("### 📊 Dashboard")
m = st.columns(7)
tot_sell = int(sku_sell["Sellable Stock"].sum())
tot_dmg  = int(sku_dmg["Damaged Stock"].sum()) if not sku_dmg.empty else 0
tot_disp = int(fba_plan["Dispatch"].sum())
tot_sold = int(sales["Qty"].sum())
n_skus   = plan["MSKU"].nunique()
n_fcs    = len(active_fcs)
avg_doc  = plan[plan["Avg/Day"]>0]["DOC"].replace(9999,np.nan).mean()
m[0].metric("Total SKUs",         fmt(n_skus))
m[1].metric("Active FCs",         fmt(n_fcs))
m[2].metric("Sellable Units",     fmt(tot_sell))
m[3].metric("Units Sold",         fmt(tot_sold))
m[4].metric("Need to Dispatch",   fmt(tot_disp))
m[5].metric("🔴 Critical SKUs",   fmt(len(r_crit)))
m[6].metric("Avg Days of Cover",  f"{avg_doc:.0f}d" if pd.notna(avg_doc) else "N/A")

al(f"✅ Stock as of <b>{date_lbl}</b>  |  Source: <b>{sales_src}</b>  |  "
   f"Window: <b>{s_win}</b> ({h_days}d)  |  FCs active: {', '.join(active_fcs)}", "g")
if r_crit.empty is False:
    al(f"🚨 <b>{len(r_crit)} SKU(s) CRITICAL</b> — under 14 days stock. Dispatch now!", "r")
if r_dead.empty is False:
    al(f"⚠️ <b>{len(r_dead)} SKU(s)</b> have stock with ZERO sales — fix listing or remove.", "a")
if tot_dmg:
    al(f"🔧 <b>{fmt(tot_dmg)} damaged units</b> — raise Removal Order / Reimbursement Claim.", "a")
if r_excess.empty is False:
    al(f"📦 <b>{len(r_excess)} SKU(s)</b> are overstocked (>{plan_days*2}d cover) — pause restock.", "y")
st.divider()


# ══════════════════════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════════════════════
T = st.tabs([
    "📋 Full Plan",
    "📦 Dispatch Needed",
    "🏭 FC-Wise View",
    "🚨 Alerts & Dead Stock",
    "📈 Sales Trends",
    "📄 Bulk Shipment File",
    "📥 Download Reports",
])


# ─────────────────────────────────────────────────────────────────────────────
# TAB 0 — FULL PLAN
# ─────────────────────────────────────────────────────────────────────────────
with T[0]:
    st.subheader("📋 Full Inventory Plan — All SKUs")
    st.caption(f"{n_skus} SKUs · {plan_days}-day plan · {svc} SL · {s_win} ({h_days}d)")

    fa,fb,fc_,fd = st.columns(4)
    with fa: f_st = st.multiselect("Status",
        ["🔴 Critical","🟠 Low","🟢 Healthy","🟡 Excess","🔵 Overstocked","⚫ No Stock"],key="fs")
    with fb: f_vl = st.multiselect("Velocity",
        ["🔥 Hot","🟢 Fast","🟡 Medium","🔵 Slow","⚫ Dead"],key="fv")
    with fc_: f_ch = st.multiselect("Channel",["FBA","FBM"],key="fch")
    with fd:  f_do = st.checkbox("Dispatch-needed only",key="fd")

    v = plan.copy()
    if f_st: v = v[v["Status"].isin(f_st)]
    if f_vl: v = v[v["Velocity"].isin(f_vl)]
    if f_ch: v = v[v["Channel"].isin(f_ch)]
    if f_do: v = v[v["Dispatch"]>0]
    st.caption(f"Showing **{len(v)}** of {n_skus} SKUs")
    plan_table(v, h=530)

    st.divider()
    st.subheader("🔍 SKU Detail")
    sel = st.selectbox("Select MSKU", sorted(plan["MSKU"].unique()), key="d0")
    if sel:
        row = plan[plan["MSKU"]==sel].iloc[0]
        st.markdown(f"**{sel}** — {trunc(row.get('Product Name',''), 90)}")
        st.caption(f"FNSKU: `{row['FNSKU']}`  |  ASIN: `{row['ASIN']}`")
        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("Sellable",  fmt(row["Sellable Stock"]))
        c2.metric("Damaged",   fmt(row["Damaged Stock"]))
        c3.metric("Avg/Day",   f"{row['Avg/Day']:.3f}")
        c4.metric("DOC",       f"{row['DOC']:.0f}d" if row["DOC"]<9999 else "∞")
        c5.metric("Dispatch",  fmt(row["Dispatch"]))
        c6.metric("Status",    row["Status"])
        fcd = fc_long[fc_long["MSKU"]==sel].copy()
        if not fcd.empty:
            fcd["FC DOC"] = np.where(row["Avg/Day"]>0,
                (fcd["FC Stock"]/row["Avg/Day"]).round(1), 9999)
            fcd["FC Status"] = fcd["FC DOC"].apply(lambda d: health(d, plan_days))
            sdf(fcd[["FC Code","FC Name","FC City","FC Cluster","FC Stock","FC DOC","FC Status"]])
        else:
            st.info("No stock at any FC for this SKU.")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 — DISPATCH NEEDED
# ─────────────────────────────────────────────────────────────────────────────
with T[1]:
    st.subheader("📦 SKUs That Need Dispatch")
    need = fba_plan[fba_plan["Dispatch"]>= max(1, min_disp)].copy()
    if need.empty:
        al("✅ All FBA SKUs are well-stocked. No dispatch required.", "g")
    else:
        st.caption(f"{len(need)} SKUs · {fmt(int(need['Dispatch'].sum()))} total units")
        plan_table(need, h=440)

        st.divider()
        sh("🏭 FC-Wise Dispatch Breakdown")
        st.caption("Units allocated to each FC proportional to current stock share")
        # sort BEFORE slicing
        fc_need = (fcp[(fcp["Channel"]=="FBA") & (fcp["FC Dispatch"]>0)]
                   .copy())
        fc_need = safe_sort(fc_need, ["FC Code","Priority"], [True,False])
        fc_need["Product Name"] = fc_need["Product Name"].apply(lambda x: trunc(x,50))
        sdf(fc_need[safe_cols(fc_need,
            ["FC Code","FC Name","FC City","FC Cluster",
             "MSKU","FNSKU","Product Name","FC Stock","Avg/Day","FC DOC","FC Dispatch"])], h=440)

    if not fbm_plan.empty and fbm_plan["Dispatch"].sum()>0:
        st.divider()
        st.subheader("📮 FBM Dispatch")
        plan_table(fbm_plan[fbm_plan["Dispatch"]>0], h=300)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 — FC-WISE VIEW
# ─────────────────────────────────────────────────────────────────────────────
with T[2]:
    st.subheader("🏭 Fulfillment Center View")
    sh("FC Summary — Stock & Dispatch")
    sdf(fc_disp_sum, h=300)

    st.divider()
    sh("Drill into a Specific FC")
    if active_fcs:
        fc_opts = [f"{fc} — {fc_name(fc)}" for fc in active_fcs]
        sel_fc  = st.selectbox("Choose FC", fc_opts, key="fc2").split(" — ")[0].strip()
        fc_skus = safe_sort(fcp[fcp["FC Code"]==sel_fc].copy(), ["Priority"], [False])
        fc_skus["Product Name"] = fc_skus["Product Name"].apply(lambda x: trunc(x,50))
        ci1,ci2,ci3,ci4 = st.columns(4)
        ci1.metric("FC Code",         sel_fc)
        ci2.metric("Total Stock",     fmt(int(fc_skus["FC Stock"].sum())))
        ci3.metric("To Dispatch",     fmt(int(fc_skus["FC Dispatch"].sum())))
        ci4.metric("City",            fc_city(sel_fc))
        st.caption(f"{fc_name(sel_fc)} · {fc_state(sel_fc)} · Cluster: {fc_cluster(sel_fc)}")
        sdf(fc_skus[safe_cols(fc_skus,
            ["FC Code","MSKU","FNSKU","Product Name","FC Stock",
             "Avg/Day","FC DOC","FC Status","FC Dispatch"])], h=420)

    st.divider()
    sh("All FC Inventory")
    sdf(fc_inv_sum, h=320)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 — ALERTS & DEAD STOCK
# ─────────────────────────────────────────────────────────────────────────────
with T[3]:
    st.subheader("🚨 Inventory Health Alerts")

    at0,at1,at2,at3,at4 = st.tabs([
        f"🔴 Critical ({len(r_crit)})",
        f"⚫ Dead Stock ({len(r_dead)})",
        f"🟡 Slow Moving ({len(r_slow)})",
        f"🔵 Excess ({len(r_excess)})",
        "🔧 Damaged",
    ])

    def risk_view(df, sort_col, asc=True):
        d = safe_sort(df.copy(), [sort_col], [asc])
        d["Product Name"] = d["Product Name"].apply(lambda x: trunc(x,58))
        sdf(d[safe_cols(d, ["MSKU","FNSKU","Product Name","Sellable Stock",
                             "Damaged Stock","Avg/Day","DOC","Status",
                             "All Time Sales","Dispatch"])])

    with at0:
        st.caption("Stock < 14 days — dispatch immediately")
        al("✅ No critical SKUs!", "g") if r_crit.empty else risk_view(r_crit, "DOC", True)

    with at1:
        st.caption("Stock sitting idle with zero sales — fix listing or plan removal")
        if r_dead.empty:
            al("✅ No dead stock.", "g")
        else:
            d2 = r_dead.copy()
            d2["Product Name"] = d2["Product Name"].apply(lambda x: trunc(x,58))
            sdf(d2[safe_cols(d2, ["MSKU","FNSKU","ASIN","Product Name",
                                   "Sellable Stock","Damaged Stock","All Time Sales"])])
            st.divider()
            sh("📍 Dead Stock Location — which FC is holding it?")
            dl = fc_long[fc_long["MSKU"].isin(r_dead["MSKU"])].copy()
            dl["Product Name"] = dl["MSKU"].map(sku_master["Title"]).fillna("").apply(lambda x: trunc(str(x),50))
            dl = safe_sort(dl, ["FC Stock"], [False])
            sdf(dl[safe_cols(dl, ["FC Code","FC Name","FC City","FC Cluster",
                                   "MSKU","Product Name","FC Stock"])], h=300)

    with at2:
        st.caption("Days of cover > 90 — consider pausing replenishment")
        al("✅ No slow-moving SKUs.", "g") if r_slow.empty else risk_view(r_slow, "DOC", False)

    with at3:
        st.caption(f"DOC > {plan_days*2} days — significantly overstocked")
        al("✅ No overstocked SKUs.", "g") if r_excess.empty else risk_view(r_excess, "DOC", False)

    with at4:
        st.caption("Raise Removal Order or Reimbursement Claim in Seller Central")
        if dmg_msku.empty:
            al("✅ No damaged stock.", "g")
        else:
            sdf(safe_sort(dmg_msku, ["Qty"], [False]), h=280)
            if not dmg_fc.empty:
                sh("Damaged Stock by FC")
                sdf(safe_sort(dmg_fc, ["Qty"], [False]), h=250)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 — SALES TRENDS
# ─────────────────────────────────────────────────────────────────────────────
with T[4]:
    st.subheader("📈 Sales Trends & Analytics")
    st.caption(f"Range: {s_min.strftime('%d %b %Y')} → {s_max.strftime('%d %b %Y')} "
               f"({all_days}d) · Window: {s_win} ({h_days}d)")

    s1,s2 = st.columns(2)
    with s1:
        sh("Weekly Sales")
        st.line_chart(wk_trend.set_index("Week")["Units"] if not wk_trend.empty else pd.Series(), height=220)
    with s2:
        sh("Monthly Sales by Channel")
        if not mo_trend.empty:
            pm = mo_trend.pivot_table(index="Month",columns="Channel",values="Units",
                                       aggfunc="sum",fill_value=0)
            st.bar_chart(pm, height=220)

    st.divider()
    sh("Top 20 Fastest-Moving SKUs")
    top20 = plan[plan["Avg/Day"]>0].nlargest(20,"Avg/Day").copy()
    top20 = safe_sort(top20, ["Avg/Day"], [False])
    top20["Product Name"] = top20["Product Name"].apply(lambda x: trunc(x,55))
    sdf(top20[safe_cols(top20, ["MSKU","FNSKU","Product Name","Avg/Day",
                                 "Velocity","Sellable Stock","DOC","Status","All Time Sales"])], h=400)

    st.divider()
    sh("SKU Sales Drill-Down")
    spick = st.selectbox("Select MSKU", sorted(plan["MSKU"].unique()), key="spick")
    if spick:
        pn_ = trunc(sku_master["Title"].get(spick,""), 80)
        st.caption(f"**{spick}** — {pn_}")
        sd = sales[sales["MSKU"]==spick].groupby("Date")["Qty"].sum().reset_index().set_index("Date")
        st.line_chart(sd["Qty"] if not sd.empty else pd.Series(), height=200)


# ─────────────────────────────────────────────────────────────────────────────
# TAB 5 — BULK SHIPMENT FILE
# ─────────────────────────────────────────────────────────────────────────────
with T[5]:
    st.subheader("📄 Amazon FBA Bulk Shipment Upload File")
    al("📋 <b>How to use:</b>  1. Edit quantities below if needed  →  "
       "2. Download file  →  3. In Seller Central: <b>Send to Amazon → "
       "Create New Shipment → Upload a file</b>  →  "
       "4. Amazon auto-creates one shipment per FC ✅", "b")

    bf1,bf2 = st.columns(2)
    with bf1:
        shp_ref  = st.text_input("Shipment Reference",
            value=f"Dispatch_{datetime.now().strftime('%d%b%Y')}", key="sref")
        shp_date = st.date_input("Expected Ship Date",
            value=datetime.now().date()+timedelta(days=2), key="sdt")
    with bf2:
        sel_fcs = st.multiselect("Include FCs (blank = all)", active_fcs, key="sfcs")
        split_fc = st.checkbox("Generate separate file per FC", key="spl")

    disp_rows = fcp[(fcp["Channel"]=="FBA") & (fcp["FC Dispatch"]>0)].copy()
    if sel_fcs:
        disp_rows = disp_rows[disp_rows["FC Code"].isin(sel_fcs)]

    if disp_rows.empty:
        al("✅ No FBA dispatch required. All stock levels are adequate.", "g")
    else:
        disp_rows["Units to Ship"] = disp_rows["FC Dispatch"].astype(int)
        disp_rows["Cases"]         = np.ceil(disp_rows["Units to Ship"]/case_qty).astype(int)
        disp_rows["Product Name"]  = disp_rows["Product Name"].apply(lambda x: trunc(x,60))
        # sort BEFORE selecting columns
        disp_rows = safe_sort(disp_rows, ["FC Code","Priority"], [True,False])
        edit_cols = ["FC Code","FC Name","MSKU","FNSKU","Product Name","Units to Ship","Cases"]
        st.caption(f"**{len(disp_rows)} lines** · "
                   f"**{fmt(disp_rows['Units to Ship'].sum())} units** · "
                   f"**{disp_rows['FC Code'].nunique()} FCs** · "
                   "Edit quantities if needed:")
        edited = st.data_editor(
            disp_rows[safe_cols(disp_rows, edit_cols)].reset_index(drop=True),
            num_rows="dynamic", use_container_width=True, key="edit_disp")

        def make_flat(rows, fc_code, shp_name):
            lines = ["TemplateType=FlatFileShipmentCreation\tVersion=2015.0403",""]
            lines += ["\t".join(["ShipmentName","ShipFromName","ShipFromAddressLine1",
                                  "ShipFromCity","ShipFromStateOrProvinceCode",
                                  "ShipFromPostalCode","ShipFromCountryCode",
                                  "ShipmentStatus","LabelPrepType",
                                  "AreCasesRequired","DestinationFulfillmentCenterId"]),
                      "\t".join([shp_name, wh_name or "My Warehouse",
                                  wh_addr or "Warehouse", wh_city or "City",
                                  (wh_state or "UP").upper()[:2],
                                  wh_pin or "000000", "IN", "WORKING",
                                  lbl_own, "YES" if case_packed else "NO", fc_code]),
                      ""]
            lines.append("\t".join(["SellerSKU","FNSKU","QuantityShipped","QuantityInCase",
                                     "PrepOwner","LabelOwner","ItemDescription","ExpectedDeliveryDate"]))
            for _, r in rows.iterrows():
                qty = int(r.get("Units to Ship",0))
                if qty<=0: continue
                lines.append("\t".join([
                    str(r.get("MSKU","")), str(r.get("FNSKU","")), str(qty),
                    str(case_qty) if case_packed else "",
                    prep_own, lbl_own,
                    str(r.get("Product Name",""))[:200], str(shp_date)]))
            return "\n".join(lines)

        st.divider()
        fc_col = "FC Code" if "FC Code" in edited.columns else "FC"
        fcs_in = sorted(edited[fc_col].unique()) if fc_col in edited.columns else []

        if not split_fc:
            blocks = []
            for fc in fcs_in:
                rows = edited[edited[fc_col]==fc]
                if not rows.empty:
                    blocks.append(make_flat(rows, fc, f"{shp_ref}_{fc}"))
            combined = "\n\n".join(blocks)
            dl1,dl2 = st.columns(2)
            with dl1:
                st.download_button(
                    f"📥 Download Bulk Shipment File ({len(fcs_in)} FCs, "
                    f"{fmt(edited['Units to Ship'].sum())} units)",
                    combined.encode("utf-8"),
                    f"FBA_Bulk_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                    "text/plain", use_container_width=True)
            with dl2:
                st.download_button("📋 Dispatch List CSV",
                    edited.to_csv(index=False),
                    f"Dispatch_{datetime.now().strftime('%Y%m%d')}.csv",
                    "text/csv", use_container_width=True)
            with st.expander("👁️ Preview flat file (first 3000 chars)"):
                st.code(combined[:3000]+("\n...(truncated)" if len(combined)>3000 else ""))
        else:
            st.markdown(f"**{len(fcs_in)} separate files:**")
            dcols_ = st.columns(min(len(fcs_in),4))
            for i,fc in enumerate(fcs_in):
                rows = edited[edited[fc_col]==fc]
                if rows.empty: continue
                txt = make_flat(rows, fc, f"{shp_ref}_{fc}")
                with dcols_[i%4]:
                    st.download_button(
                        f"📥 {fc}\n{fc_name(fc)}\n({fmt(rows['Units to Ship'].sum())} units)",
                        txt.encode("utf-8"),
                        f"FBA_{fc}_{datetime.now().strftime('%Y%m%d')}.txt",
                        "text/plain", key=f"dl_{fc}")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 6 — DOWNLOAD REPORTS
# ─────────────────────────────────────────────────────────────────────────────
with T[6]:
    st.subheader("📥 Download Full Reports")

    def build_excel():
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:

            # Helper: sort + slice + write — NEVER sort after column slice
            def ws(df, sort_cols, sort_asc, out_cols, sheet, freeze=1):
                d = safe_sort(df.copy(), sort_cols, sort_asc)
                d["Product Name"] = d["Product Name"].apply(lambda x: trunc(str(x),90)) \
                    if "Product Name" in d.columns else d.get("Product Name","")
                xl(w, d[safe_cols(d, out_cols)], sheet, freeze)

            # 1. Full Plan
            ws(plan, ["Priority"], [False], DCOLS, "Full Plan")

            # 2. FBA Dispatch Needed
            ws(fba_plan[fba_plan["Dispatch"]>0], ["Priority"], [False],
               DCOLS, "FBA Dispatch Needed")

            # 3. FC-Wise Dispatch (sort BEFORE slicing)
            _fco = fcp[fcp["Channel"]=="FBA"].copy()
            _fco = safe_sort(_fco, ["FC Code","Priority"], [True,False])
            _fco["Product Name"] = _fco["Product Name"].apply(lambda x: trunc(str(x),90))
            xl(w, _fco[safe_cols(_fco, ["FC Code","FC Name","FC City","FC State","FC Cluster",
                                         "MSKU","FNSKU","Product Name","FC Stock",
                                         "Avg/Day","FC DOC","FC Status","FC Dispatch"])],
               "FC-Wise Dispatch", freeze=2)

            # 4. FC Summary
            xl(w, fc_disp_sum, "FC Summary")

            # 5. Critical SKUs
            ws(r_crit, ["DOC"], [True], DCOLS, "Critical SKUs")

            # 6. Dead Stock
            _dd = r_dead.copy()
            _dd["Product Name"] = _dd["Product Name"].apply(lambda x: trunc(str(x),90))
            xl(w, _dd[safe_cols(_dd, ["MSKU","FNSKU","ASIN","Product Name",
                                       "Sellable Stock","Damaged Stock","All Time Sales"])],
               "Dead Stock")

            # 7. Dead Stock by FC
            _dl = fc_long[fc_long["MSKU"].isin(r_dead["MSKU"])].copy()
            _dl["Product Name"] = _dl["MSKU"].map(sku_master["Title"]).fillna("").apply(lambda x: trunc(str(x),90))
            _dl = safe_sort(_dl, ["FC Stock"], [False])
            xl(w, _dl[safe_cols(_dl, ["FC Code","FC Name","FC City","FC Cluster",
                                       "MSKU","Product Name","FC Stock"])],
               "Dead Stock by FC", freeze=2)

            # 8. Slow Moving
            ws(r_slow, ["DOC"], [False], DCOLS, "Slow Moving")

            # 9. Excess Stock
            ws(r_excess, ["DOC"], [False], DCOLS, "Excess Stock")

            # 10. Damaged
            if not dmg_msku.empty:
                xl(w, safe_sort(dmg_msku, ["Qty"], [False]), "Damaged Stock")

            # 11. Monthly Trend
            xl(w, mo_trend, "Monthly Trend")

            # 12. Weekly Trend
            xl(w, wk_trend, "Weekly Trend")

            # 13. SKU Master
            _sm = sku_master.reset_index().copy()
            _sm["Title"] = _sm["Title"].apply(lambda x: trunc(str(x),120))
            xl(w, _sm[safe_cols(_sm, ["MSKU","FNSKU","ASIN","Title"])], "SKU Master")

            # 14. FC Master Reference
            xl(w, pd.DataFrame([
                {"FC Code":k,"FC Name":v[0],"City":v[1],"State":v[2],"Cluster":v[3]}
                for k,v in FC_MASTER.items()
            ]), "FC Master")

        buf.seek(0); return buf

    d1,d2,d3 = st.columns(3)
    with d1:
        st.download_button(
            "📗 Full Excel Report\n(14 sheets — complete data)",
            data=build_excel(),
            file_name=f"FBA_Plan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with d2:
        # sort BEFORE slice for CSV
        _c1 = safe_sort(fba_plan[fba_plan["Dispatch"]>0].copy(), ["Priority"], [False])
        _c1["Product Name"] = _c1["Product Name"].apply(lambda x: trunc(str(x),90))
        st.download_button(
            "📋 Dispatch Plan CSV\n(FBA SKUs needing restock)",
            data=_c1[safe_cols(_c1, DCOLS)].to_csv(index=False),
            file_name=f"Dispatch_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv", use_container_width=True)
    with d3:
        # sort BEFORE slice for CSV — THIS was the crashing line
        _c2 = fcp[(fcp["Channel"]=="FBA") & (fcp["FC Dispatch"]>0)].copy()
        _c2 = safe_sort(_c2, ["FC Code","Priority"], [True,False])  # sort first
        _c2["Product Name"] = _c2["Product Name"].apply(lambda x: trunc(str(x),90))
        OUT_COLS = ["FC Code","FC Name","FC City","MSKU","FNSKU",
                    "Product Name","FC Stock","Avg/Day","FC DOC","FC Dispatch"]
        st.download_button(
            "🏭 FC-Wise Dispatch CSV\n(1 row per SKU per FC)",
            data=_c2[safe_cols(_c2, OUT_COLS)].to_csv(index=False),
            file_name=f"FC_Dispatch_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv", use_container_width=True)

    st.divider()
    st.markdown("""**📑 Report contents:**

| Sheet | Contents |
|---|---|
| **Full Plan** | All SKUs — MSKU, Product Name, stock, avg daily sale, dispatch, status |
| **FBA Dispatch Needed** | Only SKUs needing restock, sorted by priority |
| **FC-Wise Dispatch** | 1 row per SKU per FC — units to each location |
| **FC Summary** | 1 row per FC — total stock, units to dispatch, avg DOC |
| **Critical SKUs** | Stock < 14 days |
| **Dead Stock** | Stock with zero sales |
| **Dead Stock by FC** | Dead stock with exact FC location |
| **Slow Moving** | Sales exist but > 90 days cover |
| **Excess Stock** | DOC > 2× plan days |
| **Damaged Stock** | Non-sellable units by MSKU |
| **Monthly / Weekly Trend** | Sales trend data |
| **SKU Master** | MSKU → FNSKU → ASIN → Product Name |
| **FC Master** | All 75 Amazon India FC codes |
""")
