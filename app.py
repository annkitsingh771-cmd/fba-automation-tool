"""
FBA Smart Supply Planner — Amazon India
=======================================
Upload: Inventory Ledger (required) + MTR (optional)
• Product Name & SKU always shown
• Sales from Customer Shipments column (correct source)
• Stock from latest Ending Warehouse Balance per SKU/FC
• Bulk FBA Shipment Upload Sheet (one file, all FCs, Amazon format)
• Dead stock, slow movers, dispatch plan — all in one view
"""

import io, math, zipfile
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="FBA Supply Planner", layout="wide", page_icon="📦")
st.markdown("""<style>
[data-testid="stMetricValue"]{font-size:1.6rem;font-weight:700}
.pill{display:inline-block;padding:2px 10px;border-radius:12px;font-size:.8rem;font-weight:700;margin:1px}
.red{background:#fed7d7;color:#742a2a}  .orange{background:#feebc8;color:#7b341e}
.green{background:#c6f6d5;color:#1c4532} .blue{background:#bee3f8;color:#1a365d}
.grey{background:#e2e8f0;color:#2d3748}  .purple{background:#e9d8fd;color:#44337a}
.banner{padding:.5rem 1rem;border-radius:8px;font-weight:600;font-size:.9rem;margin:.3rem 0}
.ban-g{background:#f0fff4;border-left:4px solid #38a169;color:#1c4532}
.ban-r{background:#fff5f5;border-left:4px solid #e53e3e;color:#742a2a}
.ban-y{background:#fffff0;border-left:4px solid #d69e2e;color:#744210}
.stTabs [data-baseweb="tab"]{font-size:13px;font-weight:600}
</style>""", unsafe_allow_html=True)

st.title("📦 FBA Smart Supply Planner — Amazon India")

# ─────────────────────────────────────────────────────────────────────────────
# FC MASTER  (75 verified Amazon India FC codes)
# (FC Name, City, State, Cluster)
# ─────────────────────────────────────────────────────────────────────────────
FC = {
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
    "BLR10":("Kudlu Gate FC","Kudlu Gate","KARNATAKA","Bangalore"),
    "BLR12":("Attibele FC","Attibele","KARNATAKA","Bangalore"),
    "BLR13":("Jigani FC","Jigani","KARNATAKA","Bangalore"),
    "BLR14":("Anekal FC","Anekal","KARNATAKA","Bangalore"),
    "SCJA":("Bangalore SCJA FC","Bangalore","KARNATAKA","Bangalore"),
    "XSAB":("Bangalore XS FC","Bangalore","KARNATAKA","Bangalore"),
    "SBLA":("Bangalore SBLA FC","Bangalore","KARNATAKA","Bangalore"),
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
    "MAA1":("Irungattukottai FC","Irungattukottai","TAMIL NADU","Chennai"),
    "MAA2":("Ponneri FC","Ponneri","TAMIL NADU","Chennai"),
    "MAA3":("Sriperumbudur FC","Sriperumbudur","TAMIL NADU","Chennai"),
    "MAA4":("Ambattur FC","Ambattur","TAMIL NADU","Chennai"),
    "MAA5":("Kanchipuram FC","Kanchipuram","TAMIL NADU","Chennai"),
    "CJB1":("Coimbatore FC","Coimbatore","TAMIL NADU","Chennai"),
    "SMAB":("Chennai SMAB FC","Chennai","TAMIL NADU","Chennai"),
    "COK1":("Kochi FC","Kochi","KERALA","South Kerala"),
    "HYD3":("Shamshabad FC","Shamshabad","TELANGANA","Hyderabad"),
    "HYD6":("Kothur FC","Kothur","TELANGANA","Hyderabad"),
    "HYD7":("Medchal FC","Medchal","TELANGANA","Hyderabad"),
    "HYD8":("Shamshabad FC 2","Shamshabad","TELANGANA","Hyderabad"),
    "HYD8_HYD3":("Shamshabad Mamidipally FC","Shamshabad","TELANGANA","Hyderabad"),
    "HYD9":("Pedda Amberpet FC","Pedda Amberpet","TELANGANA","Hyderabad"),
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

def fc_name(c):    return FC.get(str(c).upper(), (str(c),))[0]
def fc_city(c):    return FC.get(str(c).upper(), ("","Unknown"))[1]
def fc_state(c):   return FC.get(str(c).upper(), ("","","Unknown"))[2]
def fc_cluster(c): return FC.get(str(c).upper(), ("","","","Other"))[3]


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
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
        st.error(f"Cannot read {f.name}: {e}")
        return pd.DataFrame()

def col(df, names):
    """Find first matching column (case-insensitive)."""
    low = {c.strip().lower(): c for c in df.columns}
    for n in names:
        if n.lower() in low:
            return low[n.lower()]
    return ""

def trunc(s, n=60):
    s = str(s)
    return s[:n]+"…" if len(s)>n else s

def n2(x):
    try: return f"{int(x):,}"
    except: return str(x)

def health(doc, plan_days):
    if doc <= 0:           return "⚫ No Stock"
    if doc < 14:           return "🔴 Critical"
    if doc < 30:           return "🟠 Low"
    if doc < plan_days:    return "🟢 Healthy"
    if doc < plan_days*2:  return "🟡 Excess"
    return "🔵 Overstocked"

def vel(avg):
    if avg <= 0:   return "⚫ Dead"
    if avg < 0.5:  return "🔵 Slow"
    if avg < 2:    return "🟡 Medium"
    if avg < 5:    return "🟢 Fast"
    return "🔥 Hot"

def banner(msg, t="g"):
    st.markdown(f'<div class="banner ban-{t}">{msg}</div>', unsafe_allow_html=True)

def styled_table(df, height=400):
    st.dataframe(df, use_container_width=True, hide_index=True, height=height)

def excel_write(writer, df, sheet):
    if df is None or df.empty:
        pd.DataFrame({"Note":["No data"]}).to_excel(writer, sheet_name=sheet, index=False)
        return
    d = df.copy()
    d.to_excel(writer, sheet_name=sheet, index=False, startrow=1)
    wb = writer.book; sh = writer.sheets[sheet]
    hf = wb.add_format({"bold":True,"bg_color":"#1a365d","font_color":"white",
                         "border":1,"align":"center","valign":"vcenter","text_wrap":True})
    for i,c in enumerate(d.columns):
        sh.write(0, i, c, hf)
        mx = d[c].astype(str).str.len().max() if not d.empty else 8
        sh.set_column(i, i, min(max(len(str(c)), mx)+2, 45))
    sh.freeze_panes(2, 0)
    sh.autofilter(0, 0, len(d), len(d.columns)-1)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    plan_days = st.number_input("Planning Days (stock coverage target)", 7, 180, 60)
    svc       = st.selectbox("Service Level", ["95%","90%","98%"])
    Z         = {"90%":1.28,"95%":1.65,"98%":2.05}[svc]
    basis     = st.selectbox("Sales Basis",
                    ["Full History","Last 30 Days","Last 60 Days","Last 90 Days"])
    st.divider()
    st.subheader("🚚 Your Warehouse")
    wh_name = st.text_input("Warehouse / Company Name", placeholder="GLOOYA Warehouse")
    wh_addr = st.text_input("Address Line 1", placeholder="123 Industrial Area")
    wh_city = st.text_input("City", placeholder="Lucknow")
    wh_state= st.text_input("State Code (2 letters)", placeholder="UP", max_chars=2)
    wh_pin  = st.text_input("PIN Code", placeholder="226001")
    st.divider()
    case_qty    = st.number_input("Default Units Per Carton", 1, 500, 12)
    case_packed = st.checkbox("Case-Packed Shipment")
    prep_own    = st.selectbox("Prep Owner",  ["AMAZON","SELLER"])
    lbl_own     = st.selectbox("Label Owner", ["AMAZON","SELLER"])
    st.divider()
    min_units = st.number_input("Min Units to Show in Dispatch", 0, 9999, 0)


# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Upload Your Files")
c1, c2 = st.columns(2)
with c1:
    inv_file = st.file_uploader(
        "**📊 Inventory Ledger** (Required) — CSV/ZIP/XLSX",
        type=["csv","zip","xlsx"], key="inv")
with c2:
    mtr_file = st.file_uploader(
        "**📋 MTR Sales Report** (Optional) — CSV/ZIP/XLSX",
        type=["csv","zip","xlsx"], key="mtr")

with st.expander("ℹ️ Which files to download from Seller Central?"):
    st.markdown("""
| File | Where to download | Why needed |
|---|---|---|
| **Inventory Ledger** | Reports → Fulfillment → Inventory Ledger | Stock + Sales (primary source) |
| **MTR** | Reports → Tax → MTR | Extra sales data (optional) |

**Inventory Ledger must have these columns:**
`Date`, `MSKU`, `FNSKU`, `Title`, `Disposition`, `Ending Warehouse Balance`,
`Customer Shipments` (negative = sold), `Customer Returns`, `Location` (FC code)
""")

if not inv_file:
    st.info("👆 Upload the **Inventory Ledger** to get started.")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# PARSE INVENTORY LEDGER
# ─────────────────────────────────────────────────────────────────────────────
raw = read_file(inv_file)
if raw.empty:
    st.error("Could not read Inventory Ledger. Check file format."); st.stop()
raw.columns = raw.columns.str.strip()

# Column detection — exact Amazon Inventory Ledger names first
C_DATE  = col(raw, ["Date"])
C_MSKU  = col(raw, ["MSKU","Sku","SKU"])
C_FNSKU = col(raw, ["FNSKU"])
C_ASIN  = col(raw, ["ASIN"])
C_TITLE = col(raw, ["Title","Product Name"])
C_DISP  = col(raw, ["Disposition"])
C_END   = col(raw, ["Ending Warehouse Balance","Quantity","Qty"])
C_LOC   = col(raw, ["Location","FC Code","Warehouse Code"])
C_SHIP  = col(raw, ["Customer Shipments"])
C_RET   = col(raw, ["Customer Returns"])

for lbl, c in [("MSKU",C_MSKU),("Ending Warehouse Balance",C_END),("Location",C_LOC)]:
    if not c:
        st.error(f"❌ Missing required column: **{lbl}**  |  Found: {list(raw.columns)}")
        st.stop()

# Build clean working frame
inv = raw.copy()
inv["MSKU"]  = inv[C_MSKU].astype(str).str.strip()
inv["FNSKU"] = inv[C_FNSKU].astype(str).str.strip() if C_FNSKU else ""
inv["ASIN"]  = inv[C_ASIN].astype(str).str.strip()  if C_ASIN  else ""
inv["Title"] = inv[C_TITLE].astype(str).str.strip()  if C_TITLE else ""
inv["Disp"]  = (inv[C_DISP].astype(str).str.upper().str.strip() if C_DISP else "SELLABLE")
inv["Stock"] = pd.to_numeric(inv[C_END], errors="coerce").fillna(0)
inv["FC"]    = inv[C_LOC].astype(str).str.upper().str.strip()

# Parse dates
n_raw = len(inv)
if C_DATE:
    inv["_dt"] = pd.to_datetime(inv[C_DATE], errors="coerce", dayfirst=False)
    inv.loc[inv["_dt"].isna(), "_dt"] = pd.to_datetime(
        inv.loc[inv["_dt"].isna(), C_DATE], errors="coerce", dayfirst=True)
    last_date   = inv["_dt"].max()
    date_label  = last_date.strftime("%d %b %Y") if pd.notna(last_date) else "Unknown"
    # DAILY DEDUP: ledger has 1 row per day — keep only LATEST row per SKU/FC/Disp
    inv = (inv.sort_values("_dt")
              .groupby(["MSKU","Disp","FC"], as_index=False)
              .last()
              .drop(columns=["_dt"]))
else:
    inv = inv.groupby(["MSKU","Disp","FC"], as_index=False).last()
    date_label = "Latest"

n_dedup = len(inv)

# SKU master (MSKU → FNSKU, Title, ASIN)
sku_master = (inv[["MSKU","FNSKU","ASIN","Title"]]
    .drop_duplicates("MSKU")
    .set_index("MSKU"))

# FC enrichment
inv["FC Name"]    = inv["FC"].apply(fc_name)
inv["FC City"]    = inv["FC"].apply(fc_city)
inv["FC State"]   = inv["FC"].apply(fc_state)
inv["FC Cluster"] = inv["FC"].apply(fc_cluster)

# Split sellable / damaged
sell = inv[inv["Disp"]=="SELLABLE"].copy()
dmg  = inv[inv["Disp"]!="SELLABLE"].copy()

# Total sellable per SKU
sku_stock = (sell.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock":"Sellable Stock"}))

# Total damaged per SKU
sku_dmg = (dmg.groupby("MSKU")["Stock"].sum()
    .reset_index().rename(columns={"Stock":"Damaged Stock"}))

# FC-wise sellable stock (long)
fc_long = (sell.groupby(["MSKU","FC","FC Name","FC City","FC State","FC Cluster"])["Stock"]
    .sum().reset_index().rename(columns={"Stock":"FC Stock"}))

# FC inventory summary
fc_summary = (fc_long.groupby(["FC","FC Name","FC City","FC State","FC Cluster"])["FC Stock"]
    .sum().reset_index().rename(columns={"FC Stock":"Total Stock"})
    .sort_values("Total Stock", ascending=False))

# Active FCs
active_fcs = sorted(sell["FC"].unique())


# ─────────────────────────────────────────────────────────────────────────────
# SALES DATA
# Primary: Customer Shipments from Inventory Ledger  (always correct)
# Backup : MTR file (optional, fills in FBM and missing SKUs)
# ─────────────────────────────────────────────────────────────────────────────

# From ledger Customer Shipments column
ledger_sales = pd.DataFrame()
if C_SHIP and C_DATE:
    ls = raw.copy()
    ls["MSKU"]    = ls[C_MSKU].astype(str).str.strip()
    ls["_dt"]     = pd.to_datetime(ls[C_DATE], errors="coerce", dayfirst=False)
    ls.loc[ls["_dt"].isna(),"_dt"] = pd.to_datetime(
        ls.loc[ls["_dt"].isna(), C_DATE], errors="coerce", dayfirst=True)
    ls["Sold"]    = pd.to_numeric(ls[C_SHIP], errors="coerce").fillna(0).abs()
    ls["Returns"] = pd.to_numeric(ls[C_RET], errors="coerce").fillna(0).clip(lower=0) if C_RET else 0
    ls["Qty"]     = (ls["Sold"] - ls["Returns"]).clip(lower=0)
    ls["Channel"] = "FBA"
    ledger_sales  = ls[ls["Qty"]>0][["MSKU","_dt","Qty","Channel"]].rename(columns={"_dt":"Date"}).dropna()

# From MTR file
mtr_sales = pd.DataFrame()
if mtr_file:
    mr = read_file(mtr_file)
    if not mr.empty:
        mr.columns = mr.columns.str.strip()
        CS = col(mr, ["Sku","SKU","MSKU","Item SKU"])
        CQ = col(mr, ["Quantity","Qty","Units"])
        CD = col(mr, ["Shipment Date","Purchase Date","Order Date","Date"])
        CF = col(mr, ["Fulfilment","Fulfillment","Fulfilment Channel","Fulfillment Channel"])
        if CS and CQ and CD:
            FTM = {"AFN":"FBA","FBA":"FBA","AMAZON_FULFILLED":"FBA",
                   "MFN":"FBM","FBM":"FBM","MERCHANT_FULFILLED":"FBM"}
            mr["MSKU"] = mr[CS].astype(str).str.strip()
            mr["Qty"]  = pd.to_numeric(mr[CQ], errors="coerce").fillna(0)
            mr["Date"] = pd.to_datetime(mr[CD], errors="coerce", dayfirst=False)
            mr.loc[mr["Date"].isna(),"Date"] = pd.to_datetime(
                mr.loc[mr["Date"].isna(), CD], errors="coerce", dayfirst=True)
            mr["Channel"] = (mr[CF].astype(str).str.upper().map(FTM).fillna("FBA") if CF else "FBA")
            mtr_sales = mr[(mr["Qty"]>0)&mr["Date"].notna()][["MSKU","Date","Qty","Channel"]].copy()

# Merge: ledger is primary; MTR adds FBM or missing SKUs
if not ledger_sales.empty and not mtr_sales.empty:
    extra = mtr_sales[~mtr_sales["MSKU"].isin(set(ledger_sales["MSKU"]))]
    sales = pd.concat([ledger_sales, extra], ignore_index=True)
    src   = "Inventory Ledger + MTR"
elif not ledger_sales.empty:
    sales = ledger_sales.copy()
    src   = "Inventory Ledger (Customer Shipments)"
elif not mtr_sales.empty:
    sales = mtr_sales.copy()
    src   = "MTR file"
else:
    st.error("No sales data found. Check that your Inventory Ledger has a 'Customer Shipments' column.")
    st.stop()

sales = sales.sort_values("Date").reset_index(drop=True)

# Date range & window
s_min   = sales["Date"].min()
s_max   = sales["Date"].max()
all_days= max((s_max - s_min).days + 1, 1)
win     = (30 if "30" in basis else 60 if "60" in basis else
           90 if "90" in basis else all_days)
cutoff  = s_max - pd.Timedelta(days=win-1)
hist    = sales[sales["Date"]>=cutoff] if win<all_days else sales
h_days  = max((hist["Date"].max()-hist["Date"].min()).days+1, 1)

# Aggregates
by_sku_ch  = (hist.groupby(["MSKU","Channel"])["Qty"].sum()
              .reset_index().rename(columns={"Qty":"Sales"}))
by_sku_all = (sales.groupby(["MSKU","Channel"])["Qty"].sum()
              .reset_index().rename(columns={"Qty":"All Time Sales"}))
daily_std  = (hist.groupby(["MSKU","Channel","Date"])["Qty"].sum()
              .reset_index().groupby(["MSKU","Channel"])["Qty"].std()
              .reset_index().rename(columns={"Qty":"Std"}))

# Monthly / Weekly trend
mo_trend = (sales.assign(Month=sales["Date"].dt.to_period("M").astype(str))
    .groupby(["Month","Channel"])["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units"}).sort_values("Month"))
wk_trend = (sales.assign(Week=sales["Date"].dt.to_period("W").astype(str))
    .groupby("Week")["Qty"].sum().reset_index()
    .rename(columns={"Qty":"Units"}).sort_values("Week"))


# ─────────────────────────────────────────────────────────────────────────────
# BUILD MASTER PLAN TABLE
# Every SKU that has stock OR sales gets a row
# Columns: MSKU | FNSKU | Product Name | [FC cols...] | Sellable Stock |
#          Damaged | Sales | Avg/Day | Safety Stock | Need | Dispatch | DOC | Status
# ─────────────────────────────────────────────────────────────────────────────
# All SKUs
all_skus = pd.concat([
    sku_stock[["MSKU"]],
    by_sku_ch[["MSKU","Channel"]].assign(dummy=1).drop_duplicates()[["MSKU"]],
], ignore_index=True).drop_duplicates()

plan = all_skus.copy()
plan["Channel"] = "FBA"   # default; will update from sales

# Bring in sales by channel
keys = pd.concat([
    by_sku_ch[["MSKU","Channel"]],
    sku_stock.assign(Channel="FBA")[["MSKU","Channel"]],
], ignore_index=True).drop_duplicates()

plan = keys.copy()
for d_, on_ in [
    (by_sku_ch,   ["MSKU","Channel"]),
    (by_sku_all,  ["MSKU","Channel"]),
    (daily_std,   ["MSKU","Channel"]),
    (sku_stock,    "MSKU"),
    (sku_dmg,      "MSKU"),
]:
    plan = plan.merge(d_, on=on_, how="left")

# Fill numbers
for c in ["Sales","All Time Sales","Sellable Stock","Damaged Stock","Std"]:
    plan[c] = pd.to_numeric(plan.get(c), errors="coerce").fillna(0)

# Add FNSKU, ASIN, Product Name from SKU master
plan["Product Name"] = plan["MSKU"].map(sku_master["Title"]).fillna("")
plan["FNSKU"]        = plan["MSKU"].map(sku_master["FNSKU"]).fillna("")
plan["ASIN"]         = plan["MSKU"].map(sku_master["ASIN"]).fillna("")

# Planning maths
plan["Avg/Day"]        = (plan["Sales"] / h_days).round(4)
plan["Safety Stock"]   = (Z * plan["Std"] * math.sqrt(plan_days)).round(0)
plan["Need"]           = (plan["Avg/Day"] * plan_days + plan["Safety Stock"]).round(0)
plan["Dispatch"]       = (plan["Need"] - plan["Sellable Stock"]).clip(lower=0).round(0)
plan["DOC"]            = np.where(
    plan["Avg/Day"]>0,
    (plan["Sellable Stock"]/plan["Avg/Day"]).round(1),
    np.where(plan["Sellable Stock"]>0, 9999, 0))
plan["Status"] = plan.apply(lambda r: health(r["DOC"], plan_days), axis=1)
plan["Velocity"] = plan["Avg/Day"].apply(vel)

mx = plan["Avg/Day"].max() or 1
plan["Priority"] = plan.apply(lambda r: round(
    (max(0,(plan_days-min(r["DOC"],plan_days))/plan_days)*0.65
     + min(r["Avg/Day"]/mx,1)*0.35)*100, 1)
    if r["Avg/Day"]>0 else 0, axis=1)

fba = plan[plan["Channel"]=="FBA"].copy()
fbm = plan[plan["Channel"]=="FBM"].copy()

# Column order for display
BASE_COLS = ["MSKU","FNSKU","Product Name","Channel","Avg/Day","Sellable Stock",
             "Damaged Stock","Sales","Need","Dispatch","DOC","Status","Velocity","Priority"]


def show_plan(df, min_d=0):
    """Show plan table — always MSKU + Product Name visible."""
    d = df.copy()
    if min_d > 0:
        d = d[d["Dispatch"] >= min_d]
    d = d.sort_values("Priority", ascending=False)
    d["Product Name"] = d["Product Name"].apply(lambda x: trunc(x, 55))
    cols = [c for c in BASE_COLS if c in d.columns]
    return d[cols].reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# FC-WISE DISPATCH ALLOCATION
# ─────────────────────────────────────────────────────────────────────────────
fcp = fc_long.merge(
    plan[["MSKU","Channel","Avg/Day","Dispatch","Sellable Stock","Need",
          "Product Name","FNSKU","ASIN","Priority","DOC"]],
    on="MSKU", how="left")

fcp["FC DOC"]     = np.where(fcp["Avg/Day"]>0,
    (fcp["FC Stock"]/fcp["Avg/Day"]).round(1),
    np.where(fcp["FC Stock"]>0, 9999, 0))
fcp["FC Status"]  = fcp.apply(lambda r: health(r["FC DOC"], plan_days), axis=1)

_tot = fcp.groupby(["MSKU","Channel"])["FC Stock"].sum().reset_index().rename(columns={"FC Stock":"_T"})
fcp  = fcp.merge(_tot, on=["MSKU","Channel"], how="left")
fcp["FC Share"]    = (fcp["FC Stock"] / fcp["_T"].replace(0,1)).fillna(1.0)
fcp["FC Dispatch"] = (fcp["Dispatch"] * fcp["FC Share"]).round(0).astype(int)
fcp.drop(columns=["_T"], inplace=True)

# FC dispatch summary
fc_disp = (fcp.groupby(["FC","FC Name","FC City","FC State","FC Cluster"])
    .agg(Total_Stock=("FC Stock","sum"), Units_to_Dispatch=("FC Dispatch","sum"),
         SKUs=("MSKU","nunique"), Avg_DOC=("FC DOC","mean"))
    .reset_index()
    .rename(columns={"Total_Stock":"Total Stock","Units_to_Dispatch":"Units to Dispatch",
                     "Avg_DOC":"Avg DOC"})
    .sort_values("Units to Dispatch", ascending=False))
fc_disp["Avg DOC"] = fc_disp["Avg DOC"].round(1)

# Risk tables
r_crit  = plan[(plan["DOC"]<14)   & (plan["Avg/Day"]>0)].copy()
r_dead  = plan[(plan["Avg/Day"]==0)& (plan["Sellable Stock"]>0)].copy()
r_slow  = plan[(plan["DOC"]>90)   & (plan["Avg/Day"]>0)].copy()
r_excess= plan[plan["DOC"]>plan_days*2].copy()


# ─────────────────────────────────────────────────────────────────────────────
# DASHBOARD METRICS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
tot_sell  = int(sku_stock["Sellable Stock"].sum())
tot_dmg   = int(sku_dmg["Damaged Stock"].sum()) if not sku_dmg.empty else 0
tot_disp  = int(fba["Dispatch"].sum())
tot_sold  = int(sales["Qty"].sum())
n_skus    = plan["MSKU"].nunique()
avg_doc   = plan[plan["Avg/Day"]>0]["DOC"].replace(9999,np.nan).mean()

m = st.columns(7)
m[0].metric("Total SKUs",        n2(n_skus))
m[1].metric("Active FCs",        n2(len(active_fcs)))
m[2].metric("Sellable Stock",    n2(tot_sell))
m[3].metric("Total Sold",        n2(tot_sold))
m[4].metric("Need to Dispatch",  n2(tot_disp))
m[5].metric("🔴 Critical",       n2(len(r_crit)))
m[6].metric("Avg Days of Cover", f"{avg_doc:.0f}d" if pd.notna(avg_doc) else "N/A")

banner(f"✅ Stock as of <b>{date_label}</b> | Sales: <b>{src}</b> | "
       f"{n_dedup} SKU-FC rows (raw: {n_raw}) | FCs active: {', '.join(active_fcs)}", "g")

if len(r_crit):    banner(f"🚨 {len(r_crit)} SKU(s) CRITICAL — stock below 14 days. Dispatch immediately!", "r")
if len(r_dead):    banner(f"⚠️ {len(r_dead)} SKU(s) have stock but ZERO sales — investigate listing.", "y")
if tot_dmg:        banner(f"🔧 {n2(tot_dmg)} damaged/non-sellable units — raise removal/reimbursement.", "y")


# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📋 Full Plan",
    "📦 Need to Dispatch",
    "🚨 Critical & Dead Stock",
    "🏭 FC-Wise View",
    "📈 Sales Trends",
    "📄 Bulk Shipment File",
    "📥 Download Reports",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 0 — FULL PLAN
# Shows ALL SKUs: MSKU, Product Name, stock, sales, dispatch needed, status
# ══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.subheader("📋 Full Inventory Plan — All SKUs")
    st.caption(
        f"{n_skus} SKUs | {plan_days}-day plan | {svc} service level | "
        f"Sales basis: {basis} ({h_days} days) | Source: {src}")

    col1, col2, col3 = st.columns(3)
    with col1:
        filt_status = st.multiselect("Filter by Status",
            ["🔴 Critical","🟠 Low","🟢 Healthy","🟡 Excess","🔵 Overstocked",
             "⚫ No Stock","⚫ Dead"], key="f_status")
    with col2:
        filt_vel = st.multiselect("Filter by Velocity",
            ["🔥 Hot","🟢 Fast","🟡 Medium","🔵 Slow","⚫ Dead"], key="f_vel")
    with col3:
        show_only_dispatch = st.checkbox("Show only SKUs that need dispatch", key="f_disp")

    disp_plan = show_plan(plan)
    if filt_status:
        disp_plan = disp_plan[disp_plan["Status"].isin(filt_status)]
    if filt_vel:
        disp_plan = disp_plan[disp_plan["Velocity"].isin(filt_vel)]
    if show_only_dispatch:
        disp_plan = disp_plan[disp_plan["Dispatch"]>0]

    st.markdown(f"**Showing {len(disp_plan)} of {n_skus} SKUs**")
    styled_table(disp_plan, height=550)

    # Per-SKU stock breakdown
    st.divider()
    st.markdown("**🔍 SKU Detail — click to expand**")
    sel = st.selectbox("Select MSKU", sorted(plan["MSKU"].unique()), key="sku_sel")
    if sel:
        row = plan[plan["MSKU"]==sel].iloc[0]
        pn  = trunc(row.get("Product Name",""), 80)
        dc1,dc2,dc3,dc4,dc5 = st.columns(5)
        dc1.metric("Sellable Stock", n2(row["Sellable Stock"]))
        dc2.metric("Avg/Day",        f"{row['Avg/Day']:.3f}")
        dc3.metric("DOC",            f"{row['DOC']:.0f}d" if row['DOC']<9999 else "∞")
        dc4.metric("Dispatch Needed",n2(row["Dispatch"]))
        dc5.metric("Status",         row["Status"])
        st.caption(f"**{pn}**  |  FNSKU: `{row['FNSKU']}`  |  ASIN: `{row['ASIN']}`")

        fc_detail = fc_long[fc_long["MSKU"]==sel].copy()
        if not fc_detail.empty:
            fc_detail = fc_detail.rename(columns={"FC":"FC Code","FC Stock":"Current Stock"})
            fc_detail["FC DOC"] = np.where(row["Avg/Day"]>0,
                (fc_detail["Current Stock"]/row["Avg/Day"]).round(1), 9999)
            fc_detail["FC Status"] = fc_detail["FC DOC"].apply(lambda d: health(d, plan_days))
            st.dataframe(
                fc_detail[["FC Code","FC Name","FC City","FC Cluster",
                           "Current Stock","FC DOC","FC Status"]],
                use_container_width=True, hide_index=True)
        else:
            st.info("No stock at any FC for this SKU.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — NEED TO DISPATCH
# Exactly what you need to send — with FC-wise breakdown
# ══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("📦 SKUs That Need Dispatch")

    need = fba[fba["Dispatch"]>= max(1, min_units)].sort_values("Priority", ascending=False).copy()
    need["Product Name"] = need["Product Name"].apply(lambda x: trunc(x, 55))

    if need.empty:
        st.success("✅ All FBA SKUs are adequately stocked! No dispatch needed.")
    else:
        st.markdown(f"**{len(need)} SKUs need restocking — {n2(int(need['Dispatch'].sum()))} total units**")

        show_cols = ["MSKU","FNSKU","Product Name","Sellable Stock","Avg/Day",
                     "DOC","Status","Need","Dispatch","Priority"]
        styled_table(need[[c for c in show_cols if c in need.columns]], height=450)

        st.divider()
        st.subheader("🏭 FC-Wise Dispatch Breakdown")
        st.caption("Shows how many units to send to each FC for each SKU (proportional to current stock share)")

        fc_need = fcp[(fcp["Channel"]=="FBA") & (fcp["FC Dispatch"]>0)].copy()
        fc_need["Product Name"] = fc_need["Product Name"].apply(lambda x: trunc(x, 45))
        fc_show = fc_need.sort_values(["FC","Priority"], ascending=[True,False])
        fc_show = fc_show.rename(columns={"FC":"FC Code"})
        styled_table(
            fc_show[["FC Code","FC Name","FC City","MSKU","FNSKU","Product Name",
                     "FC Stock","Avg/Day","FC DOC","FC Dispatch"]],
            height=450)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — CRITICAL & DEAD STOCK
# ══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("🚨 Critical & Problem Stock")
    t1,t2,t3,t4 = st.tabs(["🔴 Critical (<14 days)","⚫ Dead Stock",
                             "🟡 Slow Moving (>90 days)","🔵 Excess Stock"])

    def risk_table(df):
        d = df.copy()
        d["Product Name"] = d["Product Name"].apply(lambda x: trunc(x,55))
        return d[["MSKU","FNSKU","Product Name","Sellable Stock","Avg/Day",
                  "DOC","Status","Dispatch","All Time Sales"]].reset_index(drop=True)

    with t1:
        st.markdown(f"**{len(r_crit)} SKU(s) — less than 14 days of stock left**")
        if r_crit.empty:
            st.success("✅ No critical SKUs!")
        else:
            styled_table(risk_table(r_crit.sort_values("DOC")))

    with t2:
        st.markdown(f"**{len(r_dead)} SKU(s) — stock sitting with ZERO sales**")
        st.caption("These products need listing review, pricing change, or removal/liquidation")
        if r_dead.empty:
            st.success("✅ No dead stock found.")
        else:
            d2 = r_dead.copy()
            d2["Product Name"] = d2["Product Name"].apply(lambda x: trunc(x,55))
            d2_show = d2[["MSKU","FNSKU","ASIN","Product Name","Sellable Stock",
                          "Damaged Stock","All Time Sales"]].reset_index(drop=True)
            styled_table(d2_show)
            # Show FC location of dead stock
            dead_fc = fc_long[fc_long["MSKU"].isin(r_dead["MSKU"])].copy()
            dead_fc = dead_fc.rename(columns={"FC":"FC Code"})
            if not dead_fc.empty:
                st.markdown("**📍 Where is dead stock sitting?**")
                dead_fc["Product Name"] = dead_fc["MSKU"].map(sku_master["Title"]).apply(lambda x: trunc(str(x),45))
                styled_table(dead_fc[["FC Code","FC Name","FC City","MSKU",
                                      "Product Name","FC Stock"]].sort_values("FC Stock", ascending=False))

    with t3:
        st.markdown(f"**{len(r_slow)} SKU(s) — more than 90 days of cover**")
        if r_slow.empty:
            st.success("✅ No slow-moving SKUs.")
        else:
            styled_table(risk_table(r_slow.sort_values("DOC", ascending=False)))

    with t4:
        st.markdown(f"**{len(r_excess)} SKU(s) — cover > {plan_days*2} days (overstocked)**")
        if r_excess.empty:
            st.success("✅ No excess stock.")
        else:
            styled_table(risk_table(r_excess.sort_values("DOC", ascending=False)))

    # Damaged stock section
    if not dmg.empty:
        st.divider()
        st.subheader("🔧 Damaged / Non-Sellable Stock")
        st.caption("Raise Removal Order or Reimbursement Claim in Seller Central")
        dmg_show = (dmg.groupby(["MSKU","Disp"])["Stock"].sum().reset_index()
            .rename(columns={"Disp":"Disposition","Stock":"Units"}))
        dmg_show["Product Name"] = dmg_show["MSKU"].map(sku_master["Title"]).apply(lambda x: trunc(str(x),55))
        styled_table(dmg_show[["MSKU","Product Name","Disposition","Units"]]
                     .sort_values("Units", ascending=False))


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — FC-WISE VIEW
# ══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("🏭 Fulfillment Center View")

    st.markdown("**📊 FC Inventory & Dispatch Summary**")
    styled_table(fc_disp.rename(columns={"FC":"FC Code"}), height=300)

    st.divider()
    sel_fc = st.selectbox("🔍 Drill into a specific FC",
        [f"{r['FC']} — {r['FC Name']}" for _,r in
         fc_disp.head(30).iterrows()], key="fc_sel")
    if sel_fc:
        fc_code = sel_fc.split(" — ")[0].strip()
        fc_skus = fcp[fcp["FC"]==fc_code].copy().sort_values("Priority", ascending=False)
        fc_skus["Product Name"] = fc_skus["Product Name"].apply(lambda x: trunc(x,45))
        st.markdown(f"**All SKUs at {fc_code} — {fc_name(fc_code)}**")
        st.dataframe(
            fc_skus.rename(columns={"FC":"FC Code"})[
                ["FC Code","MSKU","FNSKU","Product Name","FC Stock",
                 "Avg/Day","FC DOC","FC Status","FC Dispatch"]
            ].reset_index(drop=True),
            use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — SALES TRENDS
# ══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("📈 Sales Trends")
    sa, sb = st.columns(2)
    with sa:
        st.markdown("**Weekly Sales**")
        if not wk_trend.empty:
            st.line_chart(wk_trend.set_index("Week")["Units"])
    with sb:
        st.markdown("**Monthly Sales by Channel**")
        if not mo_trend.empty:
            pm = mo_trend.pivot_table(index="Month",columns="Channel",
                                       values="Units",aggfunc="sum",fill_value=0)
            st.bar_chart(pm)

    st.divider()
    st.markdown("**🔝 Top SKUs by Avg Daily Sale**")
    top = (plan[plan["Avg/Day"]>0].nlargest(20,"Avg/Day")
           [["MSKU","FNSKU","Product Name","Avg/Day","Sellable Stock",
             "DOC","Status","Dispatch"]].copy())
    top["Product Name"] = top["Product Name"].apply(lambda x: trunc(x,50))
    styled_table(top, height=400)

    st.divider()
    st.markdown("**📊 SKU Sales Chart**")
    pick = st.selectbox("Select SKU", sorted(plan["MSKU"].unique()), key="trend_pick")
    sd = sales[sales["MSKU"]==pick].groupby("Date")["Qty"].sum().reset_index().set_index("Date")
    if not sd.empty:
        pn = trunc(sku_master["Title"].get(pick,""), 70)
        st.caption(f"**{pick}** — {pn}")
        st.line_chart(sd["Qty"])
    else:
        st.info("No daily sales data for this SKU.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — BULK FBA SHIPMENT FILE
# Amazon Bulk Upload format — one file, all FCs, auto-creates shipments
# ══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("📄 Bulk FBA Shipment Upload File")
    st.info(
        "**How to use:**\n"
        "1. Review the table below — edit quantities if needed\n"
        "2. Enter shipment details\n"
        "3. Click **Generate & Download** — one file covers ALL FCs\n"
        "4. In Seller Central: **Send to Amazon → Create New Shipment → Upload a file**\n"
        "5. Amazon will automatically create separate shipments for each FC ✅"
    )

    bf1, bf2, bf3 = st.columns(3)
    with bf1:
        shp_name = st.text_input("Shipment Name / Reference",
            value=f"Dispatch_{datetime.now().strftime('%d%b%Y')}", key="shp_name")
        shp_date = st.date_input("Expected Ship Date",
            value=datetime.now().date()+timedelta(days=2), key="shp_date")
    with bf2:
        shp_by_fc = st.checkbox("Split by FC (separate file per FC)", value=False)
        sel_fcs   = st.multiselect("Include only these FCs (blank = all)",
                       active_fcs, default=[], key="sel_fcs")
    with bf3:
        st.markdown("**Shipment settings from sidebar:**")
        st.markdown(f"- Prep Owner: **{prep_own}**")
        st.markdown(f"- Label Owner: **{lbl_own}**")
        st.markdown(f"- Case Packed: **{'Yes' if case_packed else 'No'}**")
        st.markdown(f"- Units/Carton: **{case_qty}**")

    # Build the dispatch table
    dispatch_fcp = fcp[(fcp["Channel"]=="FBA") & (fcp["FC Dispatch"]>0)].copy()
    if sel_fcs:
        dispatch_fcp = dispatch_fcp[dispatch_fcp["FC"].isin(sel_fcs)]

    if dispatch_fcp.empty:
        st.warning("No dispatch needed right now (or no FCs selected).")
    else:
        dispatch_fcp["Product Name"] = dispatch_fcp["Product Name"].apply(lambda x: trunc(x,60))
        dispatch_fcp["Units to Ship"] = dispatch_fcp["FC Dispatch"].astype(int)
        dispatch_fcp["Cases"]         = np.ceil(dispatch_fcp["Units to Ship"]/case_qty).astype(int)
        dispatch_fcp["FC Name Short"] = dispatch_fcp["FC"].apply(fc_name)

        edit_cols = ["FC","FC Name Short","MSKU","FNSKU","Product Name",
                     "Units to Ship","Cases"]
        st.markdown(f"**{len(dispatch_fcp)} lines | {n2(dispatch_fcp['Units to Ship'].sum())} total units | "
                    f"{dispatch_fcp['FC'].nunique()} FCs**")

        edited = st.data_editor(
            dispatch_fcp[edit_cols].rename(columns={"FC":"FC Code","FC Name Short":"FC Name"})
                       .reset_index(drop=True),
            num_rows="dynamic", use_container_width=True, key="edit_dispatch")

        st.divider()

        # ── GENERATE FLAT FILE(S) ─────────────────────────────────────────
        def make_flat(rows_df, fc_code):
            """Generate Amazon FlatFile format for one FC."""
            lines = []
            lines.append("TemplateType=FlatFileShipmentCreation\tVersion=2015.0403")
            lines.append("")
            # Shipment header
            hdr = ["ShipmentName","ShipFromName","ShipFromAddressLine1","ShipFromCity",
                   "ShipFromStateOrProvinceCode","ShipFromPostalCode","ShipFromCountryCode",
                   "ShipmentStatus","LabelPrepType","AreCasesRequired",
                   "DestinationFulfillmentCenterId"]
            val = [shp_name, wh_name or "My Warehouse",
                   wh_addr or "Warehouse", wh_city or "City",
                   (wh_state or "UP").upper()[:2], wh_pin or "000000",
                   "IN","WORKING", lbl_own,
                   "YES" if case_packed else "NO", fc_code]
            lines.append("\t".join(hdr))
            lines.append("\t".join(str(v) for v in val))
            lines.append("")
            # Items header
            lines.append("\t".join(["SellerSKU","FNSKU","QuantityShipped",
                                     "QuantityInCase","PrepOwner","LabelOwner",
                                     "ItemDescription","ExpectedDeliveryDate"]))
            for _, r in rows_df.iterrows():
                qty = int(r.get("Units to Ship",0))
                if qty <= 0: continue
                qic = str(case_qty) if case_packed else ""
                desc = str(r.get("Product Name",""))[:200]
                lines.append("\t".join([
                    str(r.get("MSKU",r.get("SellerSKU",""))),
                    str(r.get("FNSKU","")),
                    str(qty), qic, prep_own, lbl_own,
                    desc, str(shp_date)]))
            return "\n".join(lines)

        if not shp_by_fc:
            # Single combined file — Amazon auto-splits by FC
            # Build per-FC shipment blocks combined in one file
            all_blocks = []
            fcs_in = edited["FC Code"].unique() if "FC Code" in edited.columns else []
            if len(fcs_in) == 0:
                fcs_in = edited.get("FC", edited.get("FC Code", pd.Series(active_fcs))).unique()

            for fc in sorted(fcs_in):
                fc_col = "FC Code" if "FC Code" in edited.columns else "FC"
                rows   = edited[edited[fc_col]==fc]
                if rows.empty: continue
                # Rename back for make_flat
                rows = rows.copy()
                if "SellerSKU" not in rows.columns:
                    rows = rows.rename(columns={"MSKU":"MSKU"})
                all_blocks.append(make_flat(rows, fc))

            combined = "\n\n".join(all_blocks)
            st.download_button(
                f"📥 Download Bulk Shipment File — All {len(fcs_in)} FCs Combined",
                combined.encode("utf-8"),
                f"FBA_Bulk_Shipment_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                "text/plain", use_container_width=True)

            with st.expander("👁️ Preview file"):
                st.code(combined[:3000]+("\n...(truncated)" if len(combined)>3000 else ""))
        else:
            # One file per FC
            st.markdown("**📦 Download one file per FC:**")
            fc_col = "FC Code" if "FC Code" in edited.columns else "FC"
            fcs_in = edited[fc_col].unique()
            cols_dl = st.columns(min(len(fcs_in), 4))
            for i, fc in enumerate(sorted(fcs_in)):
                rows = edited[edited[fc_col]==fc]
                if rows.empty: continue
                txt  = make_flat(rows, fc)
                with cols_dl[i % 4]:
                    st.download_button(
                        f"📥 {fc}\n{fc_name(fc)}",
                        txt.encode("utf-8"),
                        f"FBA_{fc}_{datetime.now().strftime('%Y%m%d')}.txt",
                        "text/plain", key=f"fc_{fc}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — DOWNLOAD REPORTS
# ══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("📥 Download Full Reports")

    def build_excel_report():
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:

            # ── Sheet 1: Full Plan ────────────────────────────────────────
            full = plan.copy()
            full["Product Name"] = full["Product Name"].apply(lambda x: trunc(x,80))
            full = full.sort_values("Priority", ascending=False)
            excel_write(w, full[BASE_COLS], "Full Plan")

            # ── Sheet 2: Dispatch Needed ──────────────────────────────────
            disp_out = fba[fba["Dispatch"]>0].copy()
            disp_out["Product Name"] = disp_out["Product Name"].apply(lambda x: trunc(x,80))
            excel_write(w, disp_out[BASE_COLS].sort_values("Priority",ascending=False),
                        "Dispatch Needed")

            # ── Sheet 3: FC-Wise Dispatch ─────────────────────────────────
            fc_out = fcp[(fcp["Channel"]=="FBA")].copy()
            fc_out["Product Name"] = fc_out["Product Name"].apply(lambda x: trunc(x,80))
            fc_out = fc_out.rename(columns={"FC":"FC Code"})
            excel_write(w, fc_out[["FC Code","FC Name","FC City","FC State","FC Cluster",
                                    "MSKU","FNSKU","Product Name","FC Stock","Avg/Day",
                                    "FC DOC","FC Status","FC Dispatch"]]
                         .sort_values(["FC Code","Priority"],ascending=[True,False]),
                        "FC-Wise Dispatch")

            # ── Sheet 4: FC Summary ───────────────────────────────────────
            excel_write(w, fc_disp.rename(columns={"FC":"FC Code"}), "FC Summary")

            # ── Sheet 5: Critical ─────────────────────────────────────────
            crit = r_crit.copy()
            crit["Product Name"] = crit["Product Name"].apply(lambda x: trunc(x,80))
            excel_write(w, crit[BASE_COLS].sort_values("DOC"), "Critical SKUs")

            # ── Sheet 6: Dead Stock ───────────────────────────────────────
            dead = r_dead.copy()
            dead["Product Name"] = dead["Product Name"].apply(lambda x: trunc(x,80))
            dcols = ["MSKU","FNSKU","ASIN","Product Name","Sellable Stock",
                     "Damaged Stock","All Time Sales"]
            excel_write(w, dead[[c for c in dcols if c in dead.columns]], "Dead Stock")

            # ── Sheet 7: Dead Stock by FC ─────────────────────────────────
            dead_fc2 = fc_long[fc_long["MSKU"].isin(r_dead["MSKU"])].copy()
            dead_fc2["Product Name"] = dead_fc2["MSKU"].map(sku_master["Title"]).apply(lambda x: trunc(str(x),80))
            dead_fc2 = dead_fc2.rename(columns={"FC":"FC Code"})
            excel_write(w, dead_fc2[["FC Code","FC Name","FC City","MSKU","Product Name","FC Stock"]]
                         .sort_values("FC Stock",ascending=False),
                        "Dead Stock by FC")

            # ── Sheet 8: Slow Moving ──────────────────────────────────────
            slow = r_slow.copy()
            slow["Product Name"] = slow["Product Name"].apply(lambda x: trunc(x,80))
            excel_write(w, slow[BASE_COLS].sort_values("DOC",ascending=False), "Slow Moving")

            # ── Sheet 9: Excess ───────────────────────────────────────────
            exc = r_excess.copy()
            exc["Product Name"] = exc["Product Name"].apply(lambda x: trunc(x,80))
            excel_write(w, exc[BASE_COLS].sort_values("DOC",ascending=False), "Excess Stock")

            # ── Sheet 10: Damaged ─────────────────────────────────────────
            if not dmg.empty:
                dmg2 = (dmg.groupby(["MSKU","Disp"])["Stock"].sum().reset_index()
                    .rename(columns={"Disp":"Disposition","Stock":"Units"}))
                dmg2["Product Name"] = dmg2["MSKU"].map(sku_master["Title"]).apply(lambda x: trunc(str(x),80))
                excel_write(w, dmg2[["MSKU","Product Name","Disposition","Units"]]
                             .sort_values("Units",ascending=False), "Damaged Stock")

            # ── Sheet 11: Monthly Trend ───────────────────────────────────
            excel_write(w, mo_trend, "Monthly Trend")

            # ── Sheet 12: Weekly Trend ────────────────────────────────────
            excel_write(w, wk_trend, "Weekly Trend")

            # ── Sheet 13: SKU Master ──────────────────────────────────────
            sm = sku_master.reset_index()
            sm["Title"] = sm["Title"].apply(lambda x: trunc(x,80))
            excel_write(w, sm[["MSKU","FNSKU","ASIN","Title"]], "SKU Master")

            # ── Sheet 14: FC Master ───────────────────────────────────────
            fc_master = pd.DataFrame([
                {"FC Code":k,"FC Name":v[0],"City":v[1],"State":v[2],"Cluster":v[3]}
                for k,v in FC.items()])
            excel_write(w, fc_master, "FC Master")

        buf.seek(0)
        return buf

    def build_dispatch_csv():
        out = fba[fba["Dispatch"]>0].copy()
        out["Product Name"] = out["Product Name"].apply(lambda x: trunc(x,80))
        cols = [c for c in BASE_COLS if c in out.columns]
        return out[cols].sort_values("Priority", ascending=False).to_csv(index=False)

    def build_fc_dispatch_csv():
        out = fcp[(fcp["Channel"]=="FBA") & (fcp["FC Dispatch"]>0)].copy()
        out["Product Name"] = out["Product Name"].apply(lambda x: trunc(x,80))
        out = out.rename(columns={"FC":"FC Code"})
        return out[["FC Code","FC Name","FC City","MSKU","FNSKU","Product Name",
                    "FC Stock","Avg/Day","FC Dispatch"]].to_csv(index=False)

    r1,r2,r3 = st.columns(3)
    with r1:
        st.download_button(
            "📗 Full Excel Report (14 sheets)",
            data=build_excel_report(),
            file_name=f"FBA_Supply_Plan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with r2:
        st.download_button(
            "📋 Dispatch Plan CSV",
            data=build_dispatch_csv(),
            file_name=f"Dispatch_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True)
    with r3:
        st.download_button(
            "🏭 FC-Wise Dispatch CSV",
            data=build_fc_dispatch_csv(),
            file_name=f"FC_Dispatch_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True)

    st.divider()
    st.markdown("**What's in each download:**")
    st.markdown("""
| Download | Contents |
|---|---|
| **Full Excel Report** | 14 sheets: Full Plan, Dispatch Needed, FC-Wise Dispatch, FC Summary, Critical SKUs, Dead Stock, Dead Stock by FC, Slow Moving, Excess Stock, Damaged, Monthly Trend, Weekly Trend, SKU Master, FC Master |
| **Dispatch Plan CSV** | Only SKUs that need dispatch — sorted by priority |
| **FC-Wise Dispatch CSV** | One row per SKU per FC — for warehouse picking |
""")
