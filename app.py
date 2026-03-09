import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import math
from datetime import datetime, timedelta

# ================== PAGE CONFIG ==================
st.set_page_config(
    page_title="FBA Smart Supply Planner",
    layout="wide",
    page_icon="📦",
)

# ================== CUSTOM CSS ==================
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 0.25rem;
    }
    .metric-card-green {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    }
    .metric-card-red {
        background: linear-gradient(135deg, #eb3349 0%, #f45c43 100%);
    }
    .metric-card-orange {
        background: linear-gradient(135deg, #f7971e 0%, #ffd200 100%);
    }
    .alert-box {
        padding: 0.75rem 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        font-weight: 500;
    }
    .alert-red { background: #fde8e8; border-left: 4px solid #e53e3e; color: #742a2a; }
    .alert-yellow { background: #fefcbf; border-left: 4px solid #d69e2e; color: #744210; }
    .alert-green { background: #f0fff4; border-left: 4px solid #38a169; color: #1c4532; }
    .stTabs [data-baseweb="tab"] { font-size: 14px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.title("📦 FBA Smart Supply Planner — Full Automation Suite")
st.caption("Upload your MTR (sales) + Inventory Ledger → Get dispatch plan, FC allocation, Amazon flat file & insights")

# ================== FC / CLUSTER MAPPING ==================
# Correct Amazon India FC cluster mapping
FC_DATA = {
    # Delhi NCR Cluster
    "DEX3": {"name": "New Delhi FC (DEX3)", "city": "New Delhi", "state": "DELHI", "cluster": "Delhi NCR"},
    "DEX8": {"name": "New Delhi FC (DEX8)", "city": "New Delhi", "state": "DELHI", "cluster": "Delhi NCR"},
    "DEL4": {"name": "Delhi North FC (DEL4)", "city": "Delhi", "state": "DELHI", "cluster": "Delhi NCR"},
    "DEL5": {"name": "Delhi North FC (DEL5)", "city": "Delhi", "state": "DELHI", "cluster": "Delhi NCR"},
    "DEL6": {"name": "Manesar FC (DEL6)", "city": "Manesar", "state": "HARYANA", "cluster": "Delhi NCR"},
    "DEL7": {"name": "Bilaspur FC (DEL7)", "city": "Bilaspur", "state": "HARYANA", "cluster": "Delhi NCR"},
    "XDEL": {"name": "Delhi XL FC (XDEL)", "city": "New Delhi", "state": "DELHI", "cluster": "Delhi NCR"},

    # Punjab / North
    "LDH1": {"name": "Ludhiana FC (LDH1)", "city": "Ludhiana", "state": "PUNJAB", "cluster": "North"},
    "JAI1": {"name": "Jaipur FC (JAI1)", "city": "Jaipur", "state": "RAJASTHAN", "cluster": "North"},
    "LKO1": {"name": "Lucknow FC (LKO1)", "city": "Lucknow", "state": "UTTAR PRADESH", "cluster": "North"},
    "AGR1": {"name": "Agra FC (AGR1)", "city": "Agra", "state": "UTTAR PRADESH", "cluster": "North"},

    # Mumbai / West Cluster
    "BOM1": {"name": "Bhiwandi FC (BOM1)", "city": "Bhiwandi", "state": "MAHARASHTRA", "cluster": "Mumbai West"},
    "BOM3": {"name": "Nashik FC (BOM3)", "city": "Nashik", "state": "MAHARASHTRA", "cluster": "Mumbai West"},
    "BOM4": {"name": "Vasai FC (BOM4)", "city": "Vasai", "state": "MAHARASHTRA", "cluster": "Mumbai West"},
    "SAMB": {"name": "Mumbai West FC (SAMB)", "city": "Mumbai", "state": "MAHARASHTRA", "cluster": "Mumbai West"},
    "PNQ1": {"name": "Pune FC (PNQ1)", "city": "Pune", "state": "MAHARASHTRA", "cluster": "Mumbai West"},
    "PNQ2": {"name": "Pune FC (PNQ2)", "city": "Pune", "state": "MAHARASHTRA", "cluster": "Mumbai West"},
    "XBOM": {"name": "Mumbai XL FC (XBOM)", "city": "Bhiwandi", "state": "MAHARASHTRA", "cluster": "Mumbai West"},

    # Bangalore / South Cluster
    "BLR5": {"name": "Bangalore South FC (BLR5)", "city": "Bangalore", "state": "KARNATAKA", "cluster": "Bangalore South"},
    "BLR6": {"name": "Bangalore FC (BLR6)", "city": "Bangalore", "state": "KARNATAKA", "cluster": "Bangalore South"},
    "SCJA": {"name": "Bangalore FC (SCJA)", "city": "Bangalore", "state": "KARNATAKA", "cluster": "Bangalore South"},
    "XSAB": {"name": "Bangalore XS FC (XSAB)", "city": "Bangalore", "state": "KARNATAKA", "cluster": "Bangalore South"},
    "SBLA": {"name": "Bangalore SBLA FC (SBLA)", "city": "Bangalore", "state": "KARNATAKA", "cluster": "Bangalore South"},

    # Chennai / South East
    "MAA4": {"name": "Chennai South FC (MAA4)", "city": "Chennai", "state": "TAMIL NADU", "cluster": "Chennai"},
    "MAA5": {"name": "Chennai FC (MAA5)", "city": "Chennai", "state": "TAMIL NADU", "cluster": "Chennai"},
    "SMAB": {"name": "Chennai SMAB FC (SMAB)", "city": "Chennai", "state": "TAMIL NADU", "cluster": "Chennai"},

    # Hyderabad Cluster
    "HYD7": {"name": "Hyderabad FC (HYD7)", "city": "Hyderabad", "state": "TELANGANA", "cluster": "Hyderabad"},
    "HYD8": {"name": "Hyderabad FC (HYD8)", "city": "Hyderabad", "state": "TELANGANA", "cluster": "Hyderabad"},
    "HYD9": {"name": "Hyderabad FC (HYD9)", "city": "Hyderabad", "state": "TELANGANA", "cluster": "Hyderabad"},

    # Kolkata / East Cluster
    "CCU1": {"name": "Kolkata FC (CCU1)", "city": "Kolkata", "state": "WEST BENGAL", "cluster": "Kolkata East"},
    "CCU2": {"name": "Kolkata FC (CCU2)", "city": "Kolkata", "state": "WEST BENGAL", "cluster": "Kolkata East"},
    "PAT1": {"name": "Patna FC (PAT1)", "city": "Patna", "state": "BIHAR", "cluster": "Kolkata East"},

    # Gujarat / West
    "AMD1": {"name": "Ahmedabad FC (AMD1)", "city": "Ahmedabad", "state": "GUJARAT", "cluster": "Gujarat West"},
    "AMD2": {"name": "Ahmedabad FC (AMD2)", "city": "Ahmedabad", "state": "GUJARAT", "cluster": "Gujarat West"},
    "SUB1": {"name": "Surat FC (SUB1)", "city": "Surat", "state": "GUJARAT", "cluster": "Gujarat West"},
}


def fc_info(code: str) -> dict:
    code = str(code).upper().strip()
    if code in FC_DATA:
        return FC_DATA[code]
    # Try prefix match for unknown codes
    for k, v in FC_DATA.items():
        if code.startswith(k[:3]):
            return v
    return {"name": f"FC {code}", "city": "Unknown", "state": "Unknown", "cluster": "Other"}


def fc_display_name(code: str) -> str:
    return fc_info(code)["name"]


def fc_cluster(code: str) -> str:
    return fc_info(code)["cluster"]


def fc_state(code: str) -> str:
    return fc_info(code)["state"]


# ================== HELPERS ==================
def add_serial(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if df.empty:
        return df
    if "S.No" in df.columns:
        df = df.drop(columns=["S.No"])
    df.insert(0, "S.No", range(1, len(df) + 1))
    return df


def read_file(file) -> pd.DataFrame:
    try:
        name = file.name.lower()
        if name.endswith(".zip"):
            with zipfile.ZipFile(file) as z:
                for member in z.namelist():
                    if member.lower().endswith(".csv"):
                        return pd.read_csv(z.open(member), low_memory=False)
            return pd.DataFrame()
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            return pd.read_excel(file)
        else:
            return pd.read_csv(file, low_memory=False)
    except Exception as e:
        st.warning(f"File read error: {e}")
        return pd.DataFrame()


def detect_col(df: pd.DataFrame, candidates: list, substrings: list = None) -> str:
    normalized = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in normalized:
            return normalized[key]
    if substrings:
        for col in df.columns:
            low = col.lower()
            if any(sub in low for sub in substrings):
                return col
    return ""


def fmt_num(n):
    try:
        return f"{int(n):,}"
    except Exception:
        return str(n)


def days_of_cover_tag(days):
    if days < 14:
        return "🔴 Critical"
    elif days < 30:
        return "🟠 Low"
    elif days < 90:
        return "🟢 Healthy"
    elif days < 180:
        return "🟡 Excess"
    else:
        return "⚫ Dead Stock"


# ================== SIDEBAR CONTROLS ==================
with st.sidebar:
    st.header("⚙️ Planning Controls")
    planning_days = st.number_input("Planning Days (Coverage Target)", 7, 180, 60, 1)

    service_level = st.selectbox("Service Level (Safety Stock)", ["90%", "95%", "98%"])
    Z_MAP = {"90%": 1.28, "95%": 1.65, "98%": 2.05}
    z_value = float(Z_MAP[service_level])

    planning_basis = st.selectbox(
        "Sales Basis",
        ["Total Sales (full history)", "Last 30 Days", "Last 60 Days", "Last 90 Days"],
    )

    st.divider()
    st.subheader("🚚 Amazon Shipment Settings")
    seller_id = st.text_input("Seller ID (optional)", placeholder="A1B2C3D4E5F6G7")
    ship_from_name = st.text_input("Ship-From Name", placeholder="My Warehouse")
    ship_from_address = st.text_area("Ship-From Address", placeholder="123 Warehouse St, City, State - 400001")
    default_case_qty = st.number_input("Default Case Qty per Carton", 1, 500, 12)

    st.divider()
    st.subheader("🎯 Filters")
    min_dispatch_filter = st.number_input("Min Dispatch Units to Show", 0, 10000, 0)
    focus_cluster = st.multiselect(
        "Focus Clusters",
        list(set(v["cluster"] for v in FC_DATA.values())),
    )

# ================== FILE UPLOADS ==================
st.markdown("### 📁 Upload Files")
col_u1, col_u2 = st.columns(2)
with col_u1:
    mtr_files = st.file_uploader(
        "📊 Sales / MTR Report (CSV, ZIP, XLSX) — Multiple OK",
        type=["csv", "zip", "xlsx"],
        accept_multiple_files=True,
    )
with col_u2:
    inventory_file = st.file_uploader(
        "🏭 Inventory Ledger (CSV, ZIP, XLSX)",
        type=["csv", "zip", "xlsx"],
    )

# ================== MAIN LOGIC ==================
if mtr_files and inventory_file:

    # -------- LOAD SALES --------
    sales_list = [read_file(f) for f in mtr_files if not read_file(f).empty]
    if not sales_list:
        st.error("Could not read any MTR files.")
        st.stop()
    sales = pd.concat(sales_list, ignore_index=True)

    # Core column detection
    sku_col = detect_col(sales, ["Sku", "SKU", "MSKU", "ASIN", "Item SKU"], ["sku", "asin", "msku"])
    qty_col = detect_col(sales, ["Quantity", "Qty", "Units Ordered", "quantity"], ["qty", "unit", "quant"])
    date_col = detect_col(sales, ["Shipment Date", "Purchase Date", "Order Date", "Date"], ["date", "shipment"])
    state_col = detect_col(sales, ["Ship To State", "Ship-to State", "Shipping State"], ["ship to", "state"])

    missing = []
    if not sku_col: missing.append("SKU column")
    if not qty_col: missing.append("Quantity column")
    if not date_col: missing.append("Date column")
    if not state_col: missing.append("Ship To State column")

    if missing:
        st.error(f"Missing columns in Sales file: {', '.join(missing)}")
        with st.expander("Detected columns"):
            st.write(list(sales.columns))
        st.stop()

    # Rename to standard
    sales = sales.rename(columns={
        sku_col: "Sku", qty_col: "Quantity",
        date_col: "Shipment Date", state_col: "Ship To State"
    })

    sales["Sku"] = sales["Sku"].astype(str).str.strip()
    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
    sales["Ship To State"] = sales["Ship To State"].astype(str).str.upper().str.strip()

    # Fulfillment Type
    ft_col = detect_col(sales,
        ["Fulfilment", "Fulfillment", "Fulfilment Channel", "Fulfillment Channel", "Channel"],
        ["fulfil", "channel"])
    if ft_col:
        ft_map = {"AFN": "FBA", "FBA": "FBA", "AMAZON_FULFILLED": "FBA",
                  "MFN": "FBM", "FBM": "FBM", "MERCHANT_FULFILLED": "FBM"}
        sales["Fulfillment Type"] = (sales[ft_col].astype(str).str.upper().str.strip()
                                     .map(ft_map).fillna("FBA"))
    else:
        sales["Fulfillment Type"] = "FBA"

    # Ship From State
    sf_col = detect_col(sales,
        ["Ship From State", "Ship from State", "Dispatch State"],
        ["ship from", "dispatch", "origin"])
    sales["Ship From State"] = (sales[sf_col].astype(str).str.upper().str.strip()
                                 if sf_col else "UNKNOWN")

    sales = sales.dropna(subset=["Shipment Date"])
    sales = sales[sales["Quantity"] > 0]  # Only actual sales

    # Filter: only B2C orders (exclude B2B if column exists)
    b2b_col = detect_col(sales, ["Is Business Order", "Business Order", "B2B"], ["b2b", "business"])
    if b2b_col:
        sales = sales[sales[b2b_col].astype(str).str.upper().str.strip().isin(["NO", "FALSE", "0", "N", ""])]

    # -------- SALES WINDOW --------
    overall_min = sales["Shipment Date"].min()
    overall_max = sales["Shipment Date"].max()
    uploaded_days = max((overall_max - overall_min).days + 1, 1)

    if "30" in planning_basis:
        window = 30
    elif "60" in planning_basis:
        window = 60
    elif "90" in planning_basis:
        window = 90
    else:
        window = uploaded_days

    cutoff = overall_max - pd.Timedelta(days=window - 1)
    hist = sales[sales["Shipment Date"] >= cutoff].copy() if window < uploaded_days else sales.copy()
    history_window_days = max((hist["Shipment Date"].max() - hist["Shipment Date"].min()).days + 1, 1)

    # -------- LOAD INVENTORY --------
    inv = read_file(inventory_file)
    if inv.empty:
        st.error("Inventory file unreadable.")
        st.stop()
    inv.columns = inv.columns.str.strip()

    inv_sku_col = detect_col(inv, ["ASIN", "MSKU", "Sku", "SKU", "Item SKU"], ["sku", "asin", "msku"])
    inv_qty_col = detect_col(inv, ["Ending Warehouse Balance", "Quantity", "Available", "Qty Available"],
                             ["ending", "available", "balance", "qty"])

    if not inv_sku_col:
        st.error("Inventory must have a SKU/ASIN/MSKU column.")
        st.stop()
    if not inv_qty_col:
        st.error("Inventory must have a quantity/balance column.")
        st.stop()

    inv = inv.rename(columns={inv_sku_col: "Sku", inv_qty_col: "Ending Warehouse Balance"})
    inv["Sku"] = inv["Sku"].astype(str).str.strip()
    inv["Ending Warehouse Balance"] = pd.to_numeric(inv["Ending Warehouse Balance"], errors="coerce").fillna(0)

    # Fulfillment Type in inventory
    inv_ft_col = detect_col(inv,
        ["Fulfilment", "Fulfillment", "Fulfilment Channel", "Fulfillment Channel", "Inventory Type"],
        ["fulfil", "channel", "inventory type"])
    if inv_ft_col:
        ft_map2 = {"AFN": "FBA", "FBA": "FBA", "AMAZON_FULFILLED": "FBA",
                   "MFN": "FBM", "FBM": "FBM", "MERCHANT_FULFILLED": "FBM"}
        inv["Fulfillment Type"] = (inv[inv_ft_col].astype(str).str.upper().str.strip()
                                   .map(ft_map2).fillna("FBA"))
    else:
        inv["Fulfillment Type"] = "FBA"

    # Ship From State
    inv_sf_col = detect_col(inv,
        ["Ship From State", "Ship from State", "Warehouse State", "FC State", "State"],
        ["state", "ship from", "warehouse"])
    inv["Ship From State"] = (inv[inv_sf_col].astype(str).str.upper().str.strip()
                               if inv_sf_col else "UNKNOWN")

    # Warehouse Code
    inv_fc_col = detect_col(inv,
        ["Warehouse Code", "Warehouse", "FC Code", "Fulfillment Center", "FC"],
        ["warehouse", " fc", "fulfillment center"])
    inv["Warehouse Code"] = (inv[inv_fc_col].astype(str).str.upper().str.strip()
                              if inv_fc_col else "UNKNOWN")

    # Enrich FC info
    inv["FC Name"] = inv["Warehouse Code"].apply(fc_display_name)
    inv["FC Cluster"] = inv["Warehouse Code"].apply(fc_cluster)
    inv["FC State"] = inv["Warehouse Code"].apply(fc_state)

    # -------- PRODUCT NAME (if available) --------
    prod_col = detect_col(sales, ["Product Name", "Item Name", "Title", "Description"], ["product", "item name", "title"])
    if prod_col:
        prod_map = sales.groupby("Sku")[prod_col].first().to_dict()
    else:
        prod_map = {}

    # -------- STOCK AGGREGATES --------
    sku_stock_all = (inv.groupby("Sku")["Ending Warehouse Balance"].sum()
                     .reset_index().rename(columns={"Ending Warehouse Balance": "Current Stock (All)"}))

    stock_by_ft = (inv.groupby(["Sku", "Fulfillment Type"])["Ending Warehouse Balance"]
                   .sum().reset_index()
                   .rename(columns={"Ending Warehouse Balance": "Current Stock (Channel)"}))

    stock_by_site = (inv.groupby(["Sku", "Fulfillment Type", "Ship From State", "Warehouse Code"])
                     ["Ending Warehouse Balance"].sum().reset_index()
                     .rename(columns={"Ending Warehouse Balance": "Current Stock"}))
    stock_by_site["FC Name"] = stock_by_site["Warehouse Code"].apply(fc_display_name)
    stock_by_site["FC Cluster"] = stock_by_site["Warehouse Code"].apply(fc_cluster)
    stock_by_site["FC State"] = stock_by_site["Warehouse Code"].apply(fc_state)

    # -------- DEMAND CALCULATIONS --------
    hist_ft = (hist.groupby(["Sku", "Fulfillment Type"])["Quantity"]
               .sum().reset_index().rename(columns={"Quantity": "History Sales (Channel)"}))
    full_ft = (sales.groupby(["Sku", "Fulfillment Type"])["Quantity"]
               .sum().reset_index().rename(columns={"Quantity": "Total Sales (All Time)"}))

    daily_ft = (hist.groupby(["Sku", "Fulfillment Type", "Shipment Date"])["Quantity"]
                .sum().reset_index())
    std_dev_ft = (daily_ft.groupby(["Sku", "Fulfillment Type"])["Quantity"]
                  .std().reset_index().rename(columns={"Quantity": "Demand StdDev"}))

    # State-level demand (top states per SKU)
    sku_state_sales = (hist.groupby(["Sku", "Ship To State"])["Quantity"]
                       .sum().reset_index().rename(columns={"Quantity": "State Sales"}))
    top_state_per_sku = (sku_state_sales.sort_values("State Sales", ascending=False)
                         .groupby("Sku").head(3)
                         .groupby("Sku")["Ship To State"]
                         .apply(lambda x: ", ".join(x.tolist()))
                         .reset_index().rename(columns={"Ship To State": "Top Sales States"}))

    # -------- CHANNEL PLANNING TABLE --------
    base_keys = pd.concat([
        hist_ft[["Sku", "Fulfillment Type"]],
        stock_by_ft[["Sku", "Fulfillment Type"]],
    ], ignore_index=True).drop_duplicates()

    plan = base_keys.copy()
    plan = plan.merge(hist_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan = plan.merge(full_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan = plan.merge(stock_by_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan = plan.merge(std_dev_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan = plan.merge(top_state_per_sku, on="Sku", how="left")

    for col in ["History Sales (Channel)", "Total Sales (All Time)", "Current Stock (Channel)", "Demand StdDev"]:
        plan[col] = pd.to_numeric(plan[col], errors="coerce").fillna(0)

    plan["Sales Data Days (Uploaded)"] = uploaded_days
    plan["History Window Days"] = history_window_days
    plan["Planning Days"] = planning_days

    plan["Avg Daily Sale"] = plan["History Sales (Channel)"] / history_window_days
    plan["Safety Stock"] = (z_value * plan["Demand StdDev"] * math.sqrt(planning_days)).round(0)
    plan["Base Requirement"] = (plan["Avg Daily Sale"] * planning_days).round(0)
    plan["Required Stock"] = plan["Base Requirement"] + plan["Safety Stock"]
    plan["Recommended Dispatch"] = (plan["Required Stock"] - plan["Current Stock (Channel)"]).clip(lower=0).round(0)

    plan["Days of Cover"] = (plan["Current Stock (Channel)"] /
                              plan["Avg Daily Sale"].replace(0, np.nan)).fillna(0).round(1)

    plan["Stock Health"] = plan["Days of Cover"].apply(days_of_cover_tag)

    # Product name enrichment
    plan["Product Name"] = plan["Sku"].map(prod_map).fillna("")

    # -------- PRIORITY SCORE --------
    # Higher score = more urgent to replenish
    def priority_score(row):
        if row["Avg Daily Sale"] <= 0:
            return 0
        urgency = max(0, (planning_days - row["Days of Cover"]) / planning_days)
        velocity = min(row["Avg Daily Sale"] / (plan["Avg Daily Sale"].max() or 1), 1)
        return round((urgency * 0.7 + velocity * 0.3) * 100, 1)

    plan["Priority Score (0-100)"] = plan.apply(priority_score, axis=1)

    def classify_velocity(avg):
        if avg <= 0:
            return "⚫ Dead"
        elif avg < 0.5:
            return "🔵 Slow"
        elif avg < 2:
            return "🟡 Medium"
        elif avg < 5:
            return "🟢 Fast"
        else:
            return "🔥 Top Seller"

    plan["Velocity Tag"] = plan["Avg Daily Sale"].apply(classify_velocity)

    plan = add_serial(plan)
    fba_plan = plan[plan["Fulfillment Type"] == "FBA"].copy()
    fbm_plan = plan[plan["Fulfillment Type"] == "FBM"].copy()

    # -------- SITE-LEVEL PLAN --------
    site_hist = (hist.groupby(["Sku", "Fulfillment Type", "Ship From State"])["Quantity"]
                 .sum().reset_index().rename(columns={"Quantity": "History Sales (Site)"}))

    site_plan = stock_by_site.merge(site_hist, on=["Sku", "Fulfillment Type", "Ship From State"], how="left")
    site_plan["History Sales (Site)"] = site_plan["History Sales (Site)"].fillna(0)
    site_plan["Avg Daily Sale (Site)"] = site_plan["History Sales (Site)"] / history_window_days
    site_plan["Days of Cover (Site)"] = (site_plan["Current Stock"] /
                                          site_plan["Avg Daily Sale (Site)"].replace(0, np.nan)).fillna(0).round(1)

    # Merge channel-level dispatch
    site_plan = site_plan.merge(
        plan[["Sku", "Fulfillment Type", "Recommended Dispatch", "History Sales (Channel)"]],
        on=["Sku", "Fulfillment Type"], how="left"
    )

    # Demand share per site
    site_total = site_plan.groupby(["Sku", "Fulfillment Type"])["History Sales (Site)"].sum().reset_index()
    site_total.rename(columns={"History Sales (Site)": "Total Site Sales"}, inplace=True)
    site_plan = site_plan.merge(site_total, on=["Sku", "Fulfillment Type"], how="left")

    def demand_share(row):
        if row["Total Site Sales"] > 0:
            return row["History Sales (Site)"] / row["Total Site Sales"]
        n = site_plan[(site_plan["Sku"] == row["Sku"]) &
                      (site_plan["Fulfillment Type"] == row["Fulfillment Type"])].shape[0]
        return 1.0 / max(n, 1)

    site_plan["Demand Share"] = site_plan.apply(demand_share, axis=1)
    site_plan["Recommended Dispatch (FC)"] = (site_plan["Recommended Dispatch"] * site_plan["Demand Share"]).round(0)
    site_plan["FC Priority"] = site_plan["Days of Cover (Site)"].apply(
        lambda d: "🔴 Urgent" if d < 14 else ("🟠 Soon" if d < 30 else "🟢 OK"))
    site_plan["Product Name"] = site_plan["Sku"].map(prod_map).fillna("")
    site_plan = add_serial(site_plan)

    # Apply cluster filter
    if focus_cluster:
        site_plan_view = site_plan[site_plan["FC Cluster"].isin(focus_cluster)].copy()
    else:
        site_plan_view = site_plan.copy()

    # Apply min dispatch filter
    site_plan_view = site_plan_view[site_plan_view["Recommended Dispatch (FC)"] >= min_dispatch_filter]

    # -------- SKU SUMMARY --------
    sku_summary = (plan.groupby("Sku").agg(
        Channels=("Fulfillment Type", lambda x: ", ".join(sorted(set(x)))),
        Total_History_Sales=("History Sales (Channel)", "sum"),
        Total_Current_Stock=("Current Stock (Channel)", "sum"),
        Total_Required=("Required Stock", "sum"),
        Total_Dispatch=("Recommended Dispatch", "sum"),
        Avg_Daily_Sale=("Avg Daily Sale", "sum"),
        Best_Days_Cover=("Days of Cover", "max"),
    ).reset_index())
    sku_summary = sku_summary.merge(sku_stock_all, on="Sku", how="left")
    sku_summary["Product Name"] = sku_summary["Sku"].map(prod_map).fillna("")
    sku_summary["Top States"] = sku_summary["Sku"].map(
        top_state_per_sku.set_index("Sku")["Top Sales States"].to_dict()).fillna("")
    sku_summary["Velocity"] = sku_summary["Avg_Daily_Sale"].apply(classify_velocity)
    sku_summary = add_serial(sku_summary)

    # -------- CLUSTER SUMMARY --------
    cluster_summary = (site_plan.groupby("FC Cluster").agg(
        FCs=("Warehouse Code", lambda x: ", ".join(sorted(set(x)))),
        Total_Current_Stock=("Current Stock", "sum"),
        Total_Dispatch=("Recommended Dispatch (FC)", "sum"),
        Avg_DOC=("Days of Cover (Site)", "mean"),
        SKUs=("Sku", "nunique"),
    ).reset_index().rename(columns={
        "FCs": "FC Codes", "Total_Current_Stock": "Total Stock",
        "Total_Dispatch": "Dispatch Needed", "Avg_DOC": "Avg Days of Cover",
        "SKUs": "Unique SKUs"
    }))
    cluster_summary["Avg Days of Cover"] = cluster_summary["Avg Days of Cover"].round(1)
    cluster_summary = cluster_summary.sort_values("Dispatch Needed", ascending=False)
    cluster_summary = add_serial(cluster_summary)

    # -------- RISK TABLES --------
    dead_stock = plan[(plan["Avg Daily Sale"] == 0) & (plan["Current Stock (Channel)"] > 0)].copy()
    slow_moving = plan[(plan["Avg Daily Sale"] > 0) & (plan["Days of Cover"] > 90)].copy()
    excess_stock = plan[plan["Days of Cover"] > (planning_days * 2)].copy()
    critical_stock = plan[(plan["Days of Cover"] < 14) & (plan["Avg Daily Sale"] > 0)].copy()
    top_sellers = plan[plan["Avg Daily Sale"] > 0].nlargest(20, "Avg Daily Sale").copy()

    for df in [dead_stock, slow_moving, excess_stock, critical_stock, top_sellers]:
        df = add_serial(df)

    # -------- STATE ANALYSIS --------
    state_to = (sales.groupby("Ship To State")["Quantity"].sum()
                .reset_index().rename(columns={"Quantity": "Total Units Sold"})
                .sort_values("Total Units Sold", ascending=False))
    state_to["% Share"] = (state_to["Total Units Sold"] / state_to["Total Units Sold"].sum() * 100).round(1)
    state_to = add_serial(state_to)

    state_from = (sales.groupby("Ship From State")["Quantity"].sum()
                  .reset_index().rename(columns={"Quantity": "Total Units Shipped"})
                  .sort_values("Total Units Shipped", ascending=False))
    state_from = add_serial(state_from)

    # -------- TREND ANALYSIS --------
    weekly_sales = (sales.copy())
    weekly_sales["Week"] = weekly_sales["Shipment Date"].dt.to_period("W").astype(str)
    weekly_trend = (weekly_sales.groupby("Week")["Quantity"].sum()
                    .reset_index().rename(columns={"Quantity": "Units Sold"})
                    .sort_values("Week"))

    # ================== DASHBOARD METRICS ==================
    st.markdown("---")
    st.markdown("### 📊 Executive Dashboard")

    total_skus = plan["Sku"].nunique()
    total_stock = int(plan["Current Stock (Channel)"].sum())
    total_dispatch = int(plan["Recommended Dispatch"].sum())
    critical_count = len(critical_stock)
    dead_count = len(dead_stock)
    avg_doc = plan[plan["Avg Daily Sale"] > 0]["Days of Cover"].mean()

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Total SKUs", fmt_num(total_skus))
    m2.metric("Total Stock", fmt_num(total_stock))
    m3.metric("Units to Dispatch", fmt_num(total_dispatch))
    m4.metric("🔴 Critical SKUs", fmt_num(critical_count))
    m5.metric("⚫ Dead Stock SKUs", fmt_num(dead_count))
    m6.metric("Avg Days of Cover", f"{avg_doc:.0f}d" if not np.isnan(avg_doc) else "N/A")

    # Alerts
    if critical_count > 0:
        st.markdown(f'<div class="alert-box alert-red">🚨 {critical_count} SKUs have less than 14 days of stock — urgent replenishment needed!</div>', unsafe_allow_html=True)
    if dead_count > 0:
        st.markdown(f'<div class="alert-box alert-yellow">⚠️ {dead_count} SKUs have zero sales but positive stock — review for liquidation.</div>', unsafe_allow_html=True)
    if len(excess_stock) > 0:
        st.markdown(f'<div class="alert-box alert-yellow">📦 {len(excess_stock)} SKUs have excess stock (>{planning_days*2} days cover) — consider pausing replenishment.</div>', unsafe_allow_html=True)

    # ================== TABS ==================
    tabs = st.tabs([
        "🏠 Overview",
        "📦 FBA Plan",
        "📮 FBM Plan",
        "🏭 FC / Warehouse Plan",
        "🗺️ Cluster View",
        "📊 SKU Summary",
        "🚨 Risk Alerts",
        "📈 Sales Trends",
        "🌍 State Analysis",
        "📄 Amazon Flat File",
    ])

    # ---- OVERVIEW ----
    with tabs[0]:
        st.subheader("🏠 Planning Overview")
        oc1, oc2 = st.columns(2)
        with oc1:
            st.markdown("**📅 Data Summary**")
            st.info(f"""
- **Sales period:** {overall_min.strftime('%d %b %Y')} → {overall_max.strftime('%d %b %Y')} ({uploaded_days} days)
- **Planning basis:** {planning_basis}
- **History window used:** {history_window_days} days
- **Planning horizon:** {planning_days} days
- **Service level:** {service_level} (Z = {z_value})
            """)
        with oc2:
            st.markdown("**🔝 Top 5 SKUs by Dispatch Urgency**")
            top5 = plan[plan["Recommended Dispatch"] > 0].nlargest(5, "Priority Score (0-100)")[
                ["Sku", "Product Name", "Avg Daily Sale", "Days of Cover", "Recommended Dispatch", "Priority Score (0-100)"]
            ]
            st.dataframe(top5, use_container_width=True, hide_index=True)

        st.markdown("**📊 Weekly Sales Trend**")
        if not weekly_trend.empty:
            st.bar_chart(weekly_trend.set_index("Week")["Units Sold"])

        st.markdown("**🗺️ Cluster Dispatch Summary**")
        st.dataframe(cluster_summary, use_container_width=True, hide_index=True)

    # ---- FBA PLAN ----
    with tabs[1]:
        st.subheader("📦 FBA Planning (per SKU)")
        st.caption(f"Showing {len(fba_plan)} FBA SKUs | Planning: {planning_days} days @ {service_level} service level")
        fba_view = fba_plan[fba_plan["Recommended Dispatch"] >= min_dispatch_filter].copy()
        st.dataframe(fba_view.sort_values("Priority Score (0-100)", ascending=False), use_container_width=True)

    # ---- FBM PLAN ----
    with tabs[2]:
        st.subheader("📮 FBM Planning (per SKU)")
        fbm_view = fbm_plan[fbm_plan["Recommended Dispatch"] >= min_dispatch_filter].copy()
        if fbm_view.empty:
            st.info("No FBM SKUs found. All inventory appears to be FBA.")
        else:
            st.dataframe(fbm_view.sort_values("Priority Score (0-100)", ascending=False), use_container_width=True)

    # ---- FC / WAREHOUSE PLAN ----
    with tabs[3]:
        st.subheader("🏭 FC / Warehouse Wise Dispatch Plan")
        st.caption("Units to send per FC, allocated by historical demand share from that location")
        st.dataframe(
            site_plan_view.sort_values(["FC Cluster", "Recommended Dispatch (FC)"], ascending=[True, False]),
            use_container_width=True
        )

    # ---- CLUSTER VIEW ----
    with tabs[4]:
        st.subheader("🗺️ Cluster-Level Inventory View")
        st.dataframe(cluster_summary, use_container_width=True, hide_index=True)
        st.divider()
        for _, row in cluster_summary.iterrows():
            cluster_name = row["FC Cluster"]
            with st.expander(f"📍 {cluster_name} — {fmt_num(int(row['Dispatch Needed']))} units needed"):
                cluster_sites = site_plan[site_plan["FC Cluster"] == cluster_name].copy()
                st.dataframe(
                    cluster_sites[["Warehouse Code", "FC Name", "Sku", "Product Name",
                                   "Current Stock", "Avg Daily Sale (Site)",
                                   "Days of Cover (Site)", "Recommended Dispatch (FC)", "FC Priority"]],
                    use_container_width=True, hide_index=True
                )

    # ---- SKU SUMMARY ----
    with tabs[5]:
        st.subheader("📊 SKU Level Summary (All Channels)")
        st.dataframe(sku_summary.sort_values("Total_Dispatch", ascending=False), use_container_width=True)

    # ---- RISK ALERTS ----
    with tabs[6]:
        r1, r2, r3, r4 = st.tabs(["🔴 Critical (<14d)", "⚫ Dead Stock", "🟡 Slow Moving", "🔵 Excess"])
        with r1:
            st.markdown(f"**{len(critical_stock)} SKUs below 14 days of cover — ORDER IMMEDIATELY**")
            if not critical_stock.empty:
                st.dataframe(critical_stock.sort_values("Days of Cover")[
                    ["Sku", "Product Name", "Fulfillment Type", "Avg Daily Sale",
                     "Current Stock (Channel)", "Days of Cover", "Recommended Dispatch", "Top Sales States"]
                ], use_container_width=True)
        with r2:
            st.markdown(f"**{len(dead_stock)} SKUs with stock but zero sales**")
            if not dead_stock.empty:
                st.dataframe(dead_stock[["Sku", "Product Name", "Fulfillment Type",
                                          "Current Stock (Channel)", "Total Sales (All Time)"]], use_container_width=True)
        with r3:
            st.markdown(f"**{len(slow_moving)} SKUs — Days of Cover > 90 days**")
            if not slow_moving.empty:
                st.dataframe(slow_moving.sort_values("Days of Cover", ascending=False)[
                    ["Sku", "Product Name", "Fulfillment Type", "Avg Daily Sale",
                     "Current Stock (Channel)", "Days of Cover"]
                ], use_container_width=True)
        with r4:
            st.markdown(f"**{len(excess_stock)} SKUs — Cover > {planning_days*2} days (2x planning horizon)**")
            if not excess_stock.empty:
                st.dataframe(excess_stock.sort_values("Days of Cover", ascending=False)[
                    ["Sku", "Product Name", "Fulfillment Type", "Avg Daily Sale",
                     "Current Stock (Channel)", "Days of Cover", "Safety Stock"]
                ], use_container_width=True)

    # ---- SALES TRENDS ----
    with tabs[7]:
        st.subheader("📈 Sales Trends & Intelligence")

        t1, t2, t3 = st.columns(3)
        with t1:
            st.markdown("**🔥 Top 10 Sellers (by Avg Daily Sale)**")
            st.dataframe(top_sellers[["Sku", "Product Name", "Avg Daily Sale", "Days of Cover", "Velocity Tag"]].head(10),
                         use_container_width=True, hide_index=True)
        with t2:
            st.markdown("**📅 Weekly Units Sold**")
            st.line_chart(weekly_trend.set_index("Week")["Units Sold"])
        with t3:
            st.markdown("**🌍 Top States by Sales**")
            st.dataframe(state_to.head(10), use_container_width=True, hide_index=True)

        # SKU-level sales trend
        st.markdown("**🔍 Drill Down: SKU Sales Trend**")
        selected_sku = st.selectbox("Select SKU", sorted(sales["Sku"].unique()))
        sku_daily = (sales[sales["Sku"] == selected_sku]
                     .groupby("Shipment Date")["Quantity"].sum()
                     .reset_index().set_index("Shipment Date"))
        if not sku_daily.empty:
            st.line_chart(sku_daily["Quantity"])

    # ---- STATE ANALYSIS ----
    with tabs[8]:
        st.subheader("🌍 State-Level Sales Analysis")
        sc1, sc2 = st.columns(2)
        with sc1:
            st.markdown("**Ship-To State (Demand)**")
            st.dataframe(state_to, use_container_width=True, hide_index=True)
        with sc2:
            st.markdown("**Ship-From State (Supply)**")
            st.dataframe(state_from, use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("**State × SKU Sales Matrix (Top 20 SKUs)**")
        top20_skus = plan.nlargest(20, "Avg Daily Sale")["Sku"].tolist()
        state_sku = (hist[hist["Sku"].isin(top20_skus)]
                     .groupby(["Ship To State", "Sku"])["Quantity"].sum()
                     .unstack(fill_value=0))
        if not state_sku.empty:
            st.dataframe(state_sku, use_container_width=True)

    # ---- AMAZON FLAT FILE ----
    with tabs[9]:
        st.subheader("📄 Amazon FBA Shipment Flat File Generator")
        st.info("""
        This generates an Amazon-compatible **FBA Shipment Creation Flat File** (Tab-delimited .txt).
        Upload it at: **Seller Central → Manage FBA Shipments → Create New Shipment → Upload a File**.
        
        ⚠️ Review quantities before uploading. Amazon flat files require your exact MSKU/FNSKU mapping.
        """)

        ff_col1, ff_col2, ff_col3 = st.columns(3)
        with ff_col1:
            shipment_name = st.text_input("Shipment Name", f"Replenishment_{datetime.now().strftime('%Y%m%d')}")
            ship_date = st.date_input("Expected Ship Date", datetime.now() + timedelta(days=2))
        with ff_col2:
            prep_type = st.selectbox("Prep Ownership", ["AMAZON", "SELLER"])
            label_type = st.selectbox("Label Ownership", ["AMAZON", "SELLER"])
        with ff_col3:
            ship_method = st.selectbox("Shipping Method", ["SP", "LTL", "SMALL_PARCEL"])
            case_pack = st.checkbox("Case Packed Shipment", value=False)

        # Only include SKUs needing dispatch
        flat_df = fba_plan[fba_plan["Recommended Dispatch"] > 0].copy()

        if flat_df.empty:
            st.warning("No FBA SKUs require dispatch with current settings.")
        else:
            st.markdown(f"**{len(flat_df)} SKUs ready for flat file | Total units: {fmt_num(int(flat_df['Recommended Dispatch'].sum()))}**")

            # FNSKU column (if available in inventory)
            fnsku_col = detect_col(inv, ["FNSKU", "Fnsku", "fnsku"], ["fnsku"])
            if fnsku_col:
                fnsku_map = inv.drop_duplicates("Sku").set_index("Sku")[fnsku_col].to_dict()
                flat_df["FNSKU"] = flat_df["Sku"].map(fnsku_map).fillna("")
            else:
                flat_df["FNSKU"] = ""

            # Units per box
            flat_df["Units Per Box"] = default_case_qty
            flat_df["Number of Boxes"] = np.ceil(flat_df["Recommended Dispatch"] / default_case_qty).astype(int)
            flat_df["Total Units"] = flat_df["Recommended Dispatch"].astype(int)

            # Editable table
            st.markdown("**✏️ Review & Edit Before Generating:**")
            edit_cols = ["Sku", "FNSKU", "Product Name", "Total Units", "Units Per Box", "Number of Boxes"]
            edited = st.data_editor(
                flat_df[edit_cols].reset_index(drop=True),
                num_rows="dynamic",
                use_container_width=True,
            )

            def generate_amazon_flat_file(df):
                """Generate Amazon FBA Shipment Flat File (Tab-separated)"""
                lines = []

                # Header block (Amazon format)
                lines.append("TemplateType=FlatFileShipmentCreation\tVersion=2015.0403")
                lines.append("")  # blank line

                # Metadata
                meta_headers = [
                    "ShipmentName", "ShipFromName", "ShipFromAddressLine1",
                    "ShipFromCity", "ShipFromStateOrProvinceCode", "ShipFromPostalCode",
                    "ShipFromCountryCode", "ShipmentStatus", "LabelPrepType",
                    "AreCasesRequired", "ShipmentId", "DestinationFulfillmentCenterId"
                ]
                addr_parts = ship_from_address.split(",") if ship_from_address else ["", "", ""]
                city = addr_parts[1].strip() if len(addr_parts) > 1 else ""
                state_zip = addr_parts[2].strip() if len(addr_parts) > 2 else ""
                state_code = state_zip.split("-")[0].strip()[:2].upper() if "-" in state_zip else ""
                postal = state_zip.split("-")[1].strip() if "-" in state_zip else ""

                meta_values = [
                    shipment_name,
                    ship_from_name or "My Warehouse",
                    addr_parts[0].strip() if addr_parts else "",
                    city, state_code, postal, "IN",
                    "WORKING", label_type,
                    "YES" if case_pack else "NO",
                    "", ""
                ]
                lines.append("\t".join(meta_headers))
                lines.append("\t".join(str(v) for v in meta_values))
                lines.append("")  # blank line

                # Item headers
                item_headers = [
                    "ShipmentId", "SellerSKU", "FNSKU", "QuantityShipped",
                    "QuantityInCase", "PrepOwner", "LabelOwner",
                    "ItemDescription", "ExpectedDeliveryDate"
                ]
                lines.append("\t".join(item_headers))

                for _, row in df.iterrows():
                    qty_in_case = int(row.get("Units Per Box", default_case_qty)) if case_pack else ""
                    item_row = [
                        "",  # ShipmentId (blank, Amazon fills)
                        str(row["Sku"]),
                        str(row.get("FNSKU", "")),
                        str(int(row["Total Units"])),
                        str(qty_in_case),
                        prep_type,
                        label_type,
                        str(row.get("Product Name", ""))[:200],
                        str(ship_date),
                    ]
                    lines.append("\t".join(item_row))

                return "\n".join(lines)

            flat_file_content = generate_amazon_flat_file(edited)

            # Preview
            with st.expander("👁️ Preview Flat File"):
                st.code(flat_file_content[:3000] + ("\n... (truncated)" if len(flat_file_content) > 3000 else ""))

            # Download
            flat_bytes = flat_file_content.encode("utf-8")
            st.download_button(
                "📥 Download Amazon Flat File (.txt)",
                flat_bytes,
                f"Amazon_FBA_Shipment_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                "text/plain",
            )

            # Also generate FC-wise flat files
            st.divider()
            st.markdown("**📦 FC-Wise Shipment Flat Files**")
            st.caption("Download individual flat files per FC/cluster for split shipments")

            fc_groups = site_plan_view[site_plan_view["Recommended Dispatch (FC)"] > 0].groupby("Warehouse Code")

            for fc_code, fc_data in fc_groups:
                fc_name = fc_display_name(fc_code)
                fc_skus = fc_data[["Sku", "Recommended Dispatch (FC)"]].copy()
                fc_skus.rename(columns={"Recommended Dispatch (FC)": "Total Units"}, inplace=True)
                fc_skus["Product Name"] = fc_skus["Sku"].map(prod_map).fillna("")
                fc_skus["FNSKU"] = ""
                fc_skus["Units Per Box"] = default_case_qty

                fc_flat = generate_amazon_flat_file(fc_skus)
                fc_bytes = fc_flat.encode("utf-8")

                col_fc1, col_fc2 = st.columns([3, 1])
                with col_fc1:
                    st.markdown(f"**{fc_name}** — {fc_skus['Total Units'].sum():.0f} units, {len(fc_skus)} SKUs")
                with col_fc2:
                    st.download_button(
                        f"📥 {fc_code}",
                        fc_bytes,
                        f"Amazon_FBA_{fc_code}_{datetime.now().strftime('%Y%m%d')}.txt",
                        "text/plain",
                        key=f"fc_{fc_code}"
                    )

    # ================== EXCEL EXPORT ==================
    st.markdown("---")
    st.subheader("📥 Download Full Intelligence Report")

    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            wb = writer.book
            header_fmt = wb.add_format({
                'bold': True, 'bg_color': '#2D3748', 'font_color': 'white',
                'border': 1, 'align': 'center', 'valign': 'vcenter'
            })
            green_fmt = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#276221'})
            red_fmt = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            orange_fmt = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})

            sheets = {
                "📊 FBA Plan": fba_plan,
                "📮 FBM Plan": fbm_plan,
                "🏭 FC Dispatch Plan": site_plan,
                "🗺️ Cluster Summary": cluster_summary,
                "📊 SKU Summary": sku_summary,
                "🔴 Critical Stock": critical_stock,
                "⚫ Dead Stock": dead_stock,
                "🟡 Slow Moving": slow_moving,
                "🔵 Excess Stock": excess_stock,
                "🌍 State Sales": state_to,
                "📈 Weekly Trend": weekly_trend,
                "🔥 Top Sellers": top_sellers,
            }

            for sheet_name, df in sheets.items():
                safe_name = sheet_name.replace("📊", "").replace("📮", "").replace("🏭", "").replace(
                    "🗺️", "").replace("🔴", "").replace("⚫", "").replace("🟡", "").replace(
                    "🔵", "").replace("🌍", "").replace("📈", "").replace("🔥", "").strip()[:31]

                if df.empty:
                    pd.DataFrame({"Note": ["No data"]}).to_excel(writer, sheet_name=safe_name, index=False)
                    continue

                df.to_excel(writer, sheet_name=safe_name, index=False, startrow=1)
                ws = writer.sheets[safe_name]

                for col_num, col_name in enumerate(df.columns):
                    ws.write(0, col_num, col_name, header_fmt)
                    col_len = max(len(str(col_name)), df[col_name].astype(str).str.len().max() if not df.empty else 10)
                    ws.set_column(col_num, col_num, min(col_len + 2, 40))

                ws.freeze_panes(2, 0)
                ws.autofilter(0, 0, len(df), len(df.columns) - 1)

        output.seek(0)
        return output

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(
            "📥 Download Full Intelligence Report (Excel)",
            generate_excel(),
            f"FBA_Supply_Intelligence_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Quick CSV download of dispatch plan
    with dl_col2:
        dispatch_csv = fba_plan[fba_plan["Recommended Dispatch"] > 0][
            ["Sku", "Product Name", "Fulfillment Type", "Avg Daily Sale",
             "Current Stock (Channel)", "Days of Cover", "Required Stock", "Recommended Dispatch",
             "Priority Score (0-100)", "Velocity Tag", "Top Sales States"]
        ].to_csv(index=False)
        st.download_button(
            "📋 Download Dispatch Plan (CSV)",
            dispatch_csv,
            f"Dispatch_Plan_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv",
        )

else:
    st.markdown("---")
    st.info("👆 Upload your **Sales MTR** and **Inventory Ledger** files in the sidebar to begin.")
    with st.expander("📋 File Format Requirements"):
        st.markdown("""
        **Sales / MTR File (CSV or ZIP):**
        - `Sku` / `SKU` / `ASIN` — Product identifier
        - `Quantity` — Units sold
        - `Shipment Date` — Date of sale
        - `Ship To State` — Destination state
        - `Fulfilment` / `Fulfillment Channel` — FBA or FBM (optional, defaults to FBA)
        - `Ship From State` — Source state (optional)

        **Inventory Ledger File (CSV or ZIP):**
        - `ASIN` / `MSKU` / `Sku` — Product identifier
        - `Ending Warehouse Balance` — Current stock quantity
        - `Warehouse Code` / `FC Code` — Amazon FC code (optional, for FC-wise planning)
        - `Fulfillment Type` — FBA/FBM (optional)
        - `FNSKU` — Amazon FNSKU (optional, used in flat file)

        **Supported file types:** CSV, ZIP (containing CSV), XLSX
        """)

    with st.expander("🗺️ Supported FC Clusters"):
        cluster_info = {}
        for code, data in FC_DATA.items():
            c = data["cluster"]
            if c not in cluster_info:
                cluster_info[c] = []
            cluster_info[c].append(f"{code} ({data['city']})")
        for cluster, fcs in sorted(cluster_info.items()):
            st.markdown(f"**{cluster}:** {', '.join(fcs)}")
