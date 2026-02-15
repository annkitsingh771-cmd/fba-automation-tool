import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import math

# ================== PAGE CONFIG ==================
st.set_page_config(page_title="FBA Supply Intelligence", layout="wide")
st.title("📦 FBA Supply Intelligence System")

# ================== HELPERS ==================
def add_serial(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure serial number starts from 1 for every table.
    """
    df = df.copy()
    if "S.No" in df.columns:
        df = df.drop(columns=["S.No"])
    df.insert(0, "S.No", range(1, len(df) + 1))
    return df


def read_file(file) -> pd.DataFrame:
    """
    Read CSV directly or first CSV inside a ZIP.
    Returns empty DataFrame if anything fails.
    """
    try:
        if file.name.lower().endswith(".zip"):
            with zipfile.ZipFile(file) as z:
                for name in z.namelist():
                    if name.lower().endswith(".csv"):
                        return pd.read_csv(z.open(name), low_memory=False)
            return pd.DataFrame()
        else:
            return pd.read_csv(file, low_memory=False)
    except Exception:
        return pd.DataFrame()


def detect_column(df: pd.DataFrame, candidates) -> str:
    """
    From a list of candidate column names, return the first that exists in df.
    Comparison is case-insensitive and ignores surrounding spaces.
    """
    normalized = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in normalized:
            return normalized[key]
    return ""


# ================== USER INPUTS ==================
planning_days = st.number_input(
    "Stock Planning Period (Days) – Recommended coverage window",
    min_value=1,
    max_value=180,
    value=30
)

service_level_option = st.selectbox(
    "Service Level (Safety Level)",
    ["90%", "95%", "98%"]
)

z_map = {"90%": 1.28, "95%": 1.65, "98%": 2.05}
z_value = float(z_map[service_level_option])

# ================== FILE UPLOADS ==================
mtr_files = st.file_uploader(
    "Upload MTR (ZIP/CSV) – Sales data",
    type=["csv", "zip"],
    accept_multiple_files=True
)

inventory_file = st.file_uploader(
    "Upload Inventory Ledger (ZIP/CSV) – FBA/MFN stock",
    type=["csv", "zip"]
)

# ================== MAIN LOGIC ==================
if mtr_files and inventory_file:

    # ---------- LOAD & CLEAN SALES ----------
    sales_list = []
    for f in mtr_files:
        df = read_file(f)
        if not df.empty:
            sales_list.append(df)

    if not sales_list:
        st.error("Sales file not readable. Please check the uploaded MTR files.")
        st.stop()

    sales = pd.concat(sales_list, ignore_index=True)

    # Required basic columns
    required_cols = ["Sku", "Quantity", "Shipment Date", "Ship To State"]
    for col in required_cols:
        if col not in sales.columns:
            st.error(f"Missing column in Sales file: {col}")
            st.stop()

    # Base cleaning
    sales["Sku"] = sales["Sku"].astype(str).str.strip()
    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
    sales["Ship To State"] = sales["Ship To State"].astype(str).str.upper().str.strip()

    # Optional: Ship From State (for cluster / split by origin)
    ship_from_col = detect_column(
        sales,
        ["Ship From State", "Ship from State", "Ship_From_State", "Ship From", "Dispatch State"]
    )
    if ship_from_col:
        sales["Ship From State"] = sales[ship_from_col].astype(str).str.upper().str.strip()
    else:
        sales["Ship From State"] = "UNKNOWN"

    # Optional: Fulfilment / Channel (for FBA & MFN split)
    fulfilment_col = detect_column(
        sales,
        [
            "Fulfilment", "Fulfillment",
            "Fulfilment Channel", "Fulfillment Channel",
            "Channel", "Fulfillment Type"
        ]
    )
    if fulfilment_col:
        raw_fulfil = sales[fulfilment_col].astype(str).str.upper().str.strip()
        fulfil_map = {
            "AFN": "FBA",
            "FBA": "FBA",
            "MFN": "MFN",
            "FBM": "MFN"
        }
        sales["Fulfillment Type"] = raw_fulfil.map(fulfil_map).fillna(raw_fulfil)
    else:
        sales["Fulfillment Type"] = "UNKNOWN"

    # Drop rows without valid dates
    sales = sales.dropna(subset=["Shipment Date"])

    if sales.empty:
        st.error("No valid shipment dates found in Sales file.")
        st.stop()

    # ---------- LOAD & CLEAN INVENTORY ----------
    inv = read_file(inventory_file)
    if inv.empty:
        st.error("Inventory file not readable. Please check the uploaded Inventory Ledger.")
        st.stop()

    inv.columns = inv.columns.str.strip()

    # Decide the key to match with Sales Sku: prefer ASIN, else MSKU
    inventory_key_col = ""
    if "ASIN" in inv.columns:
        inventory_key_col = "ASIN"
    elif "MSKU" in inv.columns:
        inventory_key_col = "MSKU"
    else:
        st.error("Inventory file must contain either 'ASIN' or 'MSKU' column.")
        st.stop()

    if "Ending Warehouse Balance" not in inv.columns:
        st.error("Inventory file must contain 'Ending Warehouse Balance' column.")
        st.stop()

    inv[inventory_key_col] = inv[inventory_key_col].astype(str).str.strip()
    inv["Ending Warehouse Balance"] = pd.to_numeric(
        inv["Ending Warehouse Balance"], errors="coerce"
    ).fillna(0)

    # Optional Ship From State or Warehouse State / Code in Inventory
    inv_ship_from_col = detect_column(
        inv,
        ["Ship From State", "Ship from State", "Warehouse State", "FC State"]
    )
    if inv_ship_from_col:
        inv["Ship From State"] = inv[inv_ship_from_col].astype(str).str.upper().str.strip()
    else:
        inv["Ship From State"] = "UNKNOWN"

    # Optional Warehouse Code column (for more granular planning visibility)
    inv_warehouse_col = detect_column(
        inv,
        ["Warehouse Code", "Warehouse", "FC Code", "Fulfillment Center"]
    )
    if inv_warehouse_col:
        inv["Warehouse Code"] = inv[inv_warehouse_col].astype(str).str.upper().str.strip()
    else:
        inv["Warehouse Code"] = "UNKNOWN"

    # ---------- STOCK BY SKU (AGGREGATED) ----------
    # Ensure matching key name with Sales 'Sku'
    sku_stock = (
        inv.groupby(inventory_key_col)["Ending Warehouse Balance"]
        .sum()
        .reset_index()
    )
    sku_stock.rename(
        columns={
            inventory_key_col: "Sku",
            "Ending Warehouse Balance": "Current Stock"
        },
        inplace=True
    )
    sku_stock["Sku"] = sku_stock["Sku"].astype(str).str.strip()

    # ---------- STOCK BY STATE / WAREHOUSE (FOR DOWNLOAD & VIEW) ----------
    stock_by_state_cols = [inventory_key_col, "Ship From State"]
    if "Warehouse Code" in inv.columns:
        stock_by_state_cols.append("Warehouse Code")

    stock_by_state = (
        inv.groupby(stock_by_state_cols)["Ending Warehouse Balance"]
        .sum()
        .reset_index()
    )
    stock_by_state.rename(
        columns={
            inventory_key_col: "Sku",
            "Ending Warehouse Balance": "Current Stock"
        },
        inplace=True
    )
    stock_by_state["Sku"] = stock_by_state["Sku"].astype(str).str.strip()
    stock_by_state = add_serial(stock_by_state)

    # ---------- SALES PERIOD (HISTORICAL COVERAGE) ----------
    sales_period_days = (sales["Shipment Date"].max() - sales["Shipment Date"].min()).days + 1
    if sales_period_days <= 0:
        sales_period_days = 1

    # ---------- PRODUCT-LEVEL PLANNING ----------
    total_sales = (
        sales.groupby("Sku")["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Sales"})
    )

    report = total_sales.merge(sku_stock, on="Sku", how="left")
    report["Current Stock"] = report["Current Stock"].fillna(0)

    # Two planning columns requested:
    # 1. Historical sales coverage (how many days of sales data uploaded)
    # 2. Planning period (for how many days we need recommendation)
    report["Historical Sales Days"] = float(sales_period_days)
    report["Planning Period Days"] = float(planning_days)

    # Average daily sale
    report["Avg Daily Sale"] = report["Total Sales"] / report["Historical Sales Days"]

    # ---------- DEMAND VOLATILITY & SAFETY STOCK ----------
    daily_sales = (
        sales.groupby(["Sku", "Shipment Date"])["Quantity"]
        .sum()
        .reset_index()
    )

    std_dev = (
        daily_sales.groupby("Sku")["Quantity"]
        .std()
        .reset_index()
        .rename(columns={"Quantity": "Demand StdDev"})
    )

    report = report.merge(std_dev, on="Sku", how="left")
    report["Demand StdDev"] = pd.to_numeric(
        report["Demand StdDev"], errors="coerce"
    ).fillna(0)

    # Safety stock using service level (Z-value) and demand volatility
    report["Safety Stock"] = (
        z_value *
        report["Demand StdDev"] *
        math.sqrt(report["Planning Period Days"])
    )

    # Required stock = (avg daily sale * planning period) + safety stock
    report["Required Stock"] = (
        report["Avg Daily Sale"] * report["Planning Period Days"]
    ) + report["Safety Stock"]

    # Recommended dispatch = Required - Current (no negatives)
    report["Recommended Dispatch Qty"] = (
        report["Required Stock"] - report["Current Stock"]
    ).clip(lower=0).round(0)

    # Days of cover at current stock
    report["Days of Cover"] = (
        report["Current Stock"] /
        report["Avg Daily Sale"].replace(0, 1)
    )

    # Final ordering of columns (optional but cleaner)
    cols_order = [
        "S.No", "Sku", "Total Sales",
        "Historical Sales Days", "Planning Period Days",
        "Avg Daily Sale", "Demand StdDev",
        "Safety Stock", "Required Stock",
        "Current Stock", "Recommended Dispatch Qty",
        "Days of Cover"
    ]
    report = add_serial(report)
    # Reorder safely (only keep existing columns in that order)
    existing_cols = [c for c in cols_order if c in report.columns]
    other_cols = [c for c in report.columns if c not in existing_cols]
    report = report[existing_cols + other_cols]

    # ---------- DEAD STOCK ----------
    last_sale = (
        sales.groupby("Sku")["Shipment Date"]
        .max()
        .reset_index()
        .rename(columns={"Shipment Date": "Last Shipment Date"})
    )
    last_sale["Days Since Last Sale"] = (
        sales["Shipment Date"].max() - last_sale["Last Shipment Date"]
    ).dt.days

    dead_stock = last_sale.merge(sku_stock, on="Sku", how="left")
    dead_stock["Current Stock"] = dead_stock["Current Stock"].fillna(0)

    dead_stock = dead_stock[
        (dead_stock["Days Since Last Sale"] > 60) &
        (dead_stock["Current Stock"] > 0)
    ].copy()

    if not dead_stock.empty:
        dead_stock = add_serial(dead_stock)

    # ---------- SLOW MOVING ----------
    slow_moving = report[report["Days of Cover"] > 90].copy()
    if not slow_moving.empty:
        slow_moving = add_serial(slow_moving)

    # ---------- EXCESS STOCK ----------
    excess_stock = report[report["Days of Cover"] > (planning_days * 2)].copy()
    if not excess_stock.empty:
        excess_stock = add_serial(excess_stock)

    # ---------- STATE-LEVEL SALES (TO / FROM) ----------
    state_summary_to = (
        sales.groupby("Ship To State")["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Qty"})
        .sort_values("Total Qty", ascending=False)
    )
    state_summary_to = add_serial(state_summary_to)

    state_summary_from = (
        sales.groupby("Ship From State")["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Qty"})
        .sort_values("Total Qty", ascending=False)
    )
    state_summary_from = add_serial(state_summary_from)

    # ---------- CLUSTER SALES (FBA & MFN SPLIT) ----------
    # Categorized using both Ship To and Ship From states, plus Fulfillment Type.
    cluster_sales = (
        sales.groupby(["Sku", "Ship From State", "Ship To State", "Fulfillment Type"])["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Qty"})
    )
    cluster_sales = add_serial(cluster_sales)

    # ================== DISPLAY BLOCKS ==================
    st.subheader("📊 Main Planning Report (Product Level)")
    st.dataframe(report, use_container_width=True)

    st.subheader("📦 Cluster Sales (FBA & MFN Split)")
    st.dataframe(cluster_sales, use_container_width=True)

    st.subheader("🔥 Dead Stock Alert (No Sale > 60 Days)")
    st.dataframe(dead_stock, use_container_width=True)

    st.subheader("🟡 Slow Moving SKUs (Cover > 90 Days)")
    st.dataframe(slow_moving, use_container_width=True)

    st.subheader("🔵 Excess Stock Warning (Cover > 2× Planning Days)")
    st.dataframe(excess_stock, use_container_width=True)

    st.subheader("🌍 State Sales Summary – Ship To State")
    st.dataframe(state_summary_to, use_container_width=True)

    st.subheader("🏭 State Sales Summary – Ship From State")
    st.dataframe(state_summary_from, use_container_width=True)

    st.subheader("🏷️ Stock by State / Warehouse (From Inventory Ledger)")
    st.dataframe(stock_by_state, use_container_width=True)

    # ================== EXCEL EXPORT ==================
    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            report.to_excel(writer, "Main Planning", index=False)
            dead_stock.to_excel(writer, "Dead Stock", index=False)
            slow_moving.to_excel(writer, "Slow Moving", index=False)
            excess_stock.to_excel(writer, "Excess Stock", index=False)
            state_summary_to.to_excel(writer, "State Sales - Ship To", index=False)
            state_summary_from.to_excel(writer, "State Sales - Ship From", index=False)
            cluster_sales.to_excel(writer, "Cluster Sales (FBA_MFN)", index=False)
            stock_by_state.to_excel(writer, "Stock by State/Warehouse", index=False)
        output.seek(0)
        return output

    st.download_button(
        "📥 Download Full Intelligence Report",
        generate_excel(),
        "FBA_Supply_Intelligence.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Upload Sales (MTR) and Inventory Ledger files to start the analysis.")
