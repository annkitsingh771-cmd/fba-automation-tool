import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import math

# ================== PAGE CONFIG ==================
st.set_page_config(page_title="FBA Smart Supply Planner", layout="wide")
st.title("📦 FBA Smart Supply Planner")


# ================== HELPER FUNCTIONS ==================
def add_serial(df: pd.DataFrame) -> pd.DataFrame:
    """Add S.No column starting at 1."""
    df = df.copy()
    if "S.No" in df.columns:
        df = df.drop(columns=["S.No"])
    df.insert(0, "S.No", range(1, len(df) + 1))
    return df


def read_file(file) -> pd.DataFrame:
    """Read CSV directly or first CSV inside a ZIP."""
    try:
        name = file.name.lower()
        if name.endswith(".zip"):
            with zipfile.ZipFile(file) as z:
                for member in z.namelist():
                    if member.lower().endswith(".csv"):
                        return pd.read_csv(z.open(member), low_memory=False)
            return pd.DataFrame()
        else:
            return pd.read_csv(file, low_memory=False)
    except Exception:
        return pd.DataFrame()


def detect_column(df: pd.DataFrame, candidates) -> str:
    """Return first matching column name from candidates (case/space insensitive)."""
    normalized = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in normalized:
            return normalized[key]
    return ""


# ================== USER INPUTS ==================
st.markdown("### 🔧 Planning Controls")

planning_days = st.slider(
    "Planning Days (Stock coverage target)",
    min_value=7,
    max_value=180,
    value=60,
    step=1,
)

safety_buffer = st.slider(
    "Safety Buffer Multiplier (acts like Z‑score)",
    min_value=1.0,
    max_value=3.0,
    value=1.5,
    step=0.1,
)

planning_basis = st.selectbox(
    "Planning Based On",
    ["Total Sales (full history)", "Last 30 Days", "Last 60 Days", "Last 90 Days"],
)

st.markdown("---")

mtr_files = st.file_uploader(
    "Upload MTR (ZIP/CSV) – Sales data",
    type=["csv", "zip"],
    accept_multiple_files=True,
)

inventory_file = st.file_uploader(
    "Upload Inventory Ledger (ZIP/CSV) – FBA & FBM stock",
    type=["csv", "zip"],
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

    # Basic required columns
    base_required_cols = ["Sku", "Quantity", "Shipment Date", "Ship To State"]
    for col in base_required_cols:
        if col not in sales.columns:
            st.error(f"Missing column in Sales file: {col}")
            st.stop()

    # Standardize
    sales["Sku"] = sales["Sku"].astype(str).str.strip()
    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
    sales["Ship To State"] = (
        sales["Ship To State"].astype(str).str.upper().str.strip()
    )

    # Optional Ship From State
    ship_from_col = detect_column(
        sales,
        [
            "Ship From State",
            "Ship from State",
            "Ship_From_State",
            "Ship From",
            "Dispatch State",
        ],
    )
    if ship_from_col:
        sales["Ship From State"] = (
            sales[ship_from_col].astype(str).str.upper().str.strip()
        )
    else:
        sales["Ship From State"] = "UNKNOWN"

    # Optional Fulfillment / Channel (for FBA vs FBM split)
    fulfilment_col = detect_column(
        sales,
        [
            "Fulfilment",
            "Fulfillment",
            "Fulfilment Channel",
            "Fulfillment Channel",
            "Channel",
            "Fulfillment Type",
        ],
    )
    if fulfilment_col:
        raw_fulfil = sales[fulfilment_col].astype(str).str.upper().str.strip()
        fulfil_map = {
            "AFN": "FBA",
            "FBA": "FBA",
            "AMAZON_FULFILLED": "FBA",
            "MFN": "FBM",
            "FBM": "FBM",
            "MERCHANT_FULFILLED": "FBM",
        }
        sales["Fulfillment Type"] = raw_fulfil.map(fulfil_map).fillna(raw_fulfil)
    else:
        # If not present, treat all sales as FBA (typical if only FBA MTR is uploaded)
        sales["Fulfillment Type"] = "FBA"

    # Remove rows without shipment date
    sales = sales.dropna(subset=["Shipment Date"])
    if sales.empty:
        st.error("No valid shipment dates found in Sales file.")
        st.stop()

    # ---------- SALES WINDOW INFO ----------
    overall_min_date = sales["Shipment Date"].min()
    overall_max_date = sales["Shipment Date"].max()
    uploaded_sales_days = (overall_max_date - overall_min_date).days + 1
    uploaded_sales_days = max(uploaded_sales_days, 1)

    # Choose history subset based on planning_basis
    if planning_basis == "Total Sales (full history)":
        hist_sales_df = sales.copy()
        history_window_days = uploaded_sales_days
    else:
        if "30" in planning_basis:
            window = 30
        elif "60" in planning_basis:
            window = 60
        else:
            window = 90
        cutoff = overall_max_date - pd.Timedelta(days=window - 1)
        hist_sales_df = sales[sales["Shipment Date"] >= cutoff].copy()
        if hist_sales_df.empty:
            hist_sales_df = sales.copy()
            history_window_days = uploaded_sales_days
        else:
            hmin = hist_sales_df["Shipment Date"].min()
            hmax = hist_sales_df["Shipment Date"].max()
            history_window_days = (hmax - hmin).days + 1
            history_window_days = max(history_window_days, 1)

    # ---------- LOAD & CLEAN INVENTORY ----------
    inv = read_file(inventory_file)
    if inv.empty:
        st.error("Inventory file not readable. Please check the uploaded Inventory Ledger.")
        st.stop()

    inv.columns = inv.columns.str.strip()

    # Key identifier: ASIN preferred, else MSKU
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

    # Optional state / warehouse info
    inv_ship_from_col = detect_column(
        inv,
        ["Ship From State", "Ship from State", "Warehouse State", "FC State"],
    )
    if inv_ship_from_col:
        inv["Ship From State"] = (
            inv[inv_ship_from_col].astype(str).str.upper().str.strip()
        )
    else:
        inv["Ship From State"] = "UNKNOWN"

    inv_warehouse_col = detect_column(
        inv, ["Warehouse Code", "Warehouse", "FC Code", "Fulfillment Center"]
    )
    if inv_warehouse_col:
        inv["Warehouse Code"] = (
            inv[inv_warehouse_col].astype(str).str.upper().str.strip()
        )
    else:
        inv["Warehouse Code"] = "UNKNOWN"

    # Fulfillment type on inventory (FBA vs FBM)
    inv_fulfil_col = detect_column(
        inv,
        [
            "Fulfilment",
            "Fulfillment",
            "Fulfilment Channel",
            "Fulfillment Channel",
            "Channel",
            "Fulfillment Type",
            "Inventory Type",
        ],
    )
    if inv_fulfil_col:
        raw_inv_fulfil = inv[inv_fulfil_col].astype(str).str.upper().str.strip()
        fulfil_map_inv = {
            "AFN": "FBA",
            "FBA": "FBA",
            "AMAZON_FULFILLED": "FBA",
            "MFN": "FBM",
            "FBM": "FBM",
            "MERCHANT_FULFILLED": "FBM",
        }
        inv["Fulfillment Type"] = raw_inv_fulfil.map(fulfil_map_inv).fillna(
            raw_inv_fulfil
        )
    else:
        # If ledger is FBA-only (common case), mark all as FBA
        inv["Fulfillment Type"] = "FBA"

    # ---------- STOCK BY SKU (AGGREGATE) ----------
    sku_stock = (
        inv.groupby(inventory_key_col)["Ending Warehouse Balance"]
        .sum()
        .reset_index()
    )
    sku_stock.rename(
        columns={
            inventory_key_col: "Sku",
            "Ending Warehouse Balance": "Current Stock",
        },
        inplace=True,
    )
    sku_stock["Sku"] = sku_stock["Sku"].astype(str).str.strip()

    # ---------- STOCK BY STATE / WAREHOUSE & FULFILLMENT ----------
    stock_by_site = (
        inv.groupby(
            [inventory_key_col, "Fulfillment Type", "Ship From State", "Warehouse Code"]
        )["Ending Warehouse Balance"]
        .sum()
        .reset_index()
    )
    stock_by_site.rename(
        columns={
            inventory_key_col: "Sku",
            "Ending Warehouse Balance": "Current Stock",
        },
        inplace=True,
    )
    stock_by_site["Sku"] = stock_by_site["Sku"].astype(str).str.strip()

    # Also an aggregate by Sku + Fulfillment Type
    stock_by_ft = (
        stock_by_site.groupby(["Sku", "Fulfillment Type"])["Current Stock"]
        .sum()
        .reset_index()
        .rename(columns={"Current Stock": "Current Stock (FT)"})
    )

    # ---------- PRODUCT-LEVEL DEMAND (HISTORY WINDOW) ----------
    hist_total_sales = (
        hist_sales_df.groupby("Sku")["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "History Sales"})
    )

    # For reference: total sales over full uploaded window
    full_total_sales = (
        sales.groupby("Sku")["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Sales (Uploaded)"})
    )

    report = hist_total_sales.merge(full_total_sales, on="Sku", how="left")
    report = report.merge(sku_stock, on="Sku", how="left")
    report["Current Stock"] = report["Current Stock"].fillna(0)

    # Two planning columns (explicit requirement)
    report["Sales Data Days (Uploaded)"] = float(uploaded_sales_days)
    report["Planning Period Days"] = float(planning_days)

    # History window actually used for averages
    report["History Window Days (Used)"] = float(history_window_days)
    report["Avg Daily Sale"] = (
        report["History Sales"] / report["History Window Days (Used)"]
    )

    # ---------- DEMAND VOLATILITY & SAFETY STOCK ----------
    daily_sales = (
        hist_sales_df.groupby(["Sku", "Shipment Date"])["Quantity"]
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
    report["Demand StdDev"] = (
        pd.to_numeric(report["Demand StdDev"], errors="coerce").fillna(0.0)
    )

    # Safety stock using safety_buffer as Z‑factor: SS = k * σ * sqrt(T)
    report["Safety Stock"] = (
        safety_buffer
        * report["Demand StdDev"]
        * math.sqrt(float(planning_days))
    )

    # Required stock & recommendation (global per SKU)
    report["Base Requirement"] = report["Avg Daily Sale"] * float(planning_days)
    report["Required Stock"] = report["Base Requirement"] + report["Safety Stock"]

    report["Recommended Dispatch Qty"] = (
        report["Required Stock"] - report["Current Stock"]
    ).clip(lower=0).round(0)

    # Days of cover on current stock
    report["Days of Cover"] = (
        report["Current Stock"] / report["Avg Daily Sale"].replace(0, 1)
    )

    # Health classification
    def classify_row(row):
        if row["Avg Daily Sale"] <= 0 and row["Current Stock"] > 0:
            return "Dead / No Sales"
        if row["Days of Cover"] < planning_days * 0.5:
            return "At Risk (Low Stock)"
        if row["Days of Cover"] > planning_days * 2:
            return "Excess / Slow"
        return "Healthy"

    report["Health Tag"] = report.apply(classify_row, axis=1)

    report = add_serial(report)
    preferred_order = [
        "S.No",
        "Sku",
        "Health Tag",
        "History Sales",
        "Total Sales (Uploaded)",
        "Sales Data Days (Uploaded)",
        "History Window Days (Used)",
        "Planning Period Days",
        "Avg Daily Sale",
        "Demand StdDev",
        "Safety Stock",
        "Base Requirement",
        "Required Stock",
        "Current Stock",
        "Recommended Dispatch Qty",
        "Days of Cover",
    ]
    existing = [c for c in preferred_order if c in report.columns]
    rest = [c for c in report.columns if c not in existing]
    report = report[existing + rest]

    # ---------- DEAD STOCK ----------
    last_sale = (
        sales.groupby("Sku")["Shipment Date"]
        .max()
        .reset_index()
        .rename(columns={"Shipment Date": "Last Shipment Date"})
    )
    last_sale["Days Since Last Sale"] = (
        overall_max_date - last_sale["Last Shipment Date"]
    ).dt.days

    dead_stock = last_sale.merge(sku_stock, on="Sku", how="left")
    dead_stock["Current Stock"] = dead_stock["Current Stock"].fillna(0)

    dead_stock = dead_stock[
        (dead_stock["Days Since Last Sale"] > 60)
        & (dead_stock["Current Stock"] > 0)
    ].copy()
    if not dead_stock.empty:
        dead_stock = add_serial(dead_stock)

    # ---------- SLOW / EXCESS STOCK ----------
    slow_moving = report[report["Days of Cover"] > 90].copy()
    if not slow_moving.empty:
        slow_moving = add_serial(slow_moving)

    excess_stock = report[report["Days of Cover"] > (planning_days * 2)].copy()
    if not excess_stock.empty:
        excess_stock = add_serial(excess_stock)

    # ---------- STATE-LEVEL SALES ----------
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

    # ---------- DEMAND BY FULFILLMENT TYPE (FBA vs FBM) ----------
    demand_ft = (
        hist_sales_df.groupby(["Sku", "Fulfillment Type"])["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "History Sales (FT)"})
    )

    total_hist_per_sku = (
        demand_ft.groupby("Sku")["History Sales (FT)"]
        .sum()
        .reset_index()
        .rename(columns={"History Sales (FT)": "Total History Sales"})
    )

    demand_ft = demand_ft.merge(total_hist_per_sku, on="Sku", how="left")
    demand_ft["Demand Share (FT)"] = np.where(
        demand_ft["Total History Sales"] > 0,
        demand_ft["History Sales (FT)"] / demand_ft["Total History Sales"],
        0.0,
    )

    # Attach SKU-level required stock and compute allocation per fulfillment type
    demand_ft = demand_ft.merge(
        report[["Sku", "Required Stock"]], on="Sku", how="left"
    )
    demand_ft["Allocated Required Stock (FT)"] = (
        demand_ft["Required Stock"] * demand_ft["Demand Share (FT)"]
    )

    # FBA / FBM summary per SKU
    fba_fbm_summary = (
        demand_ft.groupby(["Sku", "Fulfillment Type"])
        .agg(
            History_Sales_FT=("History Sales (FT)", "sum"),
            Demand_Share_FT=("Demand Share (FT)", "mean"),
            Required_Stock_FT=("Allocated Required Stock (FT)", "sum"),
        )
        .reset_index()
    )

    fba_fbm_summary = fba_fbm_summary.merge(
        stock_by_ft, on=["Sku", "Fulfillment Type"], how="left"
    )
    fba_fbm_summary["Current Stock (FT)"] = fba_fbm_summary[
        "Current Stock (FT)"
    ].fillna(0)

    fba_fbm_summary["Recommended Dispatch (FT)"] = (
        fba_fbm_summary["Required_Stock_FT"]
        - fba_fbm_summary["Current Stock (FT)"]
    ).clip(lower=0).round(0)

    fba_fbm_summary = add_serial(fba_fbm_summary)

    # ---------- MULTI-WAREHOUSE PLAN (SKU + FT + SITE) ----------
    multi_wh_plan = stock_by_site.merge(
        demand_ft[
            ["Sku", "Fulfillment Type", "Allocated Required Stock (FT)"]
        ],
        on=["Sku", "Fulfillment Type"],
        how="left",
    )

    multi_wh_plan["Allocated Required Stock (FT)"] = multi_wh_plan[
        "Allocated Required Stock (FT)"
    ].fillna(0.0)

    multi_wh_plan["Recommended Dispatch (Site)"] = (
        multi_wh_plan["Allocated Required Stock (FT)"]
        - multi_wh_plan["Current Stock"]
    ).clip(lower=0).round(0)

    multi_wh_plan = add_serial(multi_wh_plan)

    # ================== DISPLAY ==================
    st.subheader("📊 Main Planning Report (SKU Level)")
    st.dataframe(report, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🌍 State Sales – Ship To State")
        st.dataframe(state_summary_to, use_container_width=True)
    with col2:
        st.subheader("🏭 State Sales – Ship From State")
        st.dataframe(state_summary_from, use_container_width=True)

    st.subheader("🏬 FBA / FBM Summary by SKU & Fulfillment Type")
    st.dataframe(fba_fbm_summary, use_container_width=True)

    st.subheader("📦 Multi‑Warehouse Plan (Per Site & Fulfillment Type)")
    st.dataframe(multi_wh_plan, use_container_width=True)

    st.subheader("🏷️ Stock by State / Warehouse (Raw Inventory View)")
    st.dataframe(add_serial(stock_by_site.copy()), use_container_width=True)

    st.subheader("🔥 Dead Stock (No Sale > 60 Days)")
    st.dataframe(dead_stock, use_container_width=True)

    st.subheader("🟡 Slow Moving SKUs (Cover > 90 Days)")
    st.dataframe(slow_moving, use_container_width=True)

    st.subheader("🔵 Excess Stock (Cover > 2 × Planning Days)")
    st.dataframe(excess_stock, use_container_width=True)

    # ================== EXCEL EXPORT ==================
    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            report.to_excel(writer, "Main Planning", index=False)
            dead_stock.to_excel(writer, "Dead Stock", index=False)
            slow_moving.to_excel(writer, "Slow Moving", index=False)
            excess_stock.to_excel(writer, "Excess Stock", index=False)
            state_summary_to.to_excel(writer, "State Sales - Ship To", index=False)
            state_summary_from.to_excel(
                writer, "State Sales - Ship From", index=False
            )
            fba_fbm_summary.to_excel(
                writer, "FBA_FBM by SKU", index=False
            )
            multi_wh_plan.to_excel(
                writer, "Multi-WH Plan (Site)", index=False
            )
            stock_by_site.to_excel(
                writer, "Stock by Site (Raw)", index=False
            )
        output.seek(0)
        return output

    st.download_button(
        "📥 Download Full Intelligence Report",
        generate_excel(),
        "FBA_Supply_Intelligence.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Upload Sales (MTR) and Inventory Ledger files to start the planner.")
