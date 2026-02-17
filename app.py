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
    if df.empty:
        return df
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
    help="For how many future days you want stock at FBA/FBM."
)

safety_buffer = st.slider(
    "Safety Buffer Multiplier (acts like Z‑score)",
    min_value=1.0,
    max_value=3.0,
    value=1.5,
    step=0.1,
    help="Higher value = more safety stock (protection against demand spikes)."
)

planning_basis = st.selectbox(
    "Planning Based On",
    ["Total Sales (full history)", "Last 30 Days", "Last 60 Days", "Last 90 Days"],
    help="Which sales window should be used to calculate average daily demand."
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

    # Standardize types
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

    # Fulfillment / Channel (FBA vs FBM)
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
        # If field is missing, assume all sales are FBA (common for FBA MTR exports)
        sales["Fulfillment Type"] = "FBA"

    # Drop rows without shipment date
    sales = sales.dropna(subset=["Shipment Date"])
    if sales.empty:
        st.error("No valid shipment dates found in Sales file.")
        st.stop()

    # ---------- SALES WINDOW INFO ----------
    overall_min_date = sales["Shipment Date"].min()
    overall_max_date = sales["Shipment Date"].max()
    uploaded_sales_days = (overall_max_date - overall_min_date).days + 1
    uploaded_sales_days = max(uploaded_sales_days, 1)

    # Choose history subset for planning
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

    # Ship From State & Warehouse Code
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
        # If ledger export is FBA-only, consider all as FBA
        inv["Fulfillment Type"] = "FBA"

    # ---------- GLOBAL STOCK BY SKU ----------
    sku_stock = (
        inv.groupby(inventory_key_col)["Ending Warehouse Balance"]
        .sum()
        .reset_index()
    )
    sku_stock.rename(
        columns={
            inventory_key_col: "Sku",
            "Ending Warehouse Balance": "Current Stock (All)",
        },
        inplace=True,
    )
    sku_stock["Sku"] = sku_stock["Sku"].astype(str).str.strip()

    # ---------- STOCK BY SITE (SKU + FT + STATE + WAREHOUSE) ----------
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

    # ---------- STOCK BY SKU + FULFILLMENT TYPE ----------
    stock_by_ft = (
        stock_by_site.groupby(["Sku", "Fulfillment Type"])["Current Stock"]
        .sum()
        .reset_index()
        .rename(columns={"Current Stock": "Current Stock (Channel)"})
    )

    # ---------- DEMAND BY SKU + FULFILLMENT TYPE ----------
    # Total history sales for chosen window per SKU + channel
    hist_sales_ft = (
        hist_sales_df.groupby(["Sku", "Fulfillment Type"])["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "History Sales (Channel)"})
    )

    # Full uploaded sales (for information only, not used in calc)
    full_sales_ft = (
        sales.groupby(["Sku", "Fulfillment Type"])["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Sales (Uploaded, Channel)"})
    )

    # Daily sales per SKU + FT to get channel‑level volatility
    daily_sales_ft = (
        hist_sales_df.groupby(["Sku", "Fulfillment Type", "Shipment Date"])["Quantity"]
        .sum()
        .reset_index()
    )
    std_dev_ft = (
        daily_sales_ft.groupby(["Sku", "Fulfillment Type"])["Quantity"]
        .std()
        .reset_index()
        .rename(columns={"Quantity": "Demand StdDev (Channel)"})
    )

    # ---------- BUILD CHANNEL‑LEVEL PLANNING TABLE ----------
    plan_ft = hist_sales_ft.merge(full_sales_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan_ft = plan_ft.merge(stock_by_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan_ft = plan_ft.merge(std_dev_ft, on=["Sku", "Fulfillment Type"], how="left")

    plan_ft["Current Stock (Channel)"] = plan_ft["Current Stock (Channel)"].fillna(0)
    plan_ft["Demand StdDev (Channel)"] = pd.to_numeric(
        plan_ft["Demand StdDev (Channel)"], errors="coerce"
    ).fillna(0.0)

    # Planning windows (same for all rows)
    plan_ft["Sales Data Days (Uploaded)"] = float(uploaded_sales_days)
    plan_ft["History Window Days (Used)"] = float(history_window_days)
    plan_ft["Planning Period Days"] = float(planning_days)

    # Average daily demand per channel
    plan_ft["Avg Daily Sale (Channel)"] = (
        plan_ft["History Sales (Channel)"] / plan_ft["History Window Days (Used)"]
    )

    # Safety stock per SKU + channel
    # SS = safety_buffer * σ_channel * sqrt(planning_days)
    plan_ft["Safety Stock (Channel)"] = (
        safety_buffer
        * plan_ft["Demand StdDev (Channel)"]
        * math.sqrt(float(planning_days))
    )

    # Required stock per channel
    plan_ft["Base Requirement (Channel)"] = (
        plan_ft["Avg Daily Sale (Channel)"] * float(planning_days)
    )
    plan_ft["Required Stock (Channel)"] = (
        plan_ft["Base Requirement (Channel)"] + plan_ft["Safety Stock (Channel)"]
    )

    # Recommended dispatch per channel (this fixes the earlier issue:
    # we now compare required vs current stock only within that channel)
    plan_ft["Recommended Dispatch (Channel)"] = (
        plan_ft["Required Stock (Channel)"] - plan_ft["Current Stock (Channel)"]
    ).clip(lower=0).round(0)

    # Days of cover per channel
    plan_ft["Days of Cover (Channel)"] = (
        plan_ft["Current Stock (Channel)"]
        / plan_ft["Avg Daily Sale (Channel)"].replace(0, 1)
    )

    # Health tag by channel
    def classify_channel(row):
        if row["Avg Daily Sale (Channel)"] <= 0 and row["Current Stock (Channel)"] > 0:
            return "Dead / No Sales"
        if row["Days of Cover (Channel)"] < planning_days * 0.5:
            return "At Risk (Low Stock)"
        if row["Days of Cover (Channel)"] > planning_days * 2:
            return "Excess / Slow"
        return "Healthy"

    plan_ft["Health Tag"] = plan_ft.apply(classify_channel, axis=1)

    plan_ft = add_serial(plan_ft)

    channel_order = [
        "S.No",
        "Sku",
        "Fulfillment Type",
        "Health Tag",
        "History Sales (Channel)",
        "Total Sales (Uploaded, Channel)",
        "Sales Data Days (Uploaded)",
        "History Window Days (Used)",
        "Planning Period Days",
        "Avg Daily Sale (Channel)",
        "Demand StdDev (Channel)",
        "Safety Stock (Channel)",
        "Base Requirement (Channel)",
        "Required Stock (Channel)",
        "Current Stock (Channel)",
        "Recommended Dispatch (Channel)",
        "Days of Cover (Channel)",
    ]
    existing_cols = [c for c in channel_order if c in plan_ft.columns]
    other_cols = [c for c in plan_ft.columns if c not in existing_cols]
    plan_ft = plan_ft[existing_cols + other_cols]

    # Separate FBA and FBM views
    fba_plan = plan_ft[plan_ft["Fulfillment Type"] == "FBA"].copy()
    fbm_plan = plan_ft[plan_ft["Fulfillment Type"] == "FBM"].copy()

    # ---------- COMBINED SKU SUMMARY (ALL CHANNELS) ----------
    # For quick high‑level view by SKU, sum across channels
    summary_sku = (
        plan_ft.groupby("Sku")
        .agg(
            Channels=("Fulfillment Type", lambda x: ", ".join(sorted(set(x)))),
            Total_History_Sales=("History Sales (Channel)", "sum"),
            Total_Current_Stock=("Current Stock (Channel)", "sum"),
        )
        .reset_index()
    )
    summary_sku = summary_sku.merge(sku_stock, on="Sku", how="left")
    summary_sku["Current Stock (All)"] = summary_sku["Current Stock (All)"].fillna(0)
    summary_sku = add_serial(summary_sku)

    # ---------- STATE‑LEVEL SALES ----------
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

    # ---------- MULTI‑WAREHOUSE DIAGNOSTIC ----------
    # Use demand per site (if Ship From State present in MTR) to compute local days of cover
    site_sales = (
        hist_sales_df.groupby(
            ["Sku", "Fulfillment Type", "Ship From State"]
        )["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "History Sales (Site)"})
    )
    site_sales["Avg Daily Sale (Site)"] = (
        site_sales["History Sales (Site)"] / float(history_window_days)
    )

    site_plan = stock_by_site.merge(
        site_sales, on=["Sku", "Fulfillment Type", "Ship From State"], how="left"
    )
    site_plan["History Sales (Site)"] = site_plan["History Sales (Site)"].fillna(0)
    site_plan["Avg Daily Sale (Site)"] = site_plan["Avg Daily Sale (Site)"].fillna(0)

    site_plan["Days of Cover (Site)"] = (
        site_plan["Current Stock"]
        / site_plan["Avg Daily Sale (Site)"].replace(0, 1)
    )

    site_plan = add_serial(site_plan)

    # ---------- RISK TABLES ----------
    dead_stock = plan_ft[
        (plan_ft["Avg Daily Sale (Channel)"] == 0)
        & (plan_ft["Current Stock (Channel)"] > 0)
    ].copy()
    dead_stock = add_serial(dead_stock)

    slow_moving = plan_ft[plan_ft["Days of Cover (Channel)"] > 90].copy()
    slow_moving = add_serial(slow_moving)

    excess_stock = plan_ft[plan_ft["Days of Cover (Channel)"] > (planning_days * 2)].copy()
    excess_stock = add_serial(excess_stock)

    # ================== UI TABS ==================
    tab_fba, tab_fbm, tab_all, tab_sites, tab_risk = st.tabs(
        ["FBA Planning", "FBM Planning", "All Channels Summary", "Warehouses / States", "Risk Alerts"]
    )

    with tab_fba:
        st.subheader("📦 FBA Planning (per SKU)")
        st.dataframe(fba_plan, use_container_width=True)

    with tab_fbm:
        st.subheader("📮 FBM Planning (per SKU)")
        st.dataframe(fbm_plan, use_container_width=True)

    with tab_all:
        st.subheader("📊 Combined SKU Overview (All Channels)")
        st.dataframe(summary_sku, use_container_width=True)
        st.caption(
            "This table shows total history sales and total stock across FBA + FBM. "
            "Use FBA/FBM tabs for channel‑specific dispatch recommendations."
        )

    with tab_sites:
        st.subheader("🏭 Stock by Warehouse / State (FBA & FBM)")
        st.dataframe(site_plan, use_container_width=True)

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🌍 State Sales – Ship To State")
            st.dataframe(state_summary_to, use_container_width=True)
        with col2:
            st.subheader("🚚 State Sales – Ship From State")
            st.dataframe(state_summary_from, use_container_width=True)

    with tab_risk:
        st.subheader("🔥 Dead / Zero‑Velocity Stock (Channel Level)")
        st.dataframe(dead_stock, use_container_width=True)

        st.subheader("🟡 Slow Moving (Cover > 90 Days)")
        st.dataframe(slow_moving, use_container_width=True)

        st.subheader("🔵 Excess Stock (Cover > 2 × Planning Days)")
        st.dataframe(excess_stock, use_container_width=True)

    # ================== EXCEL EXPORT ==================
    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            plan_ft.to_excel(writer, "Planning_All_Channels", index=False)
            fba_plan.to_excel(writer, "Planning_FBA", index=False)
            fbm_plan.to_excel(writer, "Planning_FBM", index=False)
            summary_sku.to_excel(writer, "Summary_SKU_All", index=False)
            site_plan.to_excel(writer, "Warehouse_State_View", index=False)
            state_summary_to.to_excel(writer, "State_Sales_ShipTo", index=False)
            state_summary_from.to_excel(writer, "State_Sales_ShipFrom", index=False)
            dead_stock.to_excel(writer, "Risk_Dead", index=False)
            slow_moving.to_excel(writer, "Risk_Slow", index=False)
            excess_stock.to_excel(writer, "Risk_Excess", index=False)
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
