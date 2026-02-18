import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import math

# ================== PAGE CONFIG ==================
st.set_page_config(page_title="FBA Smart Supply Planner", layout="wide")
st.title("📦 FBA Smart Supply Planner")

# ================== FC NAME MAPPING (CUSTOMISE AS NEEDED) ==================
FC_NAMES = {
    'DEL4': 'Delhi North FC',
    'DEL5': 'Delhi North FC',
    'DEX3': 'New Delhi FC',
    'PNQ2': 'Delhi City FC',
    'BOM1': 'Bhiwandi Mumbai FC',
    'BOM3': 'Nashik FC',
    'BOM4': 'Vasai FC',
    'SAMB': 'Mumbai West FC',
    'BLR5': 'Bangalore South FC',
    'SCJA': 'Bangalore FC',
    'XSAB': 'Bangalore XS FC',
    'SBLA': 'Bangalore SBLA FC',
    'MAA4': 'Chennai South FC',
    'MAA5': 'Chennai South FC',
    'SMAB': 'Chennai SMAB FC',
    'HYD7': 'Hyderabad South FC',
    'HYD8': 'Hyderabad South FC',
}


def fc_display_name(code: str) -> str:
    code = str(code).upper().strip()
    if not code or code == "UNKNOWN":
        return "Unknown FC"
    return FC_NAMES.get(code, f"FC {code}")


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


def detect_column_exact(df: pd.DataFrame, candidates) -> str:
    """First try exact name match (case/space insensitive)."""
    normalized = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in normalized:
            return normalized[key]
    return ""


def detect_column_fuzzy(df: pd.DataFrame, substrings) -> str:
    """
    Fuzzy detect column if exact match failed: looks for any substring
    in the column name (case-insensitive).
    """
    for col in df.columns:
        low = col.lower()
        if any(sub in low for sub in substrings):
            return col
    return ""


# ================== USER INPUTS ==================
st.markdown("### 🔧 Planning Controls")

planning_days = st.number_input(
    "Planning Days (Stock coverage target)",
    min_value=7,
    max_value=180,
    value=60,
    step=1,
)

service_level = st.selectbox(
    "Service Level (Safety Level)",
    ["90%", "95%", "98%"],
)
Z_MAP = {"90%": 1.28, "95%": 1.65, "98%": 2.05}
z_value = float(Z_MAP[service_level])

planning_basis = st.selectbox(
    "Planning Based On",
    ["Total Sales (full history)", "Last 30 Days", "Last 60 Days", "Last 90 Days"],
)

st.markdown("---")

mtr_files = st.file_uploader(
    "Upload MTR (ZIP/CSV)", type=["csv", "zip"], accept_multiple_files=True
)
inventory_file = st.file_uploader(
    "Upload Inventory Ledger (ZIP/CSV)", type=["csv", "zip"]
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
        st.error("Sales file not readable. Please check uploaded MTR files.")
        st.stop()

    sales = pd.concat(sales_list, ignore_index=True)

    base_required_cols = ["Sku", "Quantity", "Shipment Date", "Ship To State"]
    for col in base_required_cols:
        if col not in sales.columns:
            st.error(f"Missing column in Sales file: {col}")
            st.stop()

    sales["Sku"] = sales["Sku"].astype(str).str.strip()
    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
    sales["Ship To State"] = (
        sales["Ship To State"].astype(str).str.upper().str.strip()
    )

    # Fulfillment Type (FBA / FBM)
    fulfil_col = detect_column_exact(
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
    if fulfil_col:
        raw_ft = sales[fulfil_col].astype(str).str.upper().str.strip()
        ft_map = {
            "AFN": "FBA",
            "FBA": "FBA",
            "AMAZON_FULFILLED": "FBA",
            "MFN": "FBM",
            "FBM": "FBM",
            "MERCHANT_FULFILLED": "FBM",
        }
        sales["Fulfillment Type"] = raw_ft.map(ft_map).fillna(raw_ft)
    else:
        sales["Fulfillment Type"] = "FBA"

    # Ship From State (if present in MTR)
    sf_exact = detect_column_exact(
        sales,
        ["Ship From State", "Ship from State", "Ship_From_State", "Dispatch State"],
    )
    if not sf_exact:
        sf_exact = detect_column_fuzzy(sales, ["ship from", "dispatch", "origin"])
    if sf_exact:
        sales["Ship From State"] = (
            sales[sf_exact].astype(str).str.upper().str.strip()
        )
    else:
        sales["Ship From State"] = "UNKNOWN"

    sales = sales.dropna(subset=["Shipment Date"])
    if sales.empty:
        st.error("No valid shipment dates found in Sales file.")
        st.stop()

    # ---------- SALES WINDOW INFO ----------
    overall_min = sales["Shipment Date"].min()
    overall_max = sales["Shipment Date"].max()
    uploaded_sales_days = (overall_max - overall_min).days + 1
    uploaded_sales_days = max(uploaded_sales_days, 1)

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
        cutoff = overall_max - pd.Timedelta(days=window - 1)
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
        st.error("Inventory file not readable.")
        st.stop()

    inv.columns = inv.columns.str.strip()

    # Decide inventory key column
    if "ASIN" in inv.columns:
        inventory_key = "ASIN"
    elif "MSKU" in inv.columns:
        inventory_key = "MSKU"
    elif "Sku" in inv.columns:
        inventory_key = "Sku"
    else:
        st.error("Inventory must contain ASIN, MSKU or Sku column.")
        st.stop()

    if "Ending Warehouse Balance" not in inv.columns:
        st.error("Inventory must contain 'Ending Warehouse Balance' column.")
        st.stop()

    inv[inventory_key] = inv[inventory_key].astype(str).str.strip()
    inv["Sku"] = inv[inventory_key]
    inv["Ending Warehouse Balance"] = pd.to_numeric(
        inv["Ending Warehouse Balance"], errors="coerce"
    ).fillna(0)

    # Fulfillment Type in inventory
    inv_ft_col = detect_column_exact(
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
    if not inv_ft_col:
        inv_ft_col = detect_column_fuzzy(inv, ["fulfil", "fulfill", "channel"])
    if inv_ft_col:
        raw_inv_ft = inv[inv_ft_col].astype(str).str.upper().str.strip()
        ft_map_inv = {
            "AFN": "FBA",
            "FBA": "FBA",
            "AMAZON_FULFILLED": "FBA",
            "MFN": "FBM",
            "FBM": "FBM",
            "MERCHANT_FULFILLED": "FBM",
        }
        inv["Fulfillment Type"] = raw_inv_ft.map(ft_map_inv).fillna(raw_inv_ft)
    else:
        inv["Fulfillment Type"] = "FBA"

    # Ship From State (inventory)
    inv_sf_col = detect_column_exact(
        inv,
        ["Ship From State", "Ship from State", "Warehouse State", "FC State"],
    )
    if not inv_sf_col:
        inv_sf_col = detect_column_fuzzy(inv, ["state", "ship from", "warehouse"])
    if inv_sf_col:
        inv["Ship From State"] = (
            inv[inv_sf_col].astype(str).str.upper().str.strip()
        )
    else:
        inv["Ship From State"] = "UNKNOWN"

    # Warehouse Code / FC
    inv_fc_col = detect_column_exact(
        inv, ["Warehouse Code", "Warehouse", "FC Code", "Fulfillment Center"]
    )
    if not inv_fc_col:
        inv_fc_col = detect_column_fuzzy(inv, ["warehouse", "fc", "fulfillment"])
    if inv_fc_col:
        inv["Warehouse Code"] = (
            inv[inv_fc_col].astype(str).str.upper().str.strip()
        )
    else:
        inv["Warehouse Code"] = "UNKNOWN"

    # ---------- STOCK AGGREGATES ----------
    # Global stock (all channels) per SKU
    sku_stock_all = (
        inv.groupby("Sku")["Ending Warehouse Balance"]
        .sum()
        .reset_index()
        .rename(columns={"Ending Warehouse Balance": "Current Stock (All)"}))
    # Stock per SKU + Channel
    stock_by_ft = (
        inv.groupby(["Sku", "Fulfillment Type"])["Ending Warehouse Balance"]
        .sum()
        .reset_index()
        .rename(columns={"Ending Warehouse Balance": "Current Stock (Channel)"}))

    # Stock per site (SKU + FT + State + Warehouse)
    stock_by_site = (
        inv.groupby(
            ["Sku", "Fulfillment Type", "Ship From State", "Warehouse Code"]
        )["Ending Warehouse Balance"]
        .sum()
        .reset_index()
        .rename(columns={"Ending Warehouse Balance": "Current Stock"}))

    stock_by_site["FC Name"] = stock_by_site["Warehouse Code"].apply(fc_display_name)

    # ---------- DEMAND BY SKU + CHANNEL ----------
    hist_sales_ft = (
        hist_sales_df.groupby(["Sku", "Fulfillment Type"])["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "History Sales (Channel)"}))

    full_sales_ft = (
        sales.groupby(["Sku", "Fulfillment Type"])["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Sales (Uploaded, Channel)"}))

    daily_sales_ft = (
        hist_sales_df.groupby(
            ["Sku", "Fulfillment Type", "Shipment Date"]
        )["Quantity"]
        .sum()
        .reset_index())

    std_dev_ft = (
        daily_sales_ft.groupby(["Sku", "Fulfillment Type"])["Quantity"]
        .std()
        .reset_index()
        .rename(columns={"Quantity": "Demand StdDev (Channel)"}))

    # ---------- BUILD CHANNEL-LEVEL PLANNING TABLE ----------
    # Start from union of keys from sales and inventory
    base_keys = pd.concat(
        [
            hist_sales_ft[["Sku", "Fulfillment Type"]],
            stock_by_ft[["Sku", "Fulfillment Type"]],
        ],
        ignore_index=True,
    ).drop_duplicates()

    plan_ft = base_keys.merge(hist_sales_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan_ft = plan_ft.merge(full_sales_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan_ft = plan_ft.merge(stock_by_ft, on=["Sku", "Fulfillment Type"], how="left")
    plan_ft = plan_ft.merge(std_dev_ft, on=["Sku", "Fulfillment Type"], how="left")

    plan_ft["History Sales (Channel)"] = plan_ft["History Sales (Channel)"].fillna(0)
    plan_ft["Total Sales (Uploaded, Channel)"] = plan_ft[
        "Total Sales (Uploaded, Channel)"
    ].fillna(0)
    plan_ft["Current Stock (Channel)"] = plan_ft["Current Stock (Channel)"].fillna(0)
    plan_ft["Demand StdDev (Channel)"] = pd.to_numeric(
        plan_ft["Demand StdDev (Channel)"], errors="coerce"
    ).fillna(0.0)

    plan_ft["Sales Data Days (Uploaded)"] = float(uploaded_sales_days)
    plan_ft["History Window Days (Used)"] = float(history_window_days)
    plan_ft["Planning Period Days"] = float(planning_days)

    plan_ft["Avg Daily Sale (Channel)"] = (
        plan_ft["History Sales (Channel)"] / plan_ft["History Window Days (Used)"]
    )

    safety_sqrt = math.sqrt(float(planning_days))
    plan_ft["Safety Stock (Channel)"] = (
        z_value * plan_ft["Demand StdDev (Channel)"] * safety_sqrt
    )

    plan_ft["Base Requirement (Channel)"] = (
        plan_ft["Avg Daily Sale (Channel)"] * float(planning_days)
    )
    plan_ft["Required Stock (Channel)"] = (
        plan_ft["Base Requirement (Channel)"] + plan_ft["Safety Stock (Channel)"]
    )

    plan_ft["Recommended Dispatch (Channel)"] = (
        plan_ft["Required Stock (Channel)"] - plan_ft["Current Stock (Channel)"]
    ).clip(lower=0).round(0)

    plan_ft["Days of Cover (Channel)"] = (
        plan_ft["Current Stock (Channel)"]
        / plan_ft["Avg Daily Sale (Channel)"].replace(0, 1)
    )

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

    # Separate FBA / FBM
    fba_plan = plan_ft[plan_ft["Fulfillment Type"] == "FBA"].copy()
    fbm_plan = plan_ft[plan_ft["Fulfillment Type"] == "FBM"].copy()

    # ---------- SITE-LEVEL DEMAND & FC-WISE RECOMMENDATION ----------
    # Demand per site (using Ship To State as proxy; if Ship From available, you can switch)
    site_sales = (
        hist_sales_df.groupby(["Sku", "Fulfillment Type", "Ship To State"])["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "History Sales (Site)"})
    )

    # Map Ship To State onto inventory Ship From State if possible (approximate).
    # For now we aggregate sales by (Sku, FT, Ship From State) from MTR if exists.
    if "Ship From State" in hist_sales_df.columns:
        site_sales_from = (
            hist_sales_df.groupby(
                ["Sku", "Fulfillment Type", "Ship From State"]
            )["Quantity"]
            .sum()
            .reset_index()
            .rename(columns={"Quantity": "History Sales (Site)"})
        )
    else:
        site_sales_from = pd.DataFrame(
            columns=["Sku", "Fulfillment Type", "Ship From State", "History Sales (Site)"]
        )

    # Build site plan from inventory sites
    site_plan = stock_by_site.merge(
        site_sales_from,
        on=["Sku", "Fulfillment Type", "Ship From State"],
        how="left",
    )
    site_plan["History Sales (Site)"] = site_plan["History Sales (Site)"].fillna(0.0)
    site_plan["Avg Daily Sale (Site)"] = (
        site_plan["History Sales (Site)"] / float(history_window_days)
    )

    site_plan["Days of Cover (Site)"] = (
        site_plan["Current Stock"] / site_plan["Avg Daily Sale (Site)"].replace(0, 1)
    )

    # Attach channel-level recommended dispatch to every site row
    site_plan = site_plan.merge(
        plan_ft[
            ["Sku", "Fulfillment Type", "Recommended Dispatch (Channel)",
             "History Sales (Channel)"]
        ],
        on=["Sku", "Fulfillment Type"],
        how="left",
    )

    # Compute demand share per site and allocate recommended dispatch
    temp = site_plan.groupby(["Sku", "Fulfillment Type"])["History Sales (Site)"].sum().reset_index()
    temp.rename(columns={"History Sales (Site)": "Total Site History Sales"}, inplace=True)
    site_plan = site_plan.merge(
        temp, on=["Sku", "Fulfillment Type"], how="left"
    )

    # If no site-level history (0), fall back to equal allocation across sites for that SKU+FT
    def compute_share(row):
        if row["Total Site History Sales"] > 0:
            return row["History Sales (Site)"] / row["Total Site History Sales"]
        else:
            # equal split among sites for this SKU+FT
            return np.nan

    site_plan["Demand Share (Site)"] = site_plan.apply(compute_share, axis=1)

    # If NaN (no history), equal split
    count_sites = (
        site_plan.groupby(["Sku", "Fulfillment Type"])["Warehouse Code"]
        .transform("count")
        .replace(0, 1)
    )
    site_plan["Demand Share (Site)"] = site_plan["Demand Share (Site)"].fillna(
        1.0 / count_sites
    )

    site_plan["Recommended Dispatch (Site)"] = (
        site_plan["Recommended Dispatch (Channel)"]
        * site_plan["Demand Share (Site)"]
    ).round(0)

    site_plan = add_serial(site_plan)

    # ---------- GLOBAL SKU SUMMARY ----------
    sku_summary = (
        plan_ft.groupby("Sku")
        .agg(
            Channels=("Fulfillment Type", lambda x: ", ".join(sorted(set(x)))),
            Total_History_Sales=("History Sales (Channel)", "sum"),
            Total_Current_Stock=("Current Stock (Channel)", "sum"),
            Total_Recommended_Dispatch=("Recommended Dispatch (Channel)", "sum"),
        )
        .reset_index()
    )
    sku_summary = sku_summary.merge(sku_stock_all, on="Sku", how="left")
    sku_summary["Current Stock (All)"] = sku_summary["Current Stock (All)"].fillna(0)
    sku_summary = add_serial(sku_summary)

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

    # ---------- STATE-LEVEL SALES ----------
    state_to = (
        sales.groupby("Ship To State")["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Qty"})
        .sort_values("Total Qty", ascending=False)
    )
    state_to = add_serial(state_to)

    state_from = (
        sales.groupby("Ship From State")["Quantity"]
        .sum()
        .reset_index()
        .rename(columns={"Quantity": "Total Qty"})
        .sort_values("Total Qty", ascending=False)
    )
    state_from = add_serial(state_from)

    # ================== UI TABS ==================
    tab_fba, tab_fbm, tab_sites, tab_sku, tab_risk, tab_state = st.tabs(
        [
            "FBA Planning",
            "FBM Planning",
            "Warehouse / FC Plan",
            "SKU Summary",
            "Risk Alerts",
            "State Sales",
        ]
    )

    with tab_fba:
        st.subheader("📦 FBA Planning (per SKU + Channel)")
        st.dataframe(fba_plan, use_container_width=True)

    with tab_fbm:
        st.subheader("📮 FBM Planning (per SKU + Channel)")
        st.dataframe(fbm_plan, use_container_width=True)

    with tab_sites:
        st.subheader("🏭 Warehouse / FC Wise Plan (FBA & FBM)")
        st.caption(
            "Recommended Dispatch (Site) tells you how many units to send to each FC "
            "for the planning period, allocated by site-level demand share."
        )
        st.dataframe(site_plan, use_container_width=True)

    with tab_sku:
        st.subheader("📊 SKU Level Summary (All Channels)")
        st.dataframe(sku_summary, use_container_width=True)

    with tab_risk:
        st.subheader("🔥 Dead / Zero Velocity")
        st.dataframe(dead_stock, use_container_width=True)
        st.subheader("🟡 Slow Moving (Cover > 90 Days)")
        st.dataframe(slow_moving, use_container_width=True)
        st.subheader("🔵 Excess (Cover > 2 × Planning Days)")
        st.dataframe(excess_stock, use_container_width=True)

    with tab_state:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("🌍 Ship To State – Sales")
            st.dataframe(state_to, use_container_width=True)
        with col2:
            st.subheader("🚚 Ship From State – Sales")
            st.dataframe(state_from, use_container_width=True)

    # ================== EXCEL EXPORT ==================
    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            plan_ft.to_excel(writer, "Planning_All_Channels", index=False)
            fba_plan.to_excel(writer, "Planning_FBA", index=False)
            fbm_plan.to_excel(writer, "Planning_FBM", index=False)
            site_plan.to_excel(writer, "Warehouse_FC_Plan", index=False)
            sku_summary.to_excel(writer, "SKU_Summary", index=False)
            dead_stock.to_excel(writer, "Risk_Dead", index=False)
            slow_moving.to_excel(writer, "Risk_Slow", index=False)
            excess_stock.to_excel(writer, "Risk_Excess", index=False)
            state_to.to_excel(writer, "State_Sales_ShipTo", index=False)
            state_from.to_excel(writer, "State_Sales_ShipFrom", index=False)
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
