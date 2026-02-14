import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import math

st.set_page_config(page_title="FBA Supply Intelligence", layout="wide")
st.title("📦 FBA Supply Intelligence System")

# ---------------- USER INPUT ----------------
planning_days = st.number_input("Stock Planning Period (Days)", 1, 180, 30)

service_level_option = st.selectbox(
    "Service Level (Safety)",
    ["90%", "95%", "98%"]
)

z_map = {"90%": 1.28, "95%": 1.65, "98%": 2.05}
z_value = z_map[service_level_option]

# ---------------- FILE UPLOAD ----------------
mtr_files = st.file_uploader("Upload MTR (ZIP/CSV)", type=["csv","zip"], accept_multiple_files=True)
inventory_file = st.file_uploader("Upload Inventory Ledger (ZIP/CSV)", type=["csv","zip"])

def read_file(file):
    try:
        if file.name.endswith(".zip"):
            with zipfile.ZipFile(file) as z:
                for name in z.namelist():
                    if name.endswith(".csv"):
                        return pd.read_csv(z.open(name), low_memory=False)
        else:
            return pd.read_csv(file, low_memory=False)
    except:
        return pd.DataFrame()

# ---------------- MAIN ----------------
if mtr_files and inventory_file:

    # ===== LOAD SALES =====
    sales_list = []
    for f in mtr_files:
        df = read_file(f)
        if not df.empty:
            sales_list.append(df)

    if not sales_list:
        st.error("Sales file not readable.")
        st.stop()

    sales = pd.concat(sales_list, ignore_index=True)

    required_cols = ["Sku","Quantity","Shipment Date","Ship To State"]
    for col in required_cols:
        if col not in sales.columns:
            st.error(f"Missing column in Sales file: {col}")
            st.stop()

    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
    sales["Ship To State"] = sales["Ship To State"].astype(str).str.upper()

    sales = sales.dropna(subset=["Shipment Date"])

    # ===== LOAD INVENTORY =====
    inv = read_file(inventory_file)

    if inv.empty:
        st.error("Inventory file not readable.")
        st.stop()

    inv.columns = inv.columns.str.strip()

    if "MSKU" not in inv.columns or "Ending Warehouse Balance" not in inv.columns:
        st.error("Inventory must contain MSKU and Ending Warehouse Balance.")
        st.stop()

    inv["Ending Warehouse Balance"] = pd.to_numeric(
        inv["Ending Warehouse Balance"], errors="coerce"
    ).fillna(0)

    latest_stock = inv.groupby("MSKU")["Ending Warehouse Balance"].last().reset_index()
    latest_stock.columns = ["Sku","Current Stock"]

    # ===== SALES PERIOD =====
    sales_period = (sales["Shipment Date"].max() - sales["Shipment Date"].min()).days + 1
    if sales_period <= 0:
        sales_period = 1

    # ===== PRODUCT LEVEL =====
    total_sales = sales.groupby("Sku")["Quantity"].sum().reset_index()
    report = total_sales.merge(latest_stock, on="Sku", how="left")
    report["Current Stock"] = report["Current Stock"].fillna(0)

    report.rename(columns={"Quantity":"Total Sales"}, inplace=True)

    report["Sales Period (Days)"] = sales_period
    report["Avg Daily Sale"] = report["Total Sales"] / float(sales_period)

    # ===== DEMAND VOLATILITY =====
    daily_sales = sales.groupby(["Sku","Shipment Date"])["Quantity"].sum().reset_index()
    std_dev = daily_sales.groupby("Sku")["Quantity"].std().reset_index()
    std_dev.rename(columns={"Quantity":"Demand StdDev"}, inplace=True)

    report = report.merge(std_dev, on="Sku", how="left")
    report["Demand StdDev"] = pd.to_numeric(
        report["Demand StdDev"], errors="coerce"
    ).fillna(0)

    # Safety Stock Calculation (Safe)
    report["Safety Stock"] = (
        float(z_value)
        * report["Demand StdDev"]
        * math.sqrt(float(planning_days))
    )

    report["Required Stock"] = (
        report["Avg Daily Sale"] * float(planning_days)
    ) + report["Safety Stock"]

    report["Recommended Dispatch Qty"] = (
        report["Required Stock"] - report["Current Stock"]
    ).clip(lower=0)

    report["Days of Cover"] = report["Current Stock"] / report["Avg Daily Sale"].replace(0,1)

    report.insert(0, "S.No", range(1, len(report)+1))

    # ===== DEAD STOCK =====
    last_sale = sales.groupby("Sku")["Shipment Date"].max().reset_index()
    last_sale["Days Since Last Sale"] = (
        sales["Shipment Date"].max() - last_sale["Shipment Date"]
    ).dt.days

    dead_stock = last_sale.merge(latest_stock, on="Sku", how="left")
    dead_stock = dead_stock[
        (dead_stock["Days Since Last Sale"] > 60) &
        (dead_stock["Current Stock"] > 0)
    ]

    dead_stock.insert(0, "S.No", range(1, len(dead_stock)+1))

    # ===== SLOW MOVING =====
    slow_moving = report[
        (report["Days of Cover"] > 90)
    ].copy()

    slow_moving.insert(0, "S.No", range(1, len(slow_moving)+1))

    # ===== EXCESS STOCK =====
    excess_stock = report[
        report["Days of Cover"] > (planning_days * 2)
    ].copy()

    excess_stock.insert(0, "S.No", range(1, len(excess_stock)+1))

    # ===== STATE HEATMAP =====
    state_pivot = sales.pivot_table(
        values="Quantity",
        index="Ship To State",
        aggfunc="sum"
    ).sort_values("Quantity", ascending=False)

    # ===== DISPLAY =====
    st.subheader("📊 Main Planning Report")
    st.dataframe(report, use_container_width=True)

    st.subheader("🔥 Dead Stock Alert")
    st.dataframe(dead_stock, use_container_width=True)

    st.subheader("🟡 Slow Moving SKUs")
    st.dataframe(slow_moving, use_container_width=True)

    st.subheader("🔵 Excess Stock Warning")
    st.dataframe(excess_stock, use_container_width=True)

    st.subheader("🌍 State Sales Heatmap")
    st.dataframe(
        state_pivot.style.background_gradient(cmap="Reds"),
        use_container_width=True
    )

    # ===== EXCEL EXPORT =====
    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            report.to_excel(writer,"Main Planning",index=False)
            dead_stock.to_excel(writer,"Dead Stock",index=False)
            slow_moving.to_excel(writer,"Slow Moving",index=False)
            excess_stock.to_excel(writer,"Excess Stock",index=False)
            state_pivot.to_excel(writer,"State Heatmap")
        output.seek(0)
        return output

    st.download_button(
        "📥 Download Full Intelligence Report",
        generate_excel(),
        "FBA_Supply_Intelligence.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload Sales and Inventory file to start.")
