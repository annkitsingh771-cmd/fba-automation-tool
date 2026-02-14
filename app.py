import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import timedelta
import math

st.set_page_config(page_title="FBA Supply Intelligence", layout="wide")
st.title("📦 FBA Supply Intelligence System")

# ---------------- USER INPUT ----------------
planning_days = st.number_input("Stock Planning Period (Days)",1,180,30)
service_level_z = st.selectbox("Service Level (Safety)",{
    "90%":1.28,
    "95%":1.65,
    "98%":2.05
})
z_value = service_level_z

# ---------------- FILE UPLOAD ----------------
mtr_files = st.file_uploader("Upload MTR (ZIP/CSV)", type=["csv","zip"], accept_multiple_files=True)
inventory_file = st.file_uploader("Upload Inventory Ledger (ZIP/CSV)", type=["csv","zip"])

def read_file(file):
    if file.name.endswith(".zip"):
        with zipfile.ZipFile(file) as z:
            for name in z.namelist():
                if name.endswith(".csv"):
                    return pd.read_csv(z.open(name), low_memory=False)
    else:
        return pd.read_csv(file, low_memory=False)

# ---------------- MAIN ----------------
if mtr_files and inventory_file:

    sales = pd.concat([read_file(f) for f in mtr_files], ignore_index=True)

    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
    sales["Ship To State"] = sales["Ship To State"].str.upper()

    # ---------------- INVENTORY ----------------
    inv = read_file(inventory_file)
    inv.columns = inv.columns.str.strip()
    inv["Date"] = pd.to_datetime(inv["Date"], errors="coerce")

    latest_stock = inv.sort_values("Date").groupby("MSKU").tail(1)

    stock = latest_stock[["MSKU","Ending Warehouse Balance"]]
    stock.columns = ["Sku","Current Stock"]
    stock["Current Stock"] = pd.to_numeric(stock["Current Stock"], errors="coerce").fillna(0)

    # ---------------- SALES PERIOD ----------------
    sales_period = (sales["Shipment Date"].max() - sales["Shipment Date"].min()).days + 1

    total_sales = sales.groupby("Sku")["Quantity"].sum().reset_index()
    report = total_sales.merge(stock,on="Sku",how="left")
    report.fillna(0,inplace=True)

    report.rename(columns={"Quantity":"Total Sales"}, inplace=True)

    report["Sales Period (Days)"] = sales_period
    report["Avg Daily Sale"] = report["Total Sales"] / sales_period

    # ---------------- VOLATILITY SAFETY STOCK ----------------
    daily_sales = sales.groupby(["Sku","Shipment Date"])["Quantity"].sum().reset_index()

    std_dev = daily_sales.groupby("Sku")["Quantity"].std().reset_index()
    std_dev.rename(columns={"Quantity":"Demand StdDev"}, inplace=True)

    report = report.merge(std_dev,on="Sku",how="left")
    report["Demand StdDev"] = report["Demand StdDev"].fillna(0)

    report["Safety Stock"] = z_value * report["Demand StdDev"] * np.sqrt(planning_days)

    report["Required Stock"] = (
        (report["Avg Daily Sale"] * planning_days)
        + report["Safety Stock"]
    )

    report["Recommended Dispatch Qty"] = (
        report["Required Stock"] - report["Current Stock"]
    ).clip(lower=0)

    report["Days of Cover"] = report["Current Stock"] / report["Avg Daily Sale"].replace(0,1)

    # ---------------- DEAD STOCK ----------------
    last_sale = sales.groupby("Sku")["Shipment Date"].max().reset_index()
    last_sale["Days Since Last Sale"] = (
        sales["Shipment Date"].max() - last_sale["Shipment Date"]
    ).dt.days

    dead_stock = last_sale.merge(stock,on="Sku",how="left")
    dead_stock = dead_stock[
        (dead_stock["Days Since Last Sale"] > 60) &
        (dead_stock["Current Stock"] > 0)
    ]

    # ---------------- SLOW MOVING ----------------
    slow_moving = report[
        (report["Days of Cover"] > 90) &
        (report["Avg Daily Sale"] < report["Avg Daily Sale"].median())
    ]

    # ---------------- EXCESS STOCK ----------------
    excess_stock = report[
        report["Days of Cover"] > planning_days * 2
    ]

    # ---------------- STATE HEATMAP ----------------
    state_pivot = sales.pivot_table(
        values="Quantity",
        index="Ship To State",
        aggfunc="sum"
    ).sort_values("Quantity", ascending=False)

    # ---------------- DISPLAY ----------------
    st.subheader("📊 Main Planning Report")
    st.dataframe(report, use_container_width=True)

    st.subheader("🔥 Dead Stock Alert (60+ days no sale)")
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

    # ---------------- EXCEL EXPORT ----------------
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
