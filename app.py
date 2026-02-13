import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import timedelta

st.set_page_config(page_title="FBA Intelligence Pro", layout="wide")

st.title("🚀 FBA Advanced Intelligence Dashboard")
st.markdown("Upload MTR (Sales) + Inventory Ledger to unlock full analytics")

# ---------------- FILE UPLOAD ----------------

mtr_files = st.file_uploader(
    "Upload MTR Files (ZIP/CSV)",
    type=["csv","zip"],
    accept_multiple_files=True
)

inventory_file = st.file_uploader(
    "Upload Inventory Ledger (ZIP/CSV)",
    type=["csv","zip"]
)

# ---------------- FILE READER ----------------

def read_file(uploaded):
    if uploaded.name.endswith(".zip"):
        with zipfile.ZipFile(uploaded) as z:
            for name in z.namelist():
                if name.endswith(".csv"):
                    return pd.read_csv(z.open(name), low_memory=False)
    else:
        return pd.read_csv(uploaded, low_memory=False)

# ---------------- MAIN LOGIC ----------------

if mtr_files and inventory_file:

    # SALES LOAD
    sales_list = [read_file(f) for f in mtr_files]
    sales = pd.concat(sales_list, ignore_index=True)

    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")

    # INVENTORY LOAD
    inventory = read_file(inventory_file)
    inventory.columns = inventory.columns.str.strip()

    if "Ending Warehouse Balance" not in inventory.columns:
        st.error("Upload Inventory Ledger report (Ending Warehouse Balance required)")
        st.stop()

    if "MSKU" not in inventory.columns:
        st.error("MSKU column missing in inventory file")
        st.stop()

    inventory["Date"] = pd.to_datetime(inventory["Date"], errors="coerce")

    latest_stock = (
        inventory.sort_values("Date")
        .groupby("MSKU")
        .tail(1)
    )

    stock = latest_stock[["MSKU","Ending Warehouse Balance"]]
    stock.columns = ["Sku","Current Stock"]
    stock["Current Stock"] = pd.to_numeric(stock["Current Stock"], errors="coerce").fillna(0)

    # SALES SPLIT
    total_sales = sales.groupby("Sku")["Quantity"].sum().reset_index()
    fba_sales = sales[sales["Fulfillment Channel"]=="AFN"].groupby("Sku")["Quantity"].sum().reset_index()
    mfn_sales = sales[sales["Fulfillment Channel"]=="MFN"].groupby("Sku")["Quantity"].sum().reset_index()

    # FORECAST
    max_date = sales["Shipment Date"].max()
    last_30 = sales[sales["Shipment Date"] >= max_date - timedelta(days=30)]
    avg_daily = last_30.groupby("Sku")["Quantity"].sum().reset_index()
    avg_daily["Avg Daily"] = avg_daily["Quantity"] / 30

    # MERGE
    report = total_sales.merge(stock,on="Sku",how="left")
    report = report.merge(fba_sales,on="Sku",how="left",suffixes=("","_FBA"))
    report = report.merge(mfn_sales,on="Sku",how="left",suffixes=("","_MFN"))
    report = report.merge(avg_daily[["Sku","Avg Daily"]],on="Sku",how="left")

    report.fillna(0,inplace=True)

    report.rename(columns={
        "Quantity":"Total Sales",
        "Quantity_FBA":"FBA Sales",
        "Quantity_MFN":"MFN Sales"
    },inplace=True)

    report["30 Day Forecast"] = report["Avg Daily"]*30
    report["Days of Cover"] = report["Current Stock"] / report["Avg Daily"].replace(0,1)

    report["Reorder Qty (45D Cover)"] = (report["Avg Daily"]*45) - report["Current Stock"]

    report["Stock Status"] = np.where(
        report["Days of Cover"]<15,"🔴 Critical",
        np.where(report["Days of Cover"]<30,"🟠 Warning","🟢 Healthy")
    )

    # ---------------- KPI DASHBOARD ----------------

    st.subheader("📊 Executive KPI")

    col1,col2,col3,col4 = st.columns(4)

    col1.metric("Total Units Sold",int(report["Total Sales"].sum()))
    col2.metric("Total FBA Sales",int(report["FBA Sales"].sum()))
    col3.metric("Total MFN Sales",int(report["MFN Sales"].sum()))
    col4.metric("Current Total Stock",int(report["Current Stock"].sum()))

    st.divider()

    # ---------------- CHARTS ----------------

    st.subheader("📈 Sales Trend")

    daily_sales = sales.groupby("Shipment Date")["Quantity"].sum()
    st.line_chart(daily_sales)

    st.subheader("🏆 Top Performing SKUs")
    top_sku = report.sort_values("Total Sales",ascending=False).head(10)
    st.bar_chart(top_sku.set_index("Sku")["Total Sales"])

    st.subheader("🚚 FBA vs MFN Comparison")
    channel_summary = sales.groupby("Fulfillment Channel")["Quantity"].sum()
    st.bar_chart(channel_summary)

    st.subheader("🌍 State Wise Sales")
    state_sales = sales.groupby("Ship To State")["Quantity"].sum().sort_values(ascending=False)
    st.bar_chart(state_sales)

    st.subheader("📦 Inventory Planning Table")
    st.dataframe(report, use_container_width=True)

    # ---------------- EXCEL EXPORT ----------------

    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            report.to_excel(writer, sheet_name="Product Intelligence", index=False)
            state_sales.to_excel(writer, sheet_name="State Sales")
            daily_sales.to_excel(writer, sheet_name="Sales Trend")
        output.seek(0)
        return output

    excel_data = generate_excel()

    st.download_button(
        "📥 Download Full Intelligence Report",
        excel_data,
        "FBA_Advanced_Intelligence.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload Sales and Inventory file to begin.")
