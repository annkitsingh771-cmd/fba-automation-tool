import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import timedelta

st.set_page_config(page_title="FBA Smart Supply Planner", layout="wide")
st.title("📦 FBA Smart Supply Planner")

st.markdown("Upload Sales (MTR) + Inventory Ledger")

# ================= CLUSTER MAP =================
cluster_map = {
    "North Cluster": ["UTTAR PRADESH","DELHI","HARYANA","PUNJAB","RAJASTHAN"],
    "West Cluster": ["MAHARASHTRA","GUJARAT"],
    "South Cluster": ["TAMIL NADU","KARNATAKA","KERALA","TELANGANA","ANDHRA PRADESH"],
    "East Cluster": ["WEST BENGAL","ODISHA","ASSAM","BIHAR"],
    "Central Cluster": ["MADHYA PRADESH","CHHATTISGARH"]
}

fc_map = {
    "North Cluster": "DEL FC",
    "West Cluster": "BOM FC",
    "South Cluster": "BLR FC",
    "East Cluster": "CCU FC",
    "Central Cluster": "HYD FC",
    "Other Cluster": "DEL FC"
}

# ================= FILE UPLOAD =================
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

# ================= MAIN =================
if mtr_files and inventory_file:

    # ---------- LOAD SALES ----------
    sales = pd.concat([read_file(f) for f in mtr_files], ignore_index=True)
    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
    sales["Ship To State"] = sales["Ship To State"].str.upper()

    # ---------- LOAD INVENTORY ----------
    inv = read_file(inventory_file)
    inv.columns = inv.columns.str.strip()

    if "Ending Warehouse Balance" not in inv.columns:
        st.error("Inventory Ledger must contain 'Ending Warehouse Balance'")
        st.stop()

    inv["Date"] = pd.to_datetime(inv["Date"], errors="coerce")
    latest_stock = inv.sort_values("Date").groupby("MSKU").tail(1)

    stock = latest_stock[["MSKU","Ending Warehouse Balance"]]
    stock.columns = ["Sku","Current Stock"]
    stock["Current Stock"] = pd.to_numeric(stock["Current Stock"], errors="coerce").fillna(0)

    # ---------- PRODUCT LEVEL ----------
    total_sales = sales.groupby("Sku")["Quantity"].sum().reset_index()

    fba_sales = sales[sales["Fulfillment Channel"]=="AFN"].groupby("Sku")["Quantity"].sum().reset_index()
    mfn_sales = sales[sales["Fulfillment Channel"]=="MFN"].groupby("Sku")["Quantity"].sum().reset_index()

    max_date = sales["Shipment Date"].max()
    last_30 = sales[sales["Shipment Date"] >= max_date - timedelta(days=30)]
    avg_daily = last_30.groupby("Sku")["Quantity"].sum().reset_index()
    avg_daily["Avg Daily Sale"] = avg_daily["Quantity"] / 30

    report = total_sales.merge(stock,on="Sku",how="left")
    report = report.merge(fba_sales,on="Sku",how="left",suffixes=("","_FBA"))
    report = report.merge(mfn_sales,on="Sku",how="left",suffixes=("","_MFN"))
    report = report.merge(avg_daily[["Sku","Avg Daily Sale"]],on="Sku",how="left")

    report.fillna(0,inplace=True)

    report.rename(columns={
        "Quantity":"Total Sales",
        "Quantity_FBA":"FBA Sales",
        "Quantity_MFN":"MFN Sales"
    },inplace=True)

    report["30D Forecast"] = report["Avg Daily Sale"]*30
    report["Days of Cover"] = report["Current Stock"] / report["Avg Daily Sale"].replace(0,1)
    report["Recommended Reorder Qty (45D Cover)"] = (
        (report["Avg Daily Sale"]*45) - report["Current Stock"]
    ).clip(lower=0)

    report["Stock Status"] = np.where(
        report["Days of Cover"]<15,"Critical",
        np.where(report["Days of Cover"]<30,"Warning","Healthy")
    )

    # ---------- CLUSTER ----------
    def assign_cluster(state):
        for cluster, states in cluster_map.items():
            if state in states:
                return cluster
        return "Other Cluster"

    sales["Cluster Name"] = sales["Ship To State"].apply(assign_cluster)

    cluster_sales = sales.groupby(["Sku","Cluster Name"])["Quantity"].sum().reset_index()
    cluster_total = cluster_sales.groupby("Sku")["Quantity"].sum().reset_index()

    cluster_sales = cluster_sales.merge(cluster_total,on="Sku",suffixes=("","_Total"))
    cluster_sales["Cluster Contribution %"] = (
        cluster_sales["Quantity"] / cluster_sales["Quantity_Total"]
    )

    cluster_sales = cluster_sales.merge(
        avg_daily[["Sku","Avg Daily Sale"]],
        on="Sku",how="left"
    )

    cluster_sales["Cluster 30D Forecast"] = (
        cluster_sales["Avg Daily Sale"]*30*
        cluster_sales["Cluster Contribution %"]
    )

    cluster_sales["Suggested Dispatch Qty"] = (
        cluster_sales["Cluster 30D Forecast"]*1.5
    )

    cluster_sales["Mapped FC"] = cluster_sales["Cluster Name"].map(fc_map)

    # ---------- FC PLAN ----------
    fc_plan = cluster_sales.groupby("Mapped FC")["Suggested Dispatch Qty"].sum().reset_index()

    # ---------- DASHBOARD ----------
    st.subheader("📊 Product Planning")
    st.dataframe(report, use_container_width=True)

    st.subheader("📍 Cluster Allocation Plan")
    st.dataframe(cluster_sales, use_container_width=True)

    st.subheader("🏭 FC Shipment Plan")
    st.dataframe(fc_plan, use_container_width=True)

    st.subheader("📈 Sales Trend")
    trend = sales.groupby("Shipment Date")["Quantity"].sum()
    st.line_chart(trend)

    st.subheader("🌍 Cluster Distribution")
    cluster_chart = sales.groupby("Cluster Name")["Quantity"].sum()
    st.bar_chart(cluster_chart)

    # ---------- EXCEL EXPORT ----------
    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            report.to_excel(writer,"Product Planning",index=False)
            cluster_sales.to_excel(writer,"Cluster Allocation",index=False)
            fc_plan.to_excel(writer,"FC Shipment Plan",index=False)
            trend.to_excel(writer,"Sales Trend")
            cluster_chart.to_excel(writer,"Cluster Summary")
        output.seek(0)
        return output

    st.download_button(
        "📥 Download Client Ready Report",
        generate_excel(),
        "FBA_Smart_Supply_Planner.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload Sales and Inventory file to generate planning report.")
