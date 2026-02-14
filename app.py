import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import timedelta

st.set_page_config(page_title="FBA Smart Supply Planner", layout="wide")
st.title("📦 FBA Smart Supply Planner")

# ================= USER CONTROLS =================
planning_days = st.number_input("Planning Days", min_value=1, max_value=180, value=30)
safety_multiplier = st.slider("Safety Buffer Multiplier", 1.0, 2.0, 1.5)

planning_mode = st.selectbox(
    "Planning Based On",
    ["Total Sales", "FBA Sales Only"]
)

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

    sales = pd.concat([read_file(f) for f in mtr_files], ignore_index=True)
    sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
    sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
    sales["Ship To State"] = sales["Ship To State"].str.upper()

    inv = read_file(inventory_file)
    inv.columns = inv.columns.str.strip()
    inv["Date"] = pd.to_datetime(inv["Date"], errors="coerce")
    latest_stock = inv.sort_values("Date").groupby("MSKU").tail(1)

    stock = latest_stock[["MSKU","Ending Warehouse Balance"]]
    stock.columns = ["Sku","Current Stock"]
    stock["Current Stock"] = pd.to_numeric(stock["Current Stock"], errors="coerce").fillna(0)

    # -------- SALES SPLIT --------
    total_sales = sales.groupby("Sku")["Quantity"].sum().reset_index()

    fba_sales = sales[sales["Fulfillment Channel"]=="AFN"].groupby("Sku")["Quantity"].sum().reset_index()
    fba_sales.rename(columns={"Quantity":"FBA Sales"}, inplace=True)

    mfn_sales = sales[sales["Fulfillment Channel"]=="MFN"].groupby("Sku")["Quantity"].sum().reset_index()
    mfn_sales.rename(columns={"Quantity":"MFN Sales"}, inplace=True)

    # -------- MERGE PRODUCT --------
    report = total_sales.merge(stock,on="Sku",how="left")
    report = report.merge(fba_sales,on="Sku",how="left")
    report = report.merge(mfn_sales,on="Sku",how="left")
    report.fillna(0,inplace=True)

    report.rename(columns={"Quantity":"Total Sales"}, inplace=True)

    report["FBA %"] = report["FBA Sales"] / report["Total Sales"].replace(0,1)
    report["MFN %"] = report["MFN Sales"] / report["Total Sales"].replace(0,1)

    # -------- FORECAST BASE --------
    max_date = sales["Shipment Date"].max()
    last_30 = sales[sales["Shipment Date"] >= max_date - timedelta(days=30)]

    if planning_mode == "FBA Sales Only":
        base_data = last_30[last_30["Fulfillment Channel"]=="AFN"]
    else:
        base_data = last_30

    avg_daily = base_data.groupby("Sku")["Quantity"].sum().reset_index()
    avg_daily["Avg Daily"] = avg_daily["Quantity"]/30

    report = report.merge(avg_daily[["Sku","Avg Daily"]], on="Sku", how="left")
    report.fillna(0,inplace=True)

    report["Forecast Qty"] = report["Avg Daily"]*planning_days
    report["Recommended Reorder Qty"] = (
        (report["Avg Daily"]*planning_days*safety_multiplier)
        - report["Current Stock"]
    ).clip(lower=0)

    report["Days of Cover"] = report["Current Stock"] / report["Avg Daily"].replace(0,1)

    # -------- CLUSTER --------
    def assign_cluster(state):
        for cluster, states in cluster_map.items():
            if state in states:
                return cluster
        return "Other Cluster"

    sales["Cluster Name"] = sales["Ship To State"].apply(assign_cluster)

    cluster_sales = sales.groupby(
        ["Sku","Cluster Name","Fulfillment Channel"]
    )["Quantity"].sum().reset_index()

    cluster_sales["Mapped FC"] = cluster_sales["Cluster Name"].map(fc_map)

    # -------- FC PLAN --------
    fc_plan = cluster_sales.groupby("Mapped FC")["Quantity"].sum().reset_index()

    # -------- DISPLAY --------
    st.subheader("📊 Product Planning")
    st.dataframe(report, use_container_width=True)

    st.subheader("📍 Cluster Sales (FBA & MFN Split)")
    st.dataframe(cluster_sales, use_container_width=True)

    st.subheader("🏭 FC Shipment Summary")
    st.dataframe(fc_plan, use_container_width=True)

    # -------- EXCEL --------
    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            report.to_excel(writer,"Product Planning",index=False)
            cluster_sales.to_excel(writer,"Cluster Sales Split",index=False)
            fc_plan.to_excel(writer,"FC Summary",index=False)
        output.seek(0)
        return output

    st.download_button(
        "📥 Download Full Planning Report",
        generate_excel(),
        "FBA_Smart_Supply_Planner.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload Sales and Inventory file to generate planning report.")
