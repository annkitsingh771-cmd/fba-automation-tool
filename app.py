import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import timedelta

st.set_page_config(page_title="FBA Master Planner", layout="wide")
st.title("📦 FBA Advanced Inventory & Product Intelligence Tool")

st.markdown("Upload MTR ZIP/CSV + Inventory ZIP/CSV")

# ================= FILE UPLOAD =================
mtr_files = st.file_uploader(
    "Upload MTR Files (ZIP or CSV)",
    type=["csv", "zip"],
    accept_multiple_files=True
)

inventory_file = st.file_uploader(
    "Upload Inventory Report (ZIP or CSV)",
    type=["csv", "zip"]
)

# ================= HELPER FUNCTION =================
def read_file(uploaded_file):
    if uploaded_file.name.endswith(".zip"):
        with zipfile.ZipFile(uploaded_file) as z:
            file_list = z.namelist()
            for file in file_list:
                if file.endswith(".csv"):
                    return pd.read_csv(z.open(file), low_memory=False)
    else:
        return pd.read_csv(uploaded_file, low_memory=False)

# ================= PROCESS =================
if mtr_files and inventory_file:

    # ===== LOAD SALES =====
    sales_list = []
    for file in mtr_files:
        df = read_file(file)
        sales_list.append(df)

    sales_data = pd.concat(sales_list, ignore_index=True)

    sales_data["Quantity"] = pd.to_numeric(sales_data["Quantity"], errors="coerce").fillna(0)
    sales_data["Shipment Date"] = pd.to_datetime(sales_data["Shipment Date"], errors="coerce")

    # ===== LOAD INVENTORY =====
    inventory_data = read_file(inventory_file)
    inventory_data.columns = inventory_data.columns.str.strip()

    # Improved stock detection
    possible_stock_columns = [
        "Available",
        "Available Quantity",
        "afn-available-quantity",
        "quantity",
        "fulfillable-quantity"
    ]

    stock_column = None
    for col in inventory_data.columns:
        for possible in possible_stock_columns:
            if possible.lower() in col.lower():
                stock_column = col
                break

    if stock_column is None:
        st.error("❌ Stock column not detected. Please upload standard Amazon FBA Inventory report.")
        st.write("Detected columns:")
        st.write(inventory_data.columns)
        st.stop()

    # Detect SKU column
    sku_column = None
    for col in inventory_data.columns:
        if "sku" in col.lower():
            sku_column = col
            break

    if sku_column is None:
        st.error("❌ SKU column not detected in inventory file.")
        st.stop()

    stock_summary = inventory_data[[sku_column, stock_column]].copy()
    stock_summary.columns = ["Sku", "Current FBA Stock"]

    # ===== SALES SUMMARY =====
    total_sales = sales_data.groupby("Sku")["Quantity"].sum().reset_index()

    fba_sales = sales_data[sales_data["Fulfillment Channel"] == "AFN"]
    mfn_sales = sales_data[sales_data["Fulfillment Channel"] == "MFN"]

    fba_summary = fba_sales.groupby("Sku")["Quantity"].sum().reset_index()
    mfn_summary = mfn_sales.groupby("Sku")["Quantity"].sum().reset_index()

    # ===== FORECAST =====
    max_date = sales_data["Shipment Date"].max()
    last_30 = sales_data[sales_data["Shipment Date"] >= max_date - timedelta(days=30)]

    avg_daily = last_30.groupby("Sku")["Quantity"].sum().reset_index()
    avg_daily["Avg Daily Sale"] = avg_daily["Quantity"] / 30

    # ===== PRODUCT INFO =====
    product_info = sales_data[["Sku", "Asin", "Item Description"]].drop_duplicates()

    # ===== MERGE PRODUCT REPORT =====
    product_report = total_sales.merge(stock_summary, on="Sku", how="left")
    product_report = product_report.merge(fba_summary, on="Sku", how="left")
    product_report = product_report.merge(mfn_summary, on="Sku", how="left")
    product_report = product_report.merge(avg_daily[["Sku", "Avg Daily Sale"]], on="Sku", how="left")
    product_report = product_report.merge(product_info, on="Sku", how="left")

    product_report = product_report.fillna(0)

    product_report["30 Day Forecast"] = product_report["Avg Daily Sale"] * 30
    product_report["Days of Cover"] = product_report["Current FBA Stock"] / product_report["Avg Daily Sale"].replace(0,1)

    product_report["Stock Status"] = np.where(
        product_report["Days of Cover"] < 30,
        "⚠ Restock Required",
        "✅ Healthy"
    )

    # ===== STATE WISE =====
    state_sales = (
        sales_data.groupby(["Sku", "Ship To State"])["Quantity"]
        .sum()
        .reset_index()
    )

    # ================= DASHBOARD =================
    st.subheader("📦 Product Performance Report")
    st.dataframe(product_report, use_container_width=True)

    st.subheader("🌍 State Wise Sales")
    st.dataframe(state_sales, use_container_width=True)

    # ================= EXCEL EXPORT =================
    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            product_report.to_excel(writer, sheet_name="Product Report", index=False)
            state_sales.to_excel(writer, sheet_name="State Wise Sales", index=False)
        output.seek(0)
        return output

    excel_data = generate_excel()

    st.download_button(
        label="Download Full Business Report",
        data=excel_data,
        file_name="FBA_Master_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload both MTR and Inventory file (ZIP or CSV).")
