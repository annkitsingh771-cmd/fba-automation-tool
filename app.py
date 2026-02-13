import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import timedelta

st.set_page_config(page_title="FBA Master Planner", layout="wide")
st.title("📦 Advanced FBA Sales & Inventory Intelligence Tool")

st.markdown("Upload MTR (Sales) + FBA Inventory Report")

# ================= FILE UPLOAD =================
mtr_files = st.file_uploader(
    "Upload MTR CSV Files (Sales Data)",
    type=["csv"],
    accept_multiple_files=True
)

inventory_file = st.file_uploader(
    "Upload FBA Inventory Report CSV",
    type=["csv"]
)

if mtr_files and inventory_file:

    # ================= LOAD SALES =================
    sales_list = []
    for file in mtr_files:
        df = pd.read_csv(file, low_memory=False)
        sales_list.append(df)

    sales_data = pd.concat(sales_list, ignore_index=True)

    sales_data["Quantity"] = pd.to_numeric(sales_data["Quantity"], errors="coerce").fillna(0)
    sales_data["Shipment Date"] = pd.to_datetime(sales_data["Shipment Date"], errors="coerce")

    # ================= LOAD INVENTORY =================
    inventory_data = pd.read_csv(inventory_file)
    inventory_data.columns = inventory_data.columns.str.strip()

    stock_column = None
    for col in inventory_data.columns:
        if "available" in col.lower():
            stock_column = col
            break

    if stock_column is None:
        st.error("Stock column not found in inventory file.")
        st.stop()

    inventory_data = inventory_data.rename(columns={"sku": "Sku"})
    stock_summary = inventory_data[["Sku", stock_column]].copy()
    stock_summary.columns = ["Sku", "Current FBA Stock"]

    # ================= SALES SUMMARY =================
    total_sales = sales_data.groupby("Sku")["Quantity"].sum().reset_index()

    fba_sales = sales_data[sales_data["Fulfillment Channel"] == "AFN"]
    mfn_sales = sales_data[sales_data["Fulfillment Channel"] == "MFN"]

    fba_summary = fba_sales.groupby("Sku")["Quantity"].sum().reset_index()
    mfn_summary = mfn_sales.groupby("Sku")["Quantity"].sum().reset_index()

    # ================= PRODUCT INFO =================
    product_info = sales_data[["Sku", "Asin", "Item Description"]].drop_duplicates()

    # ================= FORECAST =================
    max_date = sales_data["Shipment Date"].max()
    last_30 = sales_data[sales_data["Shipment Date"] >= max_date - timedelta(days=30)]

    avg_daily = last_30.groupby("Sku")["Quantity"].sum().reset_index()
    avg_daily["Avg Daily Sale"] = avg_daily["Quantity"] / 30

    # ================= MERGE PRODUCT REPORT =================
    product_report = total_sales.merge(stock_summary, on="Sku", how="left")
    product_report = product_report.merge(fba_summary, on="Sku", how="left")
    product_report = product_report.merge(mfn_summary, on="Sku", how="left")
    product_report = product_report.merge(avg_daily[["Sku", "Avg Daily Sale"]], on="Sku", how="left")
    product_report = product_report.merge(product_info, on="Sku", how="left")

    product_report.columns = [
        "Sku",
        "Total Sales",
        "Current FBA Stock",
        "FBA Sales",
        "MFN Sales",
        "Avg Daily Sale",
        "Asin",
        "Product Name"
    ]

    product_report = product_report.fillna(0)

    product_report["30 Day Forecast"] = product_report["Avg Daily Sale"] * 30
    product_report["Days of Cover"] = product_report["Current FBA Stock"] / product_report["Avg Daily Sale"].replace(0,1)

    product_report["Stock Status"] = np.where(
        product_report["Days of Cover"] < 30,
        "⚠ Restock Required",
        "✅ Healthy"
    )

    # ================= STATE WISE =================
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
        file_name="FBA_Master_Product_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Upload both Sales (MTR) and Inventory file to continue.")
