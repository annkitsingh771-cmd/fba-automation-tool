import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

st.set_page_config(page_title="FBA Automation Tool", layout="wide")

st.title("📦 FBA Automation & Inventory Planning Tool")

st.markdown("Upload Amazon MTR CSV files to generate full professional report.")

uploaded_files = st.file_uploader(
    "Upload MTR CSV Files",
    type=["csv"],
    accept_multiple_files=True
)

if uploaded_files:

    all_data = []

    for file in uploaded_files:
        df = pd.read_csv(file, low_memory=False)
        all_data.append(df)

    data = pd.concat(all_data, ignore_index=True)

    # ===== PREPROCESS =====
    data["Quantity"] = pd.to_numeric(data["Quantity"], errors="coerce").fillna(0)
    data["Shipment Date"] = pd.to_datetime(data["Shipment Date"], errors="coerce")

    st.success("Files uploaded successfully!")

    # ===== EXECUTIVE SUMMARY =====
    total_units = data["Quantity"].sum()

    last_30_days = data[data["Shipment Date"] >= data["Shipment Date"].max() - pd.Timedelta(days=30)]
    avg_daily_sales = last_30_days["Quantity"].sum() / 30
    forecast_30 = avg_daily_sales * 30
    recommended_stock = avg_daily_sales * 45

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total Units Sold", int(total_units))
    col2.metric("Avg Daily Sales (30D)", round(avg_daily_sales, 2))
    col3.metric("30 Day Forecast", round(forecast_30, 2))
    col4.metric("Recommended Stock (45 Days Cover)", round(recommended_stock, 2))

    st.divider()

    # ===== STATE WISE SALES =====
    st.subheader("📊 State Wise Sales")

    state_summary = (
        data.groupby("Ship To State")["Quantity"]
        .sum()
        .reset_index()
        .sort_values(by="Quantity", ascending=False)
    )

    st.dataframe(state_summary, use_container_width=True)

    st.bar_chart(state_summary.set_index("Ship To State"))

    # ===== FBA vs MFN =====
    st.subheader("🚚 FBA vs MFN Analysis")

    channel_summary = (
        data.groupby("Fulfillment Channel")["Quantity"]
        .sum()
        .reset_index()
    )

    st.dataframe(channel_summary, use_container_width=True)
    st.bar_chart(channel_summary.set_index("Fulfillment Channel"))

    # ===== TOP CITIES =====
    st.subheader("🏙 Top Cities")

    city_summary = (
        data.groupby("Ship To City")["Quantity"]
        .sum()
        .reset_index()
        .sort_values(by="Quantity", ascending=False)
        .head(15)
    )

    st.dataframe(city_summary, use_container_width=True)

    # ===== DEAD STOCK =====
    st.subheader("🧠 Dead Stock Alert")

    today = data["Shipment Date"].max()

    sku_last_sale = (
        data.groupby("Sku")["Shipment Date"]
        .max()
        .reset_index()
    )

    sku_last_sale["Days Since Last Sale"] = (
        today - sku_last_sale["Shipment Date"]
    ).dt.days

    dead_stock = sku_last_sale[sku_last_sale["Days Since Last Sale"] > 60]

    if len(dead_stock) > 0:
        st.warning("Dead stock detected (No sale in 60+ days)")
        st.dataframe(dead_stock, use_container_width=True)
    else:
        st.success("No dead stock detected")

    # ===== DOWNLOAD REPORT =====
    st.subheader("📥 Download Excel Report")

    def generate_excel():
        output = pd.ExcelWriter("FBA_Report.xlsx", engine="xlsxwriter")

        state_summary.to_excel(output, sheet_name="State Wise Sales", index=False)
        channel_summary.to_excel(output, sheet_name="FBA vs MFN", index=False)
        city_summary.to_excel(output, sheet_name="Top Cities", inde_
