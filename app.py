import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import timedelta

st.set_page_config(page_title="FBA Automation Tool", layout="wide")

st.title("📦 FBA Automation & Inventory Planning Tool")
st.markdown("Upload Amazon MTR CSV files to generate a professional report.")

uploaded_files = st.file_uploader(
    "Upload MTR CSV Files",
    type=["csv"],
    accept_multiple_files=True
)

if uploaded_files:

    # ===== LOAD DATA =====
    all_data = []
    for file in uploaded_files:
        df = pd.read_csv(file, low_memory=False)
        all_data.append(df)

    data = pd.concat(all_data, ignore_index=True)

    # ===== DATA CLEANING =====
    data["Quantity"] = pd.to_numeric(data["Quantity"], errors="coerce").fillna(0)
    data["Shipment Date"] = pd.to_datetime(data["Shipment Date"], errors="coerce")

    st.success("Files uploaded successfully!")

    # ===== EXECUTIVE SUMMARY =====
    total_units = int(data["Quantity"].sum())

    max_date = data["Shipment Date"].max()
    last_30 = data[data["Shipment Date"] >= max_date - timedelta(days=30)]

    avg_daily_sales = last_30["Quantity"].sum() / 30
    forecast_30 = avg_daily_sales * 30
    recommended_stock = avg_daily_sales * 45

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total Units Sold", total_units)
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

    sku_last_sale = (
        data.groupby("Sku")["Shipment Date"]
        .max()
        .reset_index()
    )

    sku_last_sale["Days Since Last Sale"] = (
        max_date - sku_last_sale["Shipment Date"]
    ).dt.days

    dead_stock = sku_last_sale[sku_last_sale["Days Since Last Sale"] > 60]

    if len(dead_stock) > 0:
        st.warning("⚠ Dead stock detected (No sale in 60+ days)")
        st.dataframe(dead_stock, use_container_width=True)
    else:
        st.success("✅ No dead stock detected")

    # ===== EXCEL REPORT =====
    st.subheader("📥 Download Excel Report")

    def generate_excel():
        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            # Executive Summary Sheet
            summary_df = pd.DataFrame({
                "Metric": [
                    "Total Units Sold",
                    "Avg Daily Sales (30 Days)",
                    "30 Day Forecast",
                    "Recommended Stock (45 Days Cover)"
                ],
                "Value": [
                    total_units,
                    round(avg_daily_sales, 2),
                    round(forecast_30, 2),
                    round(recommended_stock, 2)
                ]
            })

            summary_df.to_excel(writer, sheet_name="Executive Summary", index=False)
            state_summary.to_excel(writer, sheet_name="State Wise Sales", index=False)
            channel_summary.to_excel(writer, sheet_name="FBA vs MFN", index=False)
            city_summary.to_excel(writer, sheet_name="Top Cities", index=False)

            if len(dead_stock) > 0:
                dead_stock.to_excel(writer, sheet_name="Dead Stock", index=False)

        output.seek(0)
        return output

    excel_data = generate_excel()

    st.download_button(
        label="Download Full Professional Report",
        data=excel_data,
        file_name="FBA_Professional_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload at least one MTR CSV file to begin analysis.")
