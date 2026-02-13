import pandas as pd
import os
from glob import glob
from datetime import datetime
import numpy as np

# ========= SETTINGS =========
INPUT_FOLDER = "input_reports"
OUTPUT_FILE = "FBA_Professional_Client_Report.xlsx"
DAYS_COVER_TARGET = 45
DEAD_STOCK_THRESHOLD_DAYS = 60
BRAND_NAME = "Glooya"
# ============================

def load_data(folder):
    files = glob(os.path.join(folder, "*.csv"))
    all_data = []
    for file in files:
        df = pd.read_csv(file, low_memory=False)
        df["Source_File"] = os.path.basename(file)
        all_data.append(df)
    return pd.concat(all_data, ignore_index=True)

def preprocess(df):
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)
    df["Shipment Date"] = pd.to_datetime(df["Shipment Date"], errors="coerce")
    return df

def sales_summary(df):
    total_units = df["Quantity"].sum()

    state_summary = df.groupby("Ship To State")["Quantity"].sum().reset_index()
    state_summary = state_summary.sort_values(by="Quantity", ascending=False)

    channel_summary = df.groupby("Fulfillment Channel")["Quantity"].sum().reset_index()

    city_summary = df.groupby("Ship To City")["Quantity"].sum().reset_index()
    city_summary = city_summary.sort_values(by="Quantity", ascending=False).head(20)

    return total_units, state_summary, channel_summary, city_summary

def forecast_30_days(df):
    last_30_days = df[df["Shipment Date"] >= df["Shipment Date"].max() - pd.Timedelta(days=30)]
    avg_daily_sales = last_30_days["Quantity"].sum() / 30
    forecast = avg_daily_sales * 30
    return round(avg_daily_sales,2), round(forecast,2)

def inventory_recommendation(avg_daily_sales):
    recommended_stock = avg_daily_sales * DAYS_COVER_TARGET
    return round(recommended_stock)

def dead_stock_analysis(df):
    today = df["Shipment Date"].max()
    last_sale = df.groupby("Sku")["Shipment Date"].max().reset_index()
    last_sale["Days Since Last Sale"] = (today - last_sale["Shipment Date"]).dt.days
    dead_stock = last_sale[last_sale["Days Since Last Sale"] > DEAD_STOCK_THRESHOLD_DAYS]
    return dead_stock

def export_excel(total_units, state_summary, channel_summary, city_summary,
                 avg_daily_sales, forecast, recommended_stock, dead_stock):

    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        workbook = writer.book

        # Executive Summary
        summary_df = pd.DataFrame({
            "Metric": [
                "Total Units Sold",
                "Average Daily Sales (Last 30 Days)",
                "30 Day Forecast",
                f"Recommended Stock ({DAYS_COVER_TARGET} Days Cover)"
            ],
            "Value": [
                total_units,
                avg_daily_sales,
                forecast,
                recommended_stock
            ]
        })
        summary_df.to_excel(writer, sheet_name="Executive Summary", index=False)

        state_summary.to_excel(writer, sheet_name="State Wise Sales", index=False)
        channel_summary.to_excel(writer, sheet_name="FBA vs MFN", index=False)
        city_summary.to_excel(writer, sheet_name="Top Cities", index=False)
        dead_stock.to_excel(writer, sheet_name="Dead Stock Alert", index=False)

        # Chart
        worksheet = writer.sheets["State Wise Sales"]
        chart = workbook.add_chart({'type': 'column'})

        chart.add_series({
            'categories': f"='State Wise Sales'!$A$2:$A${len(state_summary)+1}",
            'values':     f"='State Wise Sales'!$B$2:$B${len(state_summary)+1}",
        })

        chart.set_title({'name': 'State Wise Sales'})
        worksheet.insert_chart('D2', chart)

        # Branding Sheet
        branding_sheet = workbook.add_worksheet("Branding")
        branding_sheet.write("A1", f"{BRAND_NAME} - FBA Automation Report")
        branding_sheet.write("A3", f"Generated on: {datetime.now().strftime('%Y-%m-%d')}")

    print("Professional Report Generated Successfully!")

# ===== MAIN =====
if __name__ == "__main__":
    print("Loading Data...")
    data = load_data(INPUT_FOLDER)

    print("Processing...")
    data = preprocess(data)

    total_units, state_summary, channel_summary, city_summary = sales_summary(data)
    avg_daily_sales, forecast = forecast_30_days(data)
    recommended_stock = inventory_recommendation(avg_daily_sales)
    dead_stock = dead_stock_analysis(data)

    print("Exporting Excel...")
    export_excel(total_units, state_summary, channel_summary, city_summary,
                 avg_daily_sales, forecast, recommended_stock, dead_stock)
