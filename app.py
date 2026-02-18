import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import math
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import requests
import json
from typing import Dict, List

# ================== SP-API CONFIGURATION ==================
# Add your SP-API credentials here (get from Seller Central > Apps & Services > Develop Apps)
SP_API_CONFIG = {
    'client_id': st.secrets.get('SP_CLIENT_ID', ''),
    'client_secret': st.secrets.get('SP_CLIENT_SECRET', ''),
    'refresh_token': st.secrets.get('SP_REFRESH_TOKEN', ''),
    'lwa_app_id': st.secrets.get('SP_LWA_APP_ID', ''),
    'lwa_client_secret': st.secrets.get('SP_LWA_CLIENT_SECRET', ''),
    'aws_access_key': st.secrets.get('AWS_ACCESS_KEY', ''),
    'aws_secret_key': st.secrets.get('AWS_SECRET_KEY', ''),
    'role_arn': st.secrets.get('ROLE_ARN', ''),
    'marketplace_id': 'A21TJRUUN4KGV'  # India
}

# Indian FC Mapping
FC_MAPPING = {
    'DEL4': 'Delhi North', 'DEL5': 'Delhi North', 'PNQ2': 'Delhi City', 'DEX3': 'New Delhi',
    'BOM1': 'Bhiwandi Mumbai', 'BOM3': 'Nashik West', 'BOM4': 'Vasai Mumbai', 'SAMB': 'Mumbai West',
    'BLR5': 'Bangalore South', 'MAA4': 'Chennai South', 'MAA5': 'Chennai South', 
    'HYD7': 'Hyderabad South', 'HYD8': 'Hyderabad South', 'SCJA': 'Bangalore South',
    'XSAB': 'Bangalore South', 'SBLA': 'Bangalore South', 'SMAB': 'Chennai South',
    'UNKNOWN': 'Unknown FC'
}

def get_fc_name(fc_code):
    """Get full FC name from code."""
    return FC_MAPPING.get(fc_code.upper(), f"FC: {fc_code}")

# ================== SP-API FUNCTIONS ==================
@st.cache_data(ttl=3600)  # Cache for 1 hour
def get_sp_api_inventory():
    """Fetch live inventory from SP-API."""
    try:
        # Using python-amazon-sp-api (pip install python-amazon-sp-api)
        from sp_api.api import Inventory
        from sp_api.base import Marketplaces
        
        inventory = Inventory(
            marketplace=Marketplaces.IN,
            credentials={
                'refresh_token': SP_API_CONFIG['refresh_token'],
                'lwa_app_id': SP_API_CONFIG['lwa_app_id'],
                'lwa_client_secret': SP_API_CONFIG['lwa_client_secret'],
                'aws_access_key': SP_API_CONFIG['aws_access_key'],
                'aws_secret_key': SP_API_CONFIG['aws_secret_key'],
                'role_arn': SP_API_CONFIG['role_arn']
            }
        )
        
        response = inventory.get_inventory_summaries(
            details=True,
            marketplaceIds=[SP_API_CONFIG['marketplace_id']]
        )
        
        if response.payload and 'inventorySummaries' in response.payload:
            df = pd.DataFrame(response.payload['inventorySummaries'])
            df['fc_name'] = df['asin'].apply(lambda x: get_fc_name(x) if x else 'Unknown')
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"SP-API Error: {str(e)}")
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def get_removal_orders():
    """Get FBA removal orders."""
    try:
        from sp_api.api import FBAInbound
        fba = FBAInbound(marketplaces=[Marketplaces.IN])
        response = fba.get_lfn_removal_shipment_status()
        return pd.DataFrame(response.payload.get('shipments', []))
    except:
        return pd.DataFrame()

# ================== UI ==================
st.sidebar.title("⚙️ Controls")
use_api = st.sidebar.checkbox("Connect to SP-API (Live Data)", value=False)
planning_days = st.sidebar.number_input("Planning Days", 7, 180, 60)
safety_level = st.sidebar.selectbox("Safety Level", ["Conservative (95%)", "Balanced (90%)", "Aggressive (98%)"])
z_values = {"Conservative (95%)": 1.65, "Balanced (90%)": 1.28, "Aggressive (98%)": 2.05}
z_value = z_values[safety_level]

st.markdown("---")

# ================== FILE UPLOAD (Fallback) ==================
col1, col2 = st.columns(2)
with col1:
    mtr_files = st.file_uploader("📊 Upload MTR Files", type=["csv","zip"], accept_multiple_files=True)
with col2:
    inv_file = st.file_uploader("📦 Upload Inventory", type=["csv","zip"])

# ================== MAIN DASHBOARD ==================
if use_api and SP_API_CONFIG['client_id']:
    # SP-API Live Data
    with st.spinner("Fetching live data from Amazon Seller Central..."):
        api_inv = get_sp_api_inventory()
        removals = get_removal_orders()
    
    if not api_inv.empty:
        st.success(f"✅ Connected to SP-API! Found {len(api_inv)} SKUs across {api_inv['fnsku'].nunique()} FNSKUs")
        st.metric("Total FBA Inventory", f"{api_inv['totalQuantity'].sum():,.0f} units")
        
        # FC Breakdown
        fc_summary = api_inv.groupby('fnsku').agg({
            'totalQuantity': 'sum',
            'asin': 'first'
        }).reset_index()
        fc_summary['fc_name'] = fc_summary['fnsku'].apply(get_fc_name)
        fc_summary = fc_summary.groupby('fc_name')['totalQuantity'].sum().reset_index()
        
        st.subheader("🏭 FC Wise Inventory")
        st.dataframe(fc_summary, use_container_width=True)
        
        fig = px.bar(fc_summary, x='fc_name', y='totalQuantity', title="Inventory by FC")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("⚠️ SP-API setup incomplete. Using file upload fallback.")
        api_inv = pd.DataFrame()

# File Upload Processing (Original Features)
if (not use_api or api_inv.empty) and mtr_files and inv_file:
    # Load MTR files
    sales_dfs = []
    for f in mtr_files:
        df = read_file(f)
        if not df.empty:
            sales_dfs.append(df)
    
    if sales_dfs:
        sales = pd.concat(sales_dfs, ignore_index=True)
        
        # Clean sales data
        sales["Sku"] = sales["Sku"].astype(str).str.strip()
        sales["Quantity"] = pd.to_numeric(sales["Quantity"], errors="coerce").fillna(0)
        sales["Shipment Date"] = pd.to_datetime(sales["Shipment Date"], errors="coerce")
        sales = sales.dropna(subset=["Shipment Date"])
        
        sales_period = (sales["Shipment Date"].max() - sales["Shipment Date"].min()).days + 1
        
        # Load inventory
        inv = read_file(inv_file)
        inv["Sku"] = inv["MSKU"].astype(str).str.strip() if "MSKU" in inv.columns else inv["ASIN"].astype(str).str.strip()
        inv["Current Stock"] = pd.to_numeric(inv["Ending Warehouse Balance"], errors="coerce").fillna(0)
        
        # Main planning report
        total_sales = sales.groupby("Sku")["Quantity"].sum().reset_index()
        report = total_sales.merge(
            inv.groupby("Sku")["Current Stock"].sum().reset_index(), 
            on="Sku", how="left"
        )
        report["Current Stock"] = report["Current Stock"].fillna(0)
        report["Avg Daily Sale"] = report["Quantity"] / sales_period
        report["Safety Stock"]
