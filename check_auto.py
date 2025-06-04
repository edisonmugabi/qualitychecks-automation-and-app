import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
import os
import openpyxl

# --- Load Data ---
data = pd.read_excel("BHS.xlsx")
columns_drop = pd.read_excel("qualitychecks_columns_drop.xlsx")
crop_prices = pd.read_excel("crop_prices.xlsx")

# --- Filter and clean data ---
data = data.query("(status ==1 or status == 7) and consent_1 == 1")
data = data.drop(columns=columns_drop['columns'].tolist(), errors='ignore')
data['duration_minutes'] = pd.to_numeric(data['duration_minutes'], errors='coerce')
data['start_time'] = pd.to_datetime(data['start'], errors='coerce')
data['price_errors'] = 0

# --- Define Checks ---
data['duration2'] = ((data['duration_minutes'] < 20) | (data['duration_minutes'] > 180)).astype(int)
data['late_submission'] = (data['start_time'].dt.hour >= 20).astype(int)

# Price check
for _, row in crop_prices.iterrows():
    crop, col, p_min, p_max = row['Crop'], row['column'], row['min_price'], row['max_price']
    if col in data.columns:
        data['price_errors'] += ((data[col] < p_min) | (data[col] > p_max)).astype(int)

# Remittance check
data['remit_error'] = (data['remit_amount_ugx'] % 100 != 0).astype(int)

# Travel checks
data['travel_error'] = ((data['hh_hf_distance_km'] > 40) | (data['hh_hf_distance_km'] < 0.5) | (data['hh_hf_time_mins'] > 300)).astype(int)

# Yield checks
data['yield_error'] = (data['hh_maize_yield_kg'] > 5000).astype(int)

# Business check
data['profit_error'] = ((data['hh_business_profit'] > data['hh_business_revenue']) & data['hh_business_revenue'].notnull()).astype(int)

# Land value check
data['land_value_error'] = ((data['hh_agric_land_value'] > 200000000) | (data['hh_agric_land_value'] < 50000)).astype(int)

# Total errors
data['total_errors'] = data[['duration2', 'late_submission', 'price_errors', 'remit_error', 'travel_error', 'yield_error', 'profit_error', 'land_value_error']].sum(axis=1)

# --- Display metrics ---
st.title("RTV Daily Quality Check Dashboard")

st.metric("Total Interviews", len(data))
st.metric("Interviews with Errors", (data['total_errors'] > 0).sum())

# --- Display error breakdown ---
error_cols = ['duration2', 'late_submission', 'price_errors', 'remit_error', 'travel_error', 'yield_error', 'profit_error', 'land_value_error']
error_summary = data[error_cols].sum().reset_index()
error_summary.columns = ['Error Type', 'Count']
st.dataframe(error_summary)

# --- Error distribution by district ---
st.subheader("Error Count by District")
if 'pre_district' in data.columns:
    district_summary = data.groupby('pre_district')['total_errors'].sum().reset_index()
    fig = px.bar(district_summary, x='pre_district', y='total_errors', title='Total Errors by District')
    st.plotly_chart(fig)

# --- Data table ---
st.subheader("Detailed Data Table")
st.dataframe(data)

# --- Download data ---
st.download_button(
    label="Download Cleaned Data",
    data=data.to_csv(index=False),
    file_name="cleaned_data.csv",
    mime="text/csv"
)
