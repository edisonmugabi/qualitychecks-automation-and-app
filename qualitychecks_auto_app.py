import streamlit as st
import pandas as pd
import warnings as wn
from datetime import datetime, timedelta
from io import StringIO
import numpy as np
import re
import os

pd.set_option("styler.render.max_elements", 2585583)

# Define the relative path to the image
# image_path = os.path.join(os.path.dirname("C:\\Users\\Edison New\\Pictures\\"), "Screenshots\\RTV log.png")

image_path = os.path.join("C:\\Users\\Edison New\\Desktop\\edison jupyter\\quality checks folder\\", "RTV log.png")
st.sidebar.image(image_path, width=200)

st.title("RTV Daily Quality checks")
st.header("2025 BHS")

data_path = os.getenv("C:\\Users\\Edison New\Desktop\\edison jupyter\\quality checks folder\\", "BHS.xlsx")
data = pd.read_excel(data_path)
# def load_data():
# data = pd.read_excel(r"C:\\Users\\Edison New\Desktop\\edison jupyter\\quality checks folder\\BHS.xlsx")


# Convert 'decision_note' column to string
data['decision_note'] = data['decision_note'].astype(str)

#return df


# converting submissionDate to pandas data time format
data['SubmissionDate'] = pd.to_datetime(data['SubmissionDate'])

#generating the yesterday date for easy tracking samples of yesterday
yesterday = (datetime.today() - timedelta(days=1)).date()

# converting the data of status and consent_1 form surverycto datatypes to pandas float datatype
data[['status', 'consent_1']] = data[['status', 'consent_1']].astype(float)

# filtering the data for the previous day and data for samples fully completed
data = data[
    # (data['SubmissionDate'].dt.date == yesterday) &
    ((data['status'] == 1) | (data['status'] == 7)) &
    (data['consent_1'] == 1)
]

# converting the starttime to datatime in pandas
data['starttime'] = pd.to_datetime(data['starttime'])

# converting the endtime to datatime in pandas
data['endtime']=pd.to_datetime(data['endtime'])

#converting duration to numeric
data['duration'] = pd.to_numeric(data['duration'], errors='coerce') 
column_drop_path = os.getenv("C:\\Users\\Edison New\\Desktop\\edison jupyter\\quality checks folder\\", "qualitychecks_columns_drop.xlsx")
column_drop_df = pd.read_excel(column_drop_path)
# importing excel files with columns to drop to enhance styling process of the dataframe
# please include all columns to drop in this sheet
# column_drop_df=pd.read_excel(r"C:\\Users\\Edison New\\Desktop\\edison jupyter\\quality checks folder\\qualitychecks_columns_drop.xlsx")

# creating function to drop to drop columns and  dropping columns from data
def columns_drop(data, column_drop_df):
    columns_to_drop = column_drop_df['Column Names'].to_list()
    data = data.drop(columns=columns_to_drop, errors='ignore')
    return data
# dropping un wanted columns from data
data = columns_drop(data, column_drop_df)

# converting the duration from seconds to minutes under new variable duration2
data['duration2']=data['duration']/60

# creating the column check for samples started beyond 8PM
data['is_start_time_beyond_8pm'] = (pd.to_datetime(data['starttime']).dt.hour >= 20).astype(int)

# checking whether there if date difference between when the sample was started vs endtime of the sample
data['starttime'] = pd.to_datetime(data['starttime']).dt.date
data['endtime'] = pd.to_datetime(data['endtime']).dt.date
data['date_difference'] = (data['endtime'] - data['starttime']).apply(lambda x: x.days)



# Creating columns for samples collected by enumerator the pervious day
data['sample'] = data['enumerator_name'].map(data['enumerator_name'].value_counts())

data['is_duration_invalid']=((data['duration2']<20) | (data['duration2']>60)).astype(int)

# Creating columns for samples collected by enumerator the pervious day
data['sample'] = data['enumerator_name'].map(data['enumerator_name'].value_counts())
price_path = os.getenv("C:\\Users\\Edison New\\Desktop\\edison jupyter\\quality checks folder\\", "crop_prices.xlsx")
price_df = pd.read_excel(price_path)


# importing the excel sheet with min,max of major seasonal crops,vegetables, perennial crops
# price_df=pd.read_excel(r"C:\\Users\\Edison New\\Desktop\\edison jupyter\\quality checks folder\\crop_prices.xlsx")

list_ = ['beans', 'maize','peas','cassava']  
season = [1, 2]
state = ['fresh', 'dry']
price_1 = []

for i in season:
    for j in list_:
        for k in state:
            price1 = f'sn_{i}_{j}_Market_Price_{k}'
            if price1 in data.columns:
                price_1.append(price1)  



list=['gnuts','yams','sweetpotatoes','irish_potatoes',
      'ginger','garlic','rice','sorghum','millet',
      'soya_beans']
season=[1,2]
price_2=[]
for i in list:
    for j in season:
        price1=f'sn_{j}_{i}_Market_Price'
        if price1 in data.columns:
            price_2.append(price1)       

# merging prices of seasonal crops into one list for ease accessibility
season_crop_prices=price_1+price_2

#Converting all season crop prices to numeric from json data structures
for column in season_crop_prices:
    data[column] = pd.to_numeric(data[column], errors='coerce')

# extracting vegetables prices from the data to lists for easy accessibility
veg = [
    "Pumpkins",
    "Carrots",
    "onions",
    "green_pepper",
    "hot_pepper",
    "cabbage",
    "tomato",
    "watermelon",
    "spinach",
    "cauliflower",
    "sukuma_wiki",
    "beetroot",
    "blacknightshade",
    "white_eggplant",
    "green_eggplant",
    "purple_eggplant",
    "dodo"
]
season=[1,2]
veg_list=[]
for k in veg:
    for j in season:
        vegk=f'sn_{j}_{k}_Market_Price'
        if vegk in data.columns:
            veg_list.append(vegk)


# list of all seasonal crops ans seasons
crops=['gnuts','beans','irish_potatoes','maize','peas',
      'ginger','garlic','rice','barley','sorghum','millet',
      'soya_beans']
seasons=[1,2]

# Converting all seasonal crop total_yield and total quantity planted to numeric
for crop in crops:
    for season in seasons:
        total_yield = f'sn_{season}_{crop}_Total_Yield'
        planted = f'sn_{season}_{crop}_planted'
        
        if total_yield not in data.columns:
            data[total_yield] = np.nan
        if planted not in data.columns:
            data[planted] = np.nan
        data[total_yield] = pd.to_numeric(data[total_yield], errors='coerce')
        data[planted] = pd.to_numeric(data[planted], errors='coerce')


# creating function to compute yield per  unit
def calculate_yield_per_plant(data, crops, seasons):
    result = data.copy()
    
    for crop in crops:
        for season in seasons:
            total_yield = f'sn_{season}_{crop}_Total_Yield'
            planted = f'sn_{season}_{crop}_planted'
            yield_per_plant = f'sn_{season}_{crop}_yp'
            
            # Ensure the yield_per_plant column exists
            if yield_per_plant not in result.columns:
                result[yield_per_plant] = np.nan

            if total_yield in result.columns and planted in result.columns:
                # Identify rows where both total_yield and planted are missing
                null_mask = result[total_yield].isna() & result[planted].isna()
                result.loc[null_mask, yield_per_plant] = np.nan  # Explicitly set to NaN
                
                # Compute yield per plant where valid values exist
                mask = (
                    pd.notna(result[total_yield]) & 
                    pd.notna(result[planted]) & 
                    (result[planted] != 0)
                )
                
                if mask.any():
                    result.loc[mask, yield_per_plant] = (
                        result.loc[mask, total_yield] / result.loc[mask, planted]
                    )
                    
                    # Replace infinite values with NaN
                    result[yield_per_plant].replace([np.inf, -np.inf], np.nan, inplace=True)
    
    return result
data = calculate_yield_per_plant(data, crops, seasons)

# converting all these columns below to numeric 
required_columns = [
    'Distance_travelled_one_way_OPD_treatment', 'Time_travel_one_way_trip_OPD_treatment_minutes',
    'water_distance_collect_water_round_trip', 'hh_water_collection_Minutes', 'distance_primary_market',
    'time_primary_market', 'Size_land_owned', 'Value_land_owned', "govt_assistance_value", 
    "remittance_gifts_children", "remittance_gifts_relative", "remittance_gifts_friends",
    "casual_work_wage_member1", "casual_work_inkind_member1", "casual_work_wage_member2", 
    "casual_work_inkind_member2", "casual_work_wage_member3", "casual_work_inkind_member3", 
    "casual_work_wage_member4", "casual_work_inkind_member4", "casual_work_wage_member5", 
    "casual_work_inkind_member5", "casual_work_wage_member6", "casual_work_inkind_member6"
]

for col in required_columns:
    if col not in data.columns:
        data[col] = np.nan
data[required_columns] = data[required_columns].astype(float)

# converting the other columns to numeric
col_price = [i for i in data.columns if "price" in i or "payment_fees_labour" in i or "remittance" in i or "market_ppu" in i or "Price" in i]
data[col_price] = data[col_price].apply(pd.to_numeric, errors='coerce')

# function computing total number of errors
price_df = price_df.set_index('Crop')
def check_price_violations(row):
    errors = 0  
    for crop in price_df.index:
        if crop in row:  
            if row[crop] < price_df.loc[crop, 'Min'] or row[crop] > price_df.loc[crop, 'Max']:
                errors += 1  
    return errors
data['price_errors'] = data.apply(check_price_violations, axis=1)
price_df=price_df.reset_index('Crop')

col = [i for i in data.columns if "Value_remittance_gift" in i or "Value_payment_fees_labour" in i or 'remittance' in i]

def extract_crop(col_name):
    match = re.search(r'sn_\d+_([a-zA-Z_]+)_Yield|sn_\d+_([a-zA-Z_]+)_Value|([a-zA-Z_]+)_Yield|([a-zA-Z_]+)_Value|([a-zA-Z_]+)_remittance|([a-zA-Z_]+)_payment_fees', col_name)
    if match:
        return match.group(1) or match.group(2) or match.group(3) or match.group(4) or match.group(5) or match.group(6)
    return "Unknown"
df = pd.DataFrame({'Column Name': col, 'Crop Name': [extract_crop(coll) for coll in col]})
data[col] = data[col].apply(pd.to_numeric, errors='coerce')

business_number = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,97] 
for bus in business_number:
    bus_sa = f'business{bus}_sales'
    bus_po = f'business{bus}_profit'
    if bus_sa not in data.columns and bus_po not in data.columns:
        data[bus_sa]=np.nan
        data[bus_po]=np.nan
    
    if bus_sa in data.columns and bus_po in data.columns:
        data[[bus_sa,bus_po]]=data[[bus_sa,bus_po]].astype(float)
    else:
        print(f"Columns {bus_sa} or {bus_po} do not exist in the dataset.")


def business_error(row):
    error = 0
    for i in business_number:
        bus_sa = f'business{i}_sales'
        bus_po = f'business{i}_profit'
        if pd.notna(row[bus_po]) and pd.notna(row[bus_sa]):
            if row[bus_po] > row[bus_sa]:
                error += 1
    return error
data['bussiness_errors']=data.apply(business_error,axis=1)

# computing the error computed in the remmittance values
def count_remittance_highlights(row):
    highlight_count = 0  

    # Identify relevant columns
    cols = [col for col in row.index if any(x in str(col) for x in 
                ["Value_remittance_gift", "Value_payment_fees_labour", "remittance_gifts_friends"])]

    # Iterate through selected columns
    for col_name in cols:
        if pd.notna(row[col_name]):  
            if row[col_name] < 100 or row[col_name] % 100 != 0:
                highlight_count += 1  

    return highlight_count
data['remittance_errors'] = data.apply(count_remittance_highlights, axis=1)

def compute_travel_errors(row):
    error_count = 0  # Initialize error counter

    # Define limits for distance and time
    travel_checks = {
        'Distance_travelled_one_way_OPD_treatment': (0, 28),
        'Time_travel_one_way_trip_OPD_treatment_minutes': (0, 420),
        'water_distance_collect_water_round_trip': (0, 8),
        'hh_water_collection_Minutes': (0, 420),
        'distance_primary_market': (0, 10),
        'time_primary_market': (0, 420)
    }

    # Count errors for extreme values (0 or exceeding max limit)
    for col, (min_val, max_val) in travel_checks.items():
        if col in row:
            if row[col] == min_val or row[col] >= max_val:
                error_count += 1  # Increment error count

    # Walking speed rule: Expected travel time based on distance
    if 'Distance_travelled_one_way_OPD_treatment' in row and 'Time_travel_one_way_trip_OPD_treatment_minutes' in row:
        expected_time_opd = row['Distance_travelled_one_way_OPD_treatment'] * 40
        if row['Time_travel_one_way_trip_OPD_treatment_minutes'] > 1.2 * expected_time_opd:
            error_count += 1

    if 'water_distance_collect_water_round_trip' in row and 'hh_water_collection_Minutes' in row:
        expected_time_water = row['water_distance_collect_water_round_trip'] * 60
        if row['hh_water_collection_Minutes'] > 1.2 * expected_time_water:
            error_count += 1

    if 'distance_primary_market' in row and 'time_primary_market' in row:
        expected_time_market = row['distance_primary_market'] * 40
        if row['time_primary_market'] > 1.2 * expected_time_market:
            error_count += 1

    return error_count  # Return total errors for the row

# Apply function to compute errors
data['distance_time_errors'] = data.apply(compute_travel_errors, axis=1)

def compute_crop_yield_errors(row):
    error_count = 0  # Initialize error counter

    # Define crop yield thresholds
    crop_yield_thresholds = {
        'sn_1_beans_yp': 100, 'sn_2_beans_yp': 100,
        'sn_1_maize_yp': 250, 'sn_2_maize_yp': 250,
        'sn_1_peas_yp': 150, 'sn_2_peas_yp': 150,
        'sn_1_gnuts_yp': 150, 'sn_2_gnuts_yp': 150,
        'sn_1_irish_potatoes_yp': 10, 'sn_2_irish_potatoes_yp': 10,
        'sn_1_ginger_yp': 100, 'sn_2_ginger_yp': 100,
        'sn_1_rice_yp': 150, 'sn_2_rice_yp': 150,
        'sn_1_barley_yp': 100, 'sn_2_barley_yp': 100,
        'sn_1_sorghum_yp': 100, 'sn_2_sorghum_yp': 100,
        'sn_1_millet_yp': 200, 'sn_2_millet_yp': 200,
        'sn_1_soya_beans_yp': 100, 'sn_2_soya_beans_yp': 100,
        'sn_1_garlic_yp': 100, 'sn_2_garlic_yp': 100
    }

    # Check for crop yield exceeding thresholds
    for crop, threshold in crop_yield_thresholds.items():
        if crop in row and row[crop] > threshold:
            error_count += 1  # Increment error count

    return error_count  # Return total errors for the row

# Apply function to compute crop yield errors
data['yield_per_unit_errors'] = data.apply(compute_crop_yield_errors, axis=1)

# function computing the errors in land errors
def compute_land_value_errors(row):
    error_count = 0  # Initialize error counter

    # Land Value Check
    if 'Size_land_owned' in row and 'Value_land_owned' in row:
        expected_land_value = row['Size_land_owned'] * 15_000_000  # Expected value per unit size
        if row['Value_land_owned'] > expected_land_value:
            error_count += 1  # Increment error count

    return error_count  # Return total errors for the row

# Apply function to compute land value errors
data['land_value_errors'] = data.apply(compute_land_value_errors, axis=1)

#display pivot_errors 

pivot=data.pivot_table(index=['pre_district','enumerator_name'],values=['price_errors','bussiness_errors','land_value_errors',
                                                                        'remittance_errors','distance_time_errors','yield_per_unit_errors'],aggfunc="sum")
pivot=pivot.reset_index()
pivot['Total_errors'] = pivot[['price_errors', 'bussiness_errors', 'land_value_errors',
                               'remittance_errors', 'distance_time_errors', 'yield_per_unit_errors']].sum(axis=1)
# Inject Custom CSS for Uniform Box Sizes and Colors
st.markdown("""
    <style>
    .metric-box {
        width: 100%;  /* Ensures full width within the column */
        min-height: 130px; /* Fixed height for uniformity */
        border-radius: 12px;
        text-align: center;
        font-size: 18px;
        font-weight: bold;
        padding: 15px;
        margin: 5px;
        box-shadow: 2px 2px 10px rgba(0, 0, 0, 0.1);
    }
    .total-errors { background-color: #FF6B6B; color: white; } /* Red */
    .price-errors { background-color: #4ECDC4; color: black; } /* Aqua */
    .distance-time-errors { background-color: #FFB74D; color: black; } /* Orange */
    .bussiness-errors { background-color: #A29BFE; color: white; } /* Purple */
    .land-value-errors { background-color: #FFD700; color: black; } /* Gold */
    .remittance-errors { background-color: #66BB6A; color: white; } /* Green */
    .yield-per-unit-errors { background-color: #42A5F5; color: white; } /* Blue */
    </style>
""", unsafe_allow_html=True)

# Function to display metric inside a styled box
def styled_metric(label, value, css_class):
    st.markdown(f"""
        <div class="metric-box {css_class}">
            <p>{label}</p>
            <h1>{value}</h1>
        </div>
    """, unsafe_allow_html=True)

# First Row: Centered "Overall Total Errors"
_, col, _ = st.columns([1, 2, 1])  # Middle column takes twice the space
with col:
    styled_metric("Overall Total Errors", pivot['Total_errors'].sum(), "total-errors")

# Second Row: Three metrics with space between them
col1, _, col2, _, col3 = st.columns([1, 0.3, 1, 0.3, 1])  # Adds spacing
with col1:
    styled_metric("Overall Total Price Errors", pivot['price_errors'].sum(), "price-errors")
with col2:
    styled_metric("Overall Distance/Time Errors", pivot['distance_time_errors'].sum(), "distance-time-errors")
with col3:
    styled_metric("Overall Business Errors", pivot['bussiness_errors'].sum(), "bussiness-errors")

# Third Row: Three metrics with space between them
col1, _, col2, _, col3 = st.columns([1, 0.3, 1, 0.3, 1])  # Adds spacing
with col1:
    styled_metric("Overall Land Value Errors", pivot['land_value_errors'].sum(), "land-value-errors")
with col2:
    styled_metric("Overall Total Remittance Value Errors", pivot['remittance_errors'].sum(), "remittance-errors")
with col3:
    styled_metric("Overall Yield Per Unit Errors", pivot['yield_per_unit_errors'].sum(), "yield-per-unit-errors")
st.divider()
st.dataframe(pivot)
st.download_button(
    label="Download Data",
    data=pivot.to_csv().encode("utf-8"),
    file_name="Summary_sheet_for_errors_in_the_report.csv",
    mime="text/csv"
)
st.divider()
ppivot=pivot.groupby('pre_district')['Total_errors'].sum().reset_index()
st.subheader("Total errors per district")
st.bar_chart(ppivot.set_index('pre_district'))
# Function to apply the red background if out of range
def highlight_out_of_range(value, min_price, max_price):
    if value < min_price or value > max_price:
        return 'background-color: red'
    return ''

# First function to apply the price highlighting
def apply_highlighting(data, price_df):
    crop_prices = {row["Crop"].strip(): (row["Min"], row["Max"]) for _, row in price_df.iterrows()}

    def highlight_column(column):
        crop_name = column.name.strip()
        price_range = crop_prices.get(crop_name)
        if price_range:
            min_price, max_price = price_range
            return column.apply(lambda x: highlight_out_of_range(x, min_price, max_price))
        return [""] * len(column)

    return data.style.apply(highlight_column, subset=data.columns)

# Second function to apply time-related checks and other conditions
def highlight_time(row):
    colors = [''] * len(row)

    # Highlight remittance and payment fee values
    col = [i for i in data.columns if "Value_remittance_gift" in i or "Value_payment_fees_labour" in i or 'remittance_gifts_friends' in i]
    for col_name in col:
        if col_name in row.index and not pd.isna(row[col_name]):
            if row[col_name] < 100 or row[col_name] % 100 != 0:
                colors[row.index.get_loc(col_name)] = 'background-color:red'

    # Sample Check
    if row['sample'] < 6:
        colors[row.index.get_loc('sample')] = 'background-color:red'

    # OPD Travel Checks
    if row['Distance_travelled_one_way_OPD_treatment'] == 0 or row['Distance_travelled_one_way_OPD_treatment'] >= 28:
        colors[row.index.get_loc('Distance_travelled_one_way_OPD_treatment')] = 'background-color:red'

    if row['Time_travel_one_way_trip_OPD_treatment_minutes'] == 0 or row['Time_travel_one_way_trip_OPD_treatment_minutes'] >= 420:
        colors[row.index.get_loc('Time_travel_one_way_trip_OPD_treatment_minutes')] = 'background-color:red'

    # Water Collection Checks
    if row['water_distance_collect_water_round_trip'] == 0 or row['water_distance_collect_water_round_trip'] > 8:
        colors[row.index.get_loc('water_distance_collect_water_round_trip')] = 'background-color:red'

    if row['hh_water_collection_Minutes'] == 0 or row['hh_water_collection_Minutes'] >= 420:
        colors[row.index.get_loc('hh_water_collection_Minutes')] = 'background-color:red'

    # Market Travel Checks
    if row['distance_primary_market'] == 0 or row['distance_primary_market'] > 10:
        colors[row.index.get_loc('distance_primary_market')] = 'background-color:red'

    if row['time_primary_market'] == 0 or row['time_primary_market'] >= 420:
        colors[row.index.get_loc('time_primary_market')] = 'background-color:red'

    # 🚨 Walking Speed Rule Check (Only for high travel time)
    expected_time_opd = row['Distance_travelled_one_way_OPD_treatment'] * 40
    if row['Time_travel_one_way_trip_OPD_treatment_minutes'] > 1.2 * expected_time_opd:
        colors[row.index.get_loc('Distance_travelled_one_way_OPD_treatment')] = 'background-color:red'
        colors[row.index.get_loc('Time_travel_one_way_trip_OPD_treatment_minutes')] = 'background-color:red'

    expected_time_water = row['water_distance_collect_water_round_trip'] * 60
    if row['hh_water_collection_Minutes'] > 1.2 * expected_time_water:
        colors[row.index.get_loc('water_distance_collect_water_round_trip')] = 'background-color:red'
        colors[row.index.get_loc('hh_water_collection_Minutes')] = 'background-color:red'

    expected_time_market = row['distance_primary_market'] * 40
    if row['time_primary_market'] > 1.2 * expected_time_market:
        colors[row.index.get_loc('distance_primary_market')] = 'background-color:red'
        colors[row.index.get_loc('time_primary_market')] = 'background-color:red'

    # Business Check
    business_number = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,97]
    for bus in business_number:
        bus_sa = f'business{bus}_sales'
        bus_po = f'business{bus}_profit'
        if bus_sa in row.index and bus_po in row.index:
            if row[bus_po] > row[bus_sa]:
                colors[row.index.get_loc(bus_po)] = 'background-color:red'
                colors[row.index.get_loc(bus_sa)] = 'background-color:red'

    # Crop Yield Checks
    crop_yield_thresholds = {
        'sn_1_beans_yp': 100, 'sn_2_beans_yp': 100,
        'sn_1_maize_yp': 250, 'sn_2_maize_yp': 250,
        'sn_1_peas_yp': 150, 'sn_2_peas_yp': 150,
        'sn_1_gnuts_yp': 150, 'sn_2_gnuts_yp': 150,
        'sn_1_irish_potatoes_yp': 10, 'sn_2_irish_potatoes_yp': 10,
        'sn_1_ginger_yp': 100, 'sn_2_ginger_yp': 100,
        'sn_1_rice_yp': 150, 'sn_2_rice_yp': 150,
        'sn_1_barley_yp': 100, 'sn_2_barley_yp': 100,
        'sn_1_sorghum_yp': 100, 'sn_2_sorghum_yp': 100,
        'sn_1_millet_yp': 200, 'sn_2_millet_yp': 200,
        'sn_1_soya_beans_yp': 100, 'sn_2_soya_beans_yp': 100,
        'sn_1_garlic_yp': 100, 'sn_2_garlic_yp': 100
    }

    for crop, threshold in crop_yield_thresholds.items():
        if crop in row.index and row[crop] > threshold:
            colors[row.index.get_loc(crop.replace("_yp", "_planted"))] = 'background-color:red'
            colors[row.index.get_loc(crop.replace("_yp", "_Total_Yield"))] = 'background-color:red'
            colors[row.index.get_loc(crop)] = 'background-color:red'

    # Land Value Check
    if 'Size_land_owned' in row.index and 'Value_land_owned' in row.index:
        expected_land_value = row['Size_land_owned'] * 15000000
        if row['Value_land_owned'] > expected_land_value:
            colors[row.index.get_loc('Size_land_owned')] = 'background-color:red'
            colors[row.index.get_loc('Value_land_owned')] = 'background-color:red'

    # Time Checks
    if 'is_start_time_beyond_8pm' in row.index and row['is_start_time_beyond_8pm'] == 1:
        colors[row.index.get_loc('starttime')] = 'background-color:red'

    if 'is_duration_invalid' in row.index and row['is_duration_invalid'] == 1:
        colors[row.index.get_loc('duration2')] = 'background-color:red'

    # Date Difference Check
    if 'date_difference' in row.index and row['date_difference'] > 0:
        colors[row.index.get_loc('starttime')] = 'background-color:red'
        colors[row.index.get_loc('endtime')] = 'background-color:red'


    return colors

# Apply both highlighting functions
data = apply_highlighting(data, price_df).apply(highlight_time, axis=1)  # Apply both functions

# Streamlit UI for showing the data and download option
st.divider()
st.caption(f"Quality checks for {yesterday}")

with st.expander("View Detailed Data"):
    st.dataframe(data)

# Access the underlying DataFrame before iterating over rows
data_frame = data.data  # Access the DataFrame from the Styler object

with pd.ExcelWriter(f'quality_checks_{yesterday}.xlsx', engine='xlsxwriter',
                    engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
    data_frame.to_excel(writer, index=False, sheet_name='QualityChecks')
    workbook = writer.book
    worksheet = writer.sheets['QualityChecks']

    # Define red background format for out-of-range values
    red_format = workbook.add_format({'bg_color': '#FF0000'})

    # Apply red background for price range and other checks
    crop_prices = {row["Crop"].strip(): (row["Min"], row["Max"]) for _, row in price_df.iterrows()}
    for row_num, row in data_frame.iterrows():
        for col_num, value in enumerate(row):
            crop_name = data_frame.columns[col_num].strip()
            price_range = crop_prices.get(crop_name)
            if price_range:
                min_price, max_price = price_range
                if pd.notna(value) and (value < min_price or value > max_price):
                    worksheet.write(row_num + 1, col_num, value, red_format)
                else:
                    worksheet.write(row_num + 1, col_num, value)
            else:
                if pd.api.types.is_numeric_dtype(value):
                    worksheet.write(row_num + 1, col_num, value)
                else:
                    worksheet.write_string(row_num + 1, col_num, str(value))

        # Apply highlight_time red background checks
        colors = highlight_time(row)
        for col_num, color in enumerate(colors):
            if color:  # Apply red background from highlight_time function
                worksheet.write(row_num + 1, col_num, row[col_num], red_format)

    # Header formatting
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'})
    for col_num, value in enumerate(data_frame.columns.values):
        worksheet.write(0, col_num, value, header_format)

    workbook = writer.book
    worksheet = writer.sheets['QualityChecks']

    # Define red background format for out-of-range values
    red_format = workbook.add_format({'bg_color': '#FF0000'})

    # Apply red background for price range and other checks
    crop_prices = {row["Crop"].strip(): (row["Min"], row["Max"]) for _, row in price_df.iterrows()}
    for row_num, row in data_frame.iterrows():
        for col_num, value in enumerate(row):
            crop_name = data_frame.columns[col_num].strip()
            price_range = crop_prices.get(crop_name)
            if price_range:
                min_price, max_price = price_range
                if pd.notna(value) and (value < min_price or value > max_price):
                    worksheet.write(row_num + 1, col_num, value, red_format)
                else:
                    worksheet.write(row_num + 1, col_num, value)
            else:
                worksheet.write(row_num + 1, col_num, value)

        # Apply highlight_time red background checks
        colors = highlight_time(row)
        for col_num, color in enumerate(colors):
            if color:  # Apply red background from highlight_time function
                worksheet.write(row_num + 1, col_num, row[col_num], red_format)

    # Header formatting
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC'})
    for col_num, value in enumerate(data_frame.columns.values):
        worksheet.write(0, col_num, value, header_format)

# Provide the download link for the Excel file
st.download_button(
    label="Download data as Excel",
    data=open(f'quality_checks_{yesterday}.xlsx', 'rb').read(),
    file_name=f'quality_checks_{yesterday}.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
