import streamlit as st
import pandas as pd
from geopy.distance import geodesic
from opencage.geocoder import OpenCageGeocode

# Initialize OpenCage geocoder with your API key
geocoder = OpenCageGeocode("1b5be35755c545c0828be6700389c253")

# Function to get coordinates using OpenCage
def get_coords(address):
    try:
        results = geocoder.geocode(address)
        if results and len(results):
            return (results[0]['geometry']['lat'], results[0]['geometry']['lng'])
        else:
            return (None, None)
    except Exception:
        return (None, None)

# Load Global Pricing Workbook sheets
global_pricing_file = "Global Pricing.xlsx"
usa_data = pd.read_excel(global_pricing_file, sheet_name="USA")
canada_data = pd.read_excel(global_pricing_file, sheet_name="Canada")

# Combine USA and Canada data
all_data = pd.concat([usa_data, canada_data], ignore_index=True)

# Clean data: drop rows missing coords or price if needed
all_data.dropna(subset=['Latitude', 'Longitude', 'Price'], inplace=True)

# UI Inputs
st.title("Global Pricing Tool")

centre_number = st.text_input("Centre #")
centre_address = st.text_input("Centre Address")
pricing_currency = st.selectbox("Pricing Currency", options=["USD", "CAD"])
area_units = st.selectbox("Area Units", options=["SqM", "SqFt"])
total_area_contracted = st.number_input("Total Area Contracted", min_value=0.0)
net_internal_area = st.number_input("Net Internal Area", min_value=0.0)
monthly_market_rent = st.number_input("Monthly Market Rent", min_value=0.0)
source_of_market_rent = st.selectbox("Source of Market Rent", options=[
    "LL or Partner Provided",
    "Broker Provided or Market Report",
    "Benchmarked from similar centre"
])
service_charges = st.number_input("Service Charges", min_value=0.0)
property_tax = st.number_input("Property Tax", min_value=0.0)
total_monthly_expected_cash_flow = st.number_input("Total Monthly Expected Cash Flow Maturity", min_value=0.0)

# Load Pricing Template 2025 workbook
template_file = "Pricing Template 2025.xlsx"
template = pd.ExcelFile(template_file)

# Load the "Centre & Market Details" sheet into a DataFrame
template_df = pd.read_excel(template_file, sheet_name="Centre & Market Details", header=None)

# Get user input coordinates
user_coords = get_coords(centre_address)
if user_coords == (None, None):
    st.error("Unable to geocode the centre address. Please check the address and try again.")
    st.stop()

# Function to calculate distance between user and each comp centre
def calc_distance(row):
    if pd.isna(row['Latitude']) or pd.isna(row['Longitude']):
        return float('inf')  # effectively skip invalid coords
    return geodesic(user_coords, (row['Latitude'], row['Longitude'])).miles

all_data['distance'] = all_data.apply(calc_distance, axis=1)

# Find closest comps (exclude the centre itself if present)
closest_comps = all_data[all_data['Centre #'] != centre_number].sort_values('distance').head(5)

# Find the 2 closest comps for template output
comp1 = closest_comps.iloc[0]
comp2 = closest_comps.iloc[1]

# Calculate average price of 5 closest comps
avg_price_5_comps = closest_comps['Price'].mean()

# Calculate distance for comp1 and comp2, round to 2 decimals
comp1_distance = round(comp1['distance'], 2)
comp2_distance = round(comp2['distance'], 2)

# Quality comparisons for comp1 and comp2
def quality_label(comp_price):
    if comp_price > avg_price_5_comps:
        return "Lesser Quality"
    elif abs(comp_price - avg_price_5_comps) < 1e-2:  # close enough
        return "Same Quality"
    else:
        return "Higher Quality"

comp1_quality = quality_label(comp1['Price'])
comp2_quality = quality_label(comp2['Price'])

# Percentage difference from average for comp1 and comp2
def perc_diff(comp_price):
    return round(100 * (comp_price - avg_price_5_comps) / avg_price_5_comps, 2)

comp1_perc_diff = perc_diff(comp1['Price'])
comp2_perc_diff = perc_diff(comp2['Price'])

# Function to find closest coworking spaces using OpenStreetMap Overpass API
import requests

def find_closest_coworking(lat, lon):
    query = f"""
    [out:json];
    (
      node["office"="coworking"](around:50000,{lat},{lon});
      way["office"="coworking"](around:50000,{lat},{lon});
      relation["office"="coworking"](around:50000,{lat},{lon});
    );
    out center 5;
    """
    url = "http://overpass-api.de/api/interpreter"
    response = requests.post(url, data=query)
    data = response.json()

    places = []
    for element in data['elements']:
        if 'tags' in element and 'name' in element['tags']:
            if 'lat' in element:
                el_lat = element['lat']
                el_lon = element['lon']
            elif 'center' in element:
                el_lat = element['center']['lat']
                el_lon = element['center']['lon']
            else:
                continue
            dist = geodesic((lat, lon), (el_lat, el_lon)).miles
            places.append({'name': element['tags']['name'], 'lat': el_lat, 'lon': el_lon, 'distance': dist})

    # Sort by distance and return top 2
    places = sorted(places, key=lambda x: x['distance'])
    return places[:2]

coworking_spaces = find_closest_coworking(*user_coords)

# For output, handle missing coworking spaces gracefully
cowork1_name = coworking_spaces[0]['name'] if len(coworking_spaces) > 0 else "N/A"
cowork2_name = coworking_spaces[1]['name'] if len(coworking_spaces) > 1 else "N/A"
cowork1_dist = round(coworking_spaces[0]['distance'], 2) if len(coworking_spaces) > 0 else None
cowork2_dist = round(coworking_spaces[1]['distance'], 2) if len(coworking_spaces) > 1 else None

# Create output Excel by loading the template again and writing values

from openpyxl import load_workbook

wb = load_workbook(template_file)
ws = wb["Centre & Market Details"]

# Fill in the template cells
ws["C2"] = centre_number
ws["C3"] = centre_address
ws["D5"] = pricing_currency
ws["D6"] = area_units
ws["D8"] = total_area_contracted
ws["D9"] = None  # clear previous Net Internal Area per your request
ws["D10"] = monthly_market_rent
ws["D11"] = source_of_market_rent
ws["D12"] = service_charges
ws["D13"] = property_tax
ws["D14"] = total_monthly_expected_cash_flow

ws["D17"] = comp1['Centre #']
ws["E17"] = comp2['Centre #']

ws["D18"] = f"{comp1_distance} miles"
ws["E18"] = f"{comp2_distance} miles"

ws["D19"] = comp1_quality
ws["E19"] = comp2_quality

ws["D20"] = comp1_perc_diff
ws["E20"] = comp2_perc_diff

ws["D30"] = cowork1_name
ws["E30"] = cowork2_name

ws["D31"] = f"{cowork1_dist} miles" if cowork1_dist is not None else "N/A"
ws["E31"] = f"{cowork2_dist} miles" if cowork2_dist is not None else "N/A"

# Save output to a file
output_file = "Pricing_Template_Output.xlsx"
wb.save(output_file)

st.success(f"Pricing Template generated: {output_file}")

# Provide download link
with open(output_file, "rb") as f:
    st.download_button(label="Download Completed Pricing Template", data=f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
