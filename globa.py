import pandas as pd
import requests
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import streamlit as st
from openpyxl import load_workbook
import tempfile

# --- Load global pricing data and clean columns ---
usa_data = pd.read_excel("Global Pricing.xlsx", sheet_name="USA")
canada_data = pd.read_excel("Global Pricing.xlsx", sheet_name="Canada")

def clean_columns(df):
    df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
    return df

usa_data = clean_columns(usa_data)
canada_data = clean_columns(canada_data)

all_data = pd.concat([usa_data, canada_data], ignore_index=True)
all_data = clean_columns(all_data)

# --- Geocode user-entered address ---
geolocator = Nominatim(user_agent="pricing_app")
def get_coords(address):
    location = geolocator.geocode(address)
    return (location.latitude, location.longitude) if location else None

# --- Find closest comps and distances (rounded to 2 decimals, miles) ---
def find_closest_comps(user_coords):
    valid_data = all_data.dropna(subset=['Latitude', 'Longitude']).copy()
    valid_data = valid_data[(valid_data['Latitude'] != 0) & (valid_data['Longitude'] != 0)]

    valid_data['distance'] = valid_data.apply(
        lambda row: round(geodesic(user_coords, (row['Latitude'], row['Longitude'])).miles, 2),
        axis=1
    )

    sorted_data = valid_data.sort_values('distance')
    comps = sorted_data[['Centre #', 'Latitude', 'Longitude', 'distance']].head(2)

    comp_centres = comps['Centre #'].tolist()
    comp_distances = [f"{d} mi" for d in comps['distance'].tolist()]

    while len(comp_centres) < 2:
        comp_centres.append("")
        comp_distances.append("")

    return comp_centres, comp_distances

# --- Find coworking spaces ---
def find_online_coworking_osm(user_coords):
    lat, lon = user_coords
    overpass_url = "http://overpass-api.de/api/interpreter"
    query = f"""
    [out:json];
    node["office"="coworking"](around:5000,{lat},{lon});
    out;
    """
    response = requests.get(overpass_url, params={'data': query})
    data = response.json()
    coworking_names = []
    for element in data.get('elements', []):
        name = element['tags'].get('name')
        if name:
            coworking_names.append(name)
        if len(coworking_names) == 2:
            break
    return coworking_names

# --- Fill Pricing Template 2025.xlsx ---
def fill_pricing_template(template_path, centre_num, centre_address, currency,
                          area_units, total_area, net_internal_area,
                          monthly_rent, rent_source,
                          service_charges, property_tax,
                          comp_centres, comp_distances, coworking_names):
    wb = load_workbook(template_path)
    ws = wb['Centre & Market Details']

    ws['C2'] = centre_num
    ws['C3'] = centre_address
    ws['D5'] = currency
    ws['D6'] = area_units
    ws['D8'] = net_internal_area  # Net Internal Area
    ws['D9'] = ""                 # clear D9
    ws['D10'] = monthly_rent
    ws['D11'] = rent_source
    ws['D12'] = service_charges
    ws['D13'] = property_tax

    ws['D17'] = comp_centres[0]
    ws['E17'] = comp_centres[1]
    ws['D18'] = comp_distances[0]  # with "mi"
    ws['E18'] = comp_distances[1]  # with "mi"

    ws['D30'] = coworking_names[0] if len(coworking_names) > 0 else ""
    ws['E30'] = coworking_names[1] if len(coworking_names) > 1 else ""

    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp_file.name)
    return tmp_file.name

# --- Streamlit UI ---
st.title("Pricing Template 2025 Filler")

centre_num = st.text_input("Centre #")
centre_address = st.text_input("Centre Address")
currency = st.selectbox("Pricing Currency", ["USD", "CAD"])
area_units = st.selectbox("Area Units", ["SqM", "SqFt"])
total_area = st.number_input("Total Area Contracted", min_value=0.0, format="%.2f")
net_internal_area = st.number_input("Net Internal Area", min_value=0.0, format="%.2f")
monthly_rent = st.number_input("Monthly Market Rent", min_value=0.0, format="%.2f")
rent_source = st.selectbox("Source of Market Rent", 
                           ["LL or Partner Provided", "Broker Provided or Market Report", "Benchmarked from similar centre"])
service_charges = st.number_input("Service Charges", min_value=0.0, format="%.2f")
property_tax = st.number_input("Property Tax", min_value=0.0, format="%.2f")

if st.button("Generate Template"):
    user_coords = get_coords(centre_address)
    if not user_coords:
        st.error("Could not geocode the given address. Please try a different address.")
    else:
        comp_centres, comp_distances = find_closest_comps(user_coords)
        coworking_names = find_online_coworking_osm(user_coords)

        st.write("Closest Comps:", comp_centres)
        st.write("Distances:", comp_distances)
        st.write("Closest Coworking Spaces:", coworking_names)

        filled_file = fill_pricing_template(
            "Pricing Template 2025.xlsx",
            centre_num, centre_address, currency,
            area_units, total_area, net_internal_area,
            monthly_rent, rent_source,
            service_charges, property_tax,
            comp_centres, comp_distances, coworking_names
        )

        with open(filled_file, "rb") as f:
            st.download_button(
                label="Download Filled Pricing Template",
                data=f,
                file_name="Pricing_Template_2025_filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
