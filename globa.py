import pandas as pd
import requests
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import streamlit as st
from openpyxl import load_workbook
import tempfile

# --- Load global pricing data ---
usa_data = pd.read_excel("Global Pricing.xlsx", sheet_name="USA")
canada_data = pd.read_excel("Global Pricing.xlsx", sheet_name="Canada")
all_data = pd.concat([usa_data, canada_data], ignore_index=True)

# --- Geocode user-entered address ---
geolocator = Nominatim(user_agent="pricing_app")
def get_coords(address):
    location = geolocator.geocode(address)
    return (location.latitude, location.longitude) if location else None

# --- Find closest comps with safe lat/lon handling ---
def find_closest_comps(user_coords):
    # Drop rows with missing or zero lat/lon
    valid_data = all_data.dropna(subset=['Latitude', 'Longitude']).copy()
    valid_data = valid_data[(valid_data['Latitude'] != 0) & (valid_data['Longitude'] != 0)]

    # Calculate distances
    valid_data['distance'] = valid_data.apply(
        lambda row: geodesic(user_coords, (row['Latitude'], row['Longitude'])).miles, axis=1
    )

    # Sort by distance ascending
    sorted_data = valid_data.sort_values('distance')

    # Get top 2 centres or fill with ""
    closest_comps = sorted_data['Centre'].tolist()[:2]
    while len(closest_comps) < 2:
        closest_comps.append("")

    return closest_comps

# --- Find coworking spaces using Overpass API ---
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
def fill_pricing_template(template_path, centre_name, centre_address, currency, area_units,
                          comp_centres, coworking_names):
    wb = load_workbook(template_path)
    ws = wb['Centre & Market Details']

    ws['C2'] = centre_name
    ws['C3'] = centre_address
    ws['D5'] = currency
    ws['D7'] = area_units

    ws['D17'] = comp_centres[0]
    ws['E18'] = comp_centres[1]

    ws['D30'] = coworking_names[0] if len(coworking_names) > 0 else ""
    ws['E30'] = coworking_names[1] if len(coworking_names) > 1 else ""

    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp_file.name)
    return tmp_file.name

# --- Streamlit UI ---
st.title("Pricing Template 2025 Filler")

centre_name = st.text_input("Centre Name")
centre_address = st.text_input("Centre Address")
currency = st.selectbox("Pricing Currency", ["USD", "CAD"])
area_units = st.selectbox("Area Units", ["SqM", "Sqft"])

if st.button("Generate Template"):
    user_coords = get_coords(centre_address)
    if not user_coords:
        st.error("Could not geocode the given address. Please try a different address.")
    else:
        comp_centres = find_closest_comps(user_coords)
        coworking_names = find_online_coworking_osm(user_coords)

        st.write("Closest Comps:", comp_centres)
        st.write("Closest Coworking Spaces:", coworking_names)

        filled_file = fill_pricing_template("Pricing Template 2025.xlsx",
                                            centre_name, centre_address,
                                            currency, area_units,
                                            comp_centres, coworking_names)

        with open(filled_file, "rb") as f:
            st.download_button(
                label="Download Filled Pricing Template",
                data=f,
                file_name="Pricing_Template_2025_filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
