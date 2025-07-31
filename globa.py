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

# --- Only geocode user address ---
geolocator = Nominatim(user_agent="pricing_app")
def get_coords(address):
    location = geolocator.geocode(address)
    return (location.latitude, location.longitude) if location else None

# --- Find closest comps using stored coordinates ---
def find_closest_comps(user_coords):
    def calc_distance(row):
        return geodesic(user_coords, (row['Latitude'], row['Longitude'])).miles
    all_data['distance'] = all_data.apply(calc_distance, axis=1)
    closest_2 = all_data.nsmallest(2, 'distance')
    return closest_2['Centre'].tolist()

# --- Find coworking using Overpass ---
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

# --- Fill template ---
def fill_pricing_template(template_path, centre_name, centre_address, currency, area_units,
                          comp_centres, coworking_names):
    wb = load_workbook(template_path)
    ws = wb['Centre & Market Details']

    ws['C2'] = centre_name
    ws['C3'] = centre_address
    ws['D5'] = currency
    ws['D7'] = area_units
    ws['D17'] = comp_centres[0] if len(comp_centres) > 0 else ""
    ws['E18'] = comp_centres[1] if len(comp_centres) > 1 else ""
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
        st.error("Could not geocode the given address.")
    else:
        comp_centres = find_closest_comps(user_coords)
        coworking_names = find_online_coworking_osm(user_coords)
        filled_file = fill_pricing_template("Pricing Template 2025.xlsx",
                                            centre_name, centre_address,
                                            currency, area_units,
                                            comp_centres, coworking_names)
        with open(filled_file, "rb") as f:
            st.download_button("Download Filled Pricing Template", f,
                file_name="Pricing_Template_2025_filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
