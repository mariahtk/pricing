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

def clean_columns(df):
    df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
    return df

usa_data = clean_columns(usa_data)
canada_data = clean_columns(canada_data)
all_data = pd.concat([usa_data, canada_data], ignore_index=True)
all_data = clean_columns(all_data)

# --- Geocode ---
geolocator = Nominatim(user_agent="pricing_app")
def get_coords(address):
    location = geolocator.geocode(address)
    return (location.latitude, location.longitude) if location else None

def format_diff(value):
    if value > 0:
        return f"{abs(value)}% higher"
    elif value < 0:
        return f"{abs(value)}% lower"
    else:
        return "Same as average"

def find_closest_comps(user_coords):
    valid_data = all_data.dropna(subset=['Latitude', 'Longitude', 'Price']).copy()
    valid_data = valid_data[(valid_data['Latitude'] != 0) & (valid_data['Longitude'] != 0)]

    valid_data['distance'] = valid_data.apply(
        lambda row: round(geodesic(user_coords, (row['Latitude'], row['Longitude'])).miles, 2),
        axis=1
    )

    sorted_data = valid_data.sort_values('distance')
    comps5 = sorted_data[['Centre #', 'Latitude', 'Longitude', 'Price', 'distance']].head(5)

    avg_price = comps5['Price'].mean()

    # Comp #1
    comp1_price = comps5.iloc[0]['Price']
    if comp1_price > avg_price:
        quality1 = "Higher Quality"
    elif comp1_price < avg_price:
        quality1 = "Lesser Quality"
    else:
        quality1 = "Same Quality"
    diff1 = ((comp1_price - avg_price) / avg_price) * 100
    diff1_str = format_diff(round(diff1, 2))

    # Comp #2
    if len(comps5) > 1:
        comp2_price = comps5.iloc[1]['Price']
        if comp2_price > avg_price:
            quality2 = "Higher Quality"
        elif comp2_price < avg_price:
            quality2 = "Lesser Quality"
        else:
            quality2 = "Same Quality"
        diff2 = ((comp2_price - avg_price) / avg_price) * 100
        diff2_str = format_diff(round(diff2, 2))
    else:
        quality2 = ""
        diff2_str = ""

    comp_centres = comps5['Centre #'].head(2).tolist()
    comp_distances = [f"{d} mi" for d in comps5['distance'].head(2).tolist()]
    while len(comp_centres) < 2:
        comp_centres.append("")
        comp_distances.append("")

    return comp_centres, comp_distances, quality1, quality2, diff1_str, diff2_str

def find_online_coworking_osm(user_coords):
    lat, lon = user_coords
    overpass_url = "http://overpass-api.de/api/interpreter"
    radius = 10000  # 10 km radius
    query = f"""
    [out:json];
    node["office"="coworking"](around:{radius},{lat},{lon});
    out;
    """
    response = requests.get(overpass_url, params={'data': query})
    data = response.json()

    coworking_spaces = []
    for element in data.get('elements', []):
        name = element['tags'].get('name')
        if name:
            c_lat = element['lat']
            c_lon = element['lon']
            dist = round(geodesic(user_coords, (c_lat, c_lon)).miles, 2)
            coworking_spaces.append((name, dist))
    coworking_spaces = sorted(coworking_spaces, key=lambda x: x[1])
    return coworking_spaces[:2]

def fill_pricing_template(template_path, centre_num, centre_address, currency,
                          area_units, total_area, net_internal_area,
                          monthly_rent, rent_source,
                          service_charges, property_tax,
                          comp_centres, comp_distances,
                          quality1, quality2, diff1_str, diff2_str,
                          coworking_names, coworking_distances):
    wb = load_workbook(template_path)
    ws = wb['Centre & Market Details']

    ws['C2'] = centre_num
    ws['C3'] = centre_address
    ws['D5'] = currency
    ws['D6'] = area_units
    ws['D8'] = net_internal_area
    ws['D9'] = ""  # clear D9 as requested
    ws['D10'] = monthly_rent
    ws['D11'] = rent_source
    ws['D12'] = service_charges
    ws['D13'] = property_tax

    ws['D17'] = comp_centres[0]
    ws['E17'] = comp_centres[1]
    ws['D18'] = comp_distances[0]
    ws['E18'] = comp_distances[1]
    ws['D19'] = quality1
    ws['E19'] = quality2
    ws['D20'] = diff1_str
    ws['E20'] = diff2_str

    # --- Coworking spaces name & distance exactly as requested ---
    ws['D30'] = coworking_names[0] if len(coworking_names) > 0 else ""
    ws['E30'] = coworking_names[1] if len(coworking_names) > 1 else ""
    ws['D31'] = coworking_distances[0] if len(coworking_distances) > 0 else ""
    ws['E31'] = coworking_distances[1] if len(coworking_distances) > 1 else ""

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
        comp_centres, comp_distances, quality1, quality2, diff1_str, diff2_str = find_closest_comps(user_coords)
        coworking_spaces = find_online_coworking_osm(user_coords)

        # Prepare coworking output or blanks
        if len(coworking_spaces) == 0:
            coworking_names = ["", ""]
            coworking_distances = ["", ""]
        elif len(coworking_spaces) == 1:
            coworking_names = [coworking_spaces[0][0], ""]
            coworking_distances = [f"{coworking_spaces[0][1]} mi", ""]
        else:
            coworking_names = [coworking_spaces[0][0], coworking_spaces[1][0]]
            coworking_distances = [f"{coworking_spaces[0][1]} mi", f"{coworking_spaces[1][1]} mi"]

        # Show outputs clearly in UI
        st.markdown("### Closest Comps")
        st.write(f"Comp #1: {comp_centres[0]} at {comp_distances[0]}, Quality: {quality1} ({diff1_str})")
        st.write(f"Comp #2: {comp_centres[1]} at {comp_distances[1]}, Quality: {quality2} ({diff2_str})")

        st.markdown("### Closest Coworking Spaces")
        st.write(f"1st Closest: {coworking_names[0]} ({coworking_distances[0]})")
        st.write(f"2nd Closest: {coworking_names[1]} ({coworking_distances[1]})")

        filled_file = fill_pricing_template(
            "Pricing Template 2025.xlsx",
            centre_num, centre_address, currency,
            area_units, total_area, net_internal_area,
            monthly_rent, rent_source,
            service_charges, property_tax,
            comp_centres, comp_distances,
            quality1, quality2, diff1_str, diff2_str,
            coworking_names, coworking_distances
        )

        with open(filled_file, "rb") as f:
            st.download_button(
                label="Download Filled Pricing Template",
                data=f,
                file_name="Pricing_Template_2025_filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
