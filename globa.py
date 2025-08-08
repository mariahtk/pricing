import pandas as pd
import requests
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import streamlit as st
from openpyxl import load_workbook
import tempfile
import pdfplumber
import re
import os

# --- Hide Streamlit Branding and Toolbar ---
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}     /* Hide hamburger menu */
    footer {visibility: hidden;}       /* Hide footer (Made with Streamlit) */
    header {visibility: hidden;}       /* Hide top header (share, edit, github) */
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

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

# --- Geolocator ---
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
    quality1 = "Higher Quality" if comp1_price > avg_price else "Lesser Quality" if comp1_price < avg_price else "Same Quality"
    diff1 = ((comp1_price - avg_price) / avg_price) * 100
    diff1_str = format_diff(round(diff1, 2))

    # Comp #2
    if len(comps5) > 1:
        comp2_price = comps5.iloc[1]['Price']
        quality2 = "Higher Quality" if comp2_price > avg_price else "Lesser Quality" if comp2_price < avg_price else "Same Quality"
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

    return comp_centres, comp_distances, quality1, quality2, diff1_str, diff2_str, avg_price

def find_online_coworking_osm(user_coords):
    lat, lon = user_coords
    overpass_url = "http://overpass-api.de/api/interpreter"

    radius = 10000
    step = 10000
    max_radius = 100000
    coworking_spaces = []

    while radius <= max_radius:
        query = f"""
        [out:json];
        node["office"="coworking"](around:{radius},{lat},{lon});
        out;
        """
        response = requests.get(overpass_url, params={'data': query})
        data = response.json()

        new_spaces = []
        for element in data.get('elements', []):
            name = element['tags'].get('name')
            if name:
                c_lat = element['lat']
                c_lon = element['lon']
                dist = round(geodesic(user_coords, (c_lat, c_lon)).miles, 2)
                new_spaces.append((name, dist, c_lat, c_lon))

        all_spaces = coworking_spaces + new_spaces
        unique_spaces = {}
        for name, dist, c_lat, c_lon in all_spaces:
            if name not in unique_spaces or dist < unique_spaces[name][0]:
                unique_spaces[name] = (dist, c_lat, c_lon)
        coworking_spaces = sorted(
            [(name, val[0], val[1], val[2]) for name, val in unique_spaces.items()],
            key=lambda x: x[1]
        )

        if len(coworking_spaces) >= 2:
            break
        radius += step

    if len(coworking_spaces) == 0:
        coworking_spaces = [("No coworking space found nearby", 0, None, None)]
    return coworking_spaces[:2]

city_rent_lookup_sqft = {
    "New York": 70, "San Francisco": 65, "Chicago": 40, "Los Angeles": 45,
    "Seattle": 50, "Boston": 55, "Austin": 35, "Denver": 30, "Miami": 30,
    "Washington": 50, "Atlanta": 28, "Dallas": 32, "Houston": 28,
    "Toronto": 50, "Vancouver": 45, "Montreal": 35, "Calgary": 30,
    "Ottawa": 32, "Edmonton": 28, "Winnipeg": 25, "Quebec": 25,
}
city_rent_lookup_sqm = {city: val * 10.7639 for city, val in city_rent_lookup_sqft.items()}

def get_city_from_coords(lat, lon):
    location = geolocator.reverse((lat, lon), exactly_one=True)
    if location:
        addr = location.raw.get('address', {})
        city = (addr.get('city') or addr.get('town') or addr.get('municipality') or
                addr.get('village') or addr.get('hamlet'))
        if city:
            return city
        state = addr.get('state')
        if state:
            return state
    return None

def estimate_coworking_price(lat, lon, area_units):
    fixed_office_size = 150
    city = get_city_from_coords(lat, lon)
    if not city:
        price_per_unit = 5 if area_units == "SqFt" else 55
    else:
        price_per_unit = city_rent_lookup_sqft.get(city, 5) if area_units == "SqFt" else city_rent_lookup_sqm.get(city, 55)
    estimated_price = price_per_unit * fixed_office_size
    return round(min(estimated_price, 2000), 2)

# --- Safe numeric conversion ---
def safe_to_float(val):
    try:
        return float(val.replace(",", "")) if val and val != "." else 0
    except:
        return 0

# --- Financial Model Extraction ---
def extract_from_excel(uploaded_file):
    try:
        wb = load_workbook(uploaded_file, data_only=True)
        ws = wb["10Yr Model"] if "10Yr Model" in wb.sheetnames else wb.active
        text = " ".join([str(cell.value) for row in ws.iter_rows() for cell in row if cell.value])
        currency = "USD" if "USD" in text else "CAD" if "CAD" in text else "USD"

        gross_area = market_rent = cashflow = None
        for row in ws.iter_rows(values_only=True):
            row_text = [str(x).strip().lower() if x else "" for x in row]
            if "gross area (sqft)" in " ".join(row_text):
                gross_area = next((x for x in row if isinstance(x, (int, float))), 0)
            if "market rent value" in " ".join(row_text):
                market_rent = next((x for x in row if isinstance(x, (int, float))), 0)
            if "net partner cashflow" in " ".join(row_text) and "year 1" in " ".join(row_text):
                cashflow = next((x for x in row if isinstance(x, (int, float))), 0)

        net_internal_area = gross_area * 0.5 if gross_area else 0
        monthly_cashflow = cashflow / 12 if cashflow else 0
        return currency, gross_area or 0, net_internal_area, market_rent or 0, monthly_cashflow or 0
    except Exception as e:
        st.warning(f"Could not parse Excel model: {e}")
        return None

def extract_from_pdf(uploaded_file):
    try:
        text = ""
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        currency = "USD" if "USD" in text else "CAD" if "CAD" in text else "USD"

        # Patterns for total area contracted (look for "Rentable Area" or "Total Area Contracted")
        rentable_area_match = re.search(r"Rentable Area.*?sqft.*?([\d,\.]+)", text, re.IGNORECASE)
        total_area_contracted_match = re.search(r"Total Area Contracted.*?([\d,\.]+)", text, re.IGNORECASE)
        gross_area_match = re.search(r"Gross Area.*?sqft.*?([\d,\.]+)", text, re.IGNORECASE)

        # Pick first found in order: Total Area Contracted, Rentable Area, Gross Area
        total_area = None
        for match in [total_area_contracted_match, rentable_area_match, gross_area_match]:
            if match:
                total_area = safe_to_float(match.group(1))
                break
        if total_area is None:
            total_area = 0

        # Patterns for net internal area (look for "Sellable Office Area")
        sellable_area_match = re.search(r"Sellable Office Area.*?sqft.*?([\d,\.]+)", text, re.IGNORECASE)
        net_internal_area = safe_to_float(sellable_area_match.group(1)) if sellable_area_match else total_area * 0.5

        # Patterns for market rent: check "Market Rent Value" or "Headline Rent (as reviewed by partner)"
        market_rent_match = re.search(r"Market Rent Value.*?([\d,\.]+)", text, re.IGNORECASE)
        headline_rent_match = re.search(r"Headline Rent \(as reviewed by partner\).*?([\d,\.]+)", text, re.IGNORECASE)
        market_rent = None
        for match in [market_rent_match, headline_rent_match]:
            if match:
                market_rent = safe_to_float(match.group(1))
                break
        if market_rent is None:
            market_rent = 0

        # Pattern for total cash flow: "Net Partner Cashflow" and "Year 1"
        cashflow_match = re.search(r"Net Partner Cashflow.*?Year 1.*?([\d,\.]+)", text, re.IGNORECASE)
        cashflow = safe_to_float(cashflow_match.group(1)) if cashflow_match else 0
        monthly_cashflow = cashflow / 12 if cashflow else 0

        return currency, total_area, net_internal_area, market_rent, monthly_cashflow
    except Exception as e:
        st.warning(f"Could not parse PDF model: {e}")
        return None

# --- Pricing Template Filler ---
def fill_pricing_template(template_path, centre_num, centre_address, currency,
                          area_units, total_area, net_internal_area,
                          monthly_rent, rent_source,
                          service_charges, property_tax,
                          comp_centres, comp_distances,
                          quality1, quality2, diff1_str, diff2_str,
                          coworking_names, coworking_distances,
                          coworking_price1, coworking_price2,
                          total_cash_flow):
    if not os.path.exists(template_path):
        st.error(f"Template file not found: {template_path}")
        return None

    wb = load_workbook(template_path)
    ws = wb['Centre & Market Details']

    ws['C2'] = centre_num
    ws['C3'] = centre_address
    ws['D5'] = currency
    ws['D6'] = area_units
    ws['D7'] = total_area                # Total Area Contracted here
    ws['D8'] = net_internal_area
    ws['D9'] = ""
    ws['D10'] = monthly_rent             # Market Rent here
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

    ws['D30'] = coworking_names[0] if len(coworking_names) > 0 else ""
    ws['E30'] = coworking_names[1] if len(coworking_names) > 1 else ""
    ws['D31'] = coworking_distances[0] if len(coworking_distances) > 0 else ""
    ws['E31'] = coworking_distances[1] if len(coworking_distances) > 1 else ""
    ws['D33'] = coworking_price1 if coworking_price1 is not None else ""
    ws['E33'] = coworking_price2 if coworking_price2 is not None else ""
    ws['D35'] = total_cash_flow          # Total Monthly Expected Cash Flow Maturity

    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp_file.name)
    return tmp_file.name

# --- Streamlit UI ---
st.title("Pricing Template 2025 Filler")
uploaded_model = st.file_uploader("Upload Financial Model (PDF)", type=["xlsx", "xls", "pdf"])

# Variables for extracted/default values
currency = None
total_area = 0.0
net_internal_area = 0.0
monthly_rent = 0.0
total_cash_flow = 0.0

if uploaded_model:
    if uploaded_model.name.endswith(".pdf"):
        parsed = extract_from_pdf(uploaded_model)
    else:
        parsed = extract_from_excel(uploaded_model)

    if parsed:
        currency, total_area, net_internal_area, monthly_rent, total_cash_flow = parsed
        st.success("Values auto-extracted from model.")
    else:
        currency = st.selectbox("Pricing Currency", ["USD", "CAD"])
        total_area = st.number_input("Total Area Contracted", min_value=0.0, format="%.2f")
        net_internal_area = st.number_input("Net Internal Area", min_value=0.0, format="%.2f")
        monthly_rent = st.number_input("Monthly Market Rent", min_value=0.0, format="%.2f")
        total_cash_flow = st.number_input("Total Monthly Expected Cash Flow Maturity", min_value=0.0, format="%.2f")
else:
    currency = st.selectbox("Pricing Currency", ["USD", "CAD"])
    total_area = st.number_input("Total Area Contracted", min_value=0.0, format="%.2f")
    net_internal_area = st.number_input("Net Internal Area", min_value=0.0, format="%.2f")
    monthly_rent = st.number_input("Monthly Market Rent", min_value=0.0, format="%.2f")
    total_cash_flow = st.number_input("Total Monthly Expected Cash Flow Maturity", min_value=0.0, format="%.2f")

centre_num = st.text_input("Centre #")
centre_address = st.text_input("Centre Address")
area_units = st.selectbox("Area Units", ["SqM", "SqFt"])
rent_source = st.selectbox("Source of Market Rent", ["LL or Partner Provided", "Broker Provided or Market Report", "Benchmarked from similar centre"])
service_charges = st.number_input("Service Charges", min_value=0.0, format="%.2f")
property_tax = st.number_input("Property Tax", min_value=0.0, format="%.2f")

# Force monthly_rent to float for number_input default to avoid Streamlit errors
try:
    monthly_rent_val = float(monthly_rent)
except:
    monthly_rent_val = 0.0

monthly_rent_override = st.number_input("Override Monthly Market Rent", value=monthly_rent_val, min_value=0.0, format="%.2f")

# Show comps and coworking spaces info if centre_address entered
if centre_address:
    user_coords = get_coords(centre_address)
    if user_coords:
        comp_centres, comp_distances, quality1, quality2, diff1_str, diff2_str, avg_price = find_closest_comps(user_coords)
        st.markdown("### Closest Comps")
        st.write(f"**Comp #1:** {comp_centres[0]} — {comp_distances[0]} — Quality: {quality1} — {diff1_str}")
        if comp_centres[1]:
            st.write(f"**Comp #2:** {comp_centres[1]} — {comp_distances[1]} — Quality: {quality2} — {diff2_str}")

        coworking_spaces = find_online_coworking_osm(user_coords)
        coworking_names = [c[0] for c in coworking_spaces]
        coworking_distances = [f"{c[1]} mi" for c in coworking_spaces]
        coworking_price1 = estimate_coworking_price(coworking_spaces[0][2], coworking_spaces[0][3], area_units) if len(coworking_spaces) > 0 else None
        coworking_price2 = estimate_coworking_price(coworking_spaces[1][2], coworking_spaces[1][3], area_units) if len(coworking_spaces) > 1 else None

        st.markdown("### Nearby Coworking Spaces")
        for i, (name, dist) in enumerate(zip(coworking_names, coworking_distances)):
            price = coworking_price1 if i == 0 else coworking_price2 if i == 1 else None
            price_str = f"${price}" if price else "N/A"
            st.write(f"**{name}** — {dist} — Estimated Price: {price_str}")
    else:
        st.warning("Could not geocode the given address. Please enter a valid address to see comps and coworking info.")

if st.button("Generate Pricing Template"):
    if not centre_num or not centre_address:
        st.error("Please enter Centre # and Centre Address")
    else:
        # Use override rent if provided
        final_monthly_rent = monthly_rent_override if monthly_rent_override > 0 else monthly_rent_val

        file_path = fill_pricing_template(
            "Pricing_Template_2025.xlsx", centre_num, centre_address, currency,
            area_units, total_area, net_internal_area,
            final_monthly_rent, rent_source,
            service_charges, property_tax,
            comp_centres if centre_address else ["", ""],
            comp_distances if centre_address else ["", ""],
            quality1 if centre_address else "",
            quality2 if centre_address else "",
            diff1_str if centre_address else "",
            diff2_str if centre_address else "",
            coworking_names if centre_address else ["", ""],
            coworking_distances if centre_address else ["", ""],
            coworking_price1 if centre_address else None,
            coworking_price2 if centre_address else None,
            total_cash_flow
        )

        if file_path:
            with open(file_path, "rb") as f:
                st.download_button(
                    label="Download Filled Pricing Template",
                    data=f,
                    file_name="Pricing_Template_Filled_2025.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            os.unlink(file_path)
