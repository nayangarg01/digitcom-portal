import pandas as pd
import folium
from pathlib import Path

# --- NEW: Define input and output paths ---
input_filepath = 'test_data/DC087.xlsx'
output_dir = Path('route_maps')

# Ensure the output directory exists (creates it if you haven't already)
output_dir.mkdir(parents=True, exist_ok=True)

# --- NEW: Extract the number and format the short name ---
file_stem = Path(input_filepath).stem # Gets 'DC087'
# Filters out letters and converts to int to drop the leading zero (087 -> 87)
file_number = int("".join(filter(str.isdigit, file_stem))) 
short_name = f"RM{file_number}"

# 1. Load the file and clean headers
df = pd.read_excel(input_filepath)
df.columns = df.columns.str.strip()

lat_col = 'LAT'
lon_col = 'LONG'
group_col = 'CLUBBING' 
site_col = 'eNBsiteID'

# 2. Extract Route Letter
df['Route_Group'] = df[group_col].astype(str).str[0]

# Define the Jaipur Bagru Warehouse
depot_lat = 26.8139
depot_lon = 75.5450

# 3. Setup the Map (Centered directly on the warehouse)
m = folium.Map(location=[depot_lat, depot_lon], zoom_start=7)

# Add a prominent marker for the Warehouse
folium.Marker(
    [depot_lat, depot_lon],
    popup="WAREHOUSE: Jaipur - Bagru",
    icon=folium.Icon(color='black', icon='home', prefix='fa')
).add_to(m)

colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue', 'gray', 'pink', 'lightgreen']

print(f"Plotting manual routes for {short_name} from the Warehouse Hub...")

# 4. Loop and Plot
unique_routes = df['Route_Group'].dropna().unique()

for idx, route in enumerate(unique_routes):
    if route == 'n': 
        continue
        
    color = colors[idx % len(colors)]
    route_data = df[df['Route_Group'] == route].sort_values(by=group_col)
    
    # Force every route to start exactly at the warehouse
    route_coords = [(depot_lat, depot_lon)]
    
    for _, row in route_data.iterrows():
        lat, lon = row[lat_col], row[lon_col]
        
        if pd.notna(lat) and pd.notna(lon):
            route_coords.append((lat, lon))
            folium.Marker(
                [lat, lon],
                popup=f"{row[site_col]} ({row[group_col]})",
                icon=folium.Icon(color=color)
            ).add_to(m)
        
    # Draw the line if we have the warehouse + at least 1 site
    if len(route_coords) > 1:
        folium.PolyLine(route_coords, color=color, weight=4, opacity=0.8).add_to(m)

# --- NEW: Save to the specific folder with the short name ---
output_filepath = output_dir / f"{short_name}.html"
m.save(output_filepath)

print(f"Success! Saved map as '{output_filepath}'.")