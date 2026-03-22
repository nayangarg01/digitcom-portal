import pandas as pd
import folium

# 1. Load the file and clean headers
df = pd.read_excel('Optimized_DC011.xlsx')
df.columns = df.columns.str.strip()

lat_col = 'LAT'
lon_col = 'LONG'
group_col = 'CLUBBING' 
site_col = 'eNBsiteID'

# 2. Extract Route Letter
df['Route_Group'] = df[group_col].astype(str).str[0]

# --- NEW: Define the Jaipur Bagru Warehouse ---
depot_lat = 26.8139
depot_lon = 75.5450

# 3. Setup the Map (Centered directly on the warehouse)
m = folium.Map(location=[depot_lat, depot_lon], zoom_start=7)

# --- NEW: Add a prominent marker for the Warehouse ---
folium.Marker(
    [depot_lat, depot_lon],
    popup="WAREHOUSE: Jaipur - Bagru",
    icon=folium.Icon(color='black', icon='home', prefix='fa')
).add_to(m)

colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue', 'gray', 'pink', 'lightgreen']

print("Plotting manual routes from the Warehouse Hub...")

# 4. Loop and Plot
unique_routes = df['Route_Group'].dropna().unique()

for idx, route in enumerate(unique_routes):
    if route == 'n': 
        continue
        
    color = colors[idx % len(colors)]
    route_data = df[df['Route_Group'] == route].sort_values(by=group_col)
    
    # --- NEW: Force every route to start exactly at the warehouse ---
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

m.save("manual_routes_hub_spoke.html")
print("Success! Open 'manual_routes_hub_spoke.html' in your browser.")