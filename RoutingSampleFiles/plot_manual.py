import pandas as pd
import folium
from pathlib import Path

# --- NEW: Define input and output paths ---
input_dir = Path(__file__).parent / 'test_data_dates'
output_dir = Path(__file__).parent / 'route_maps'

# Ensure the output directory exists
output_dir.mkdir(parents=True, exist_ok=True)

# Warehouse Coords
WAREHOUSES = {
    'JAIPUR': (26.8139, 75.5450),
    'JODHPUR': (26.1245, 73.0543),
    'DEFAULT': (26.8139, 75.5450)
}

# 1. Loop through the 5 sample files
for input_filepath in input_dir.glob('*.xlsx'):
    short_name = input_filepath.stem # e.g. '2025-06-14_Small'
    
    # Load the file and clean headers
    df = pd.read_excel(input_filepath)
    df.columns = df.columns.str.strip()
    
    # Identify columns (robust search for the spaces I added)
    lat_col = next((c for c in df.columns if c.lower() in ['lat', 'latitude', 'lat ']), 'LAT ')
    lon_col = next((c for c in df.columns if c.lower() in ['long', 'longitude', 'long ']), 'LONG')
    group_col = 'CLUBBING_OLD'
    site_col = 'eNBsiteID'
    wh_col = next((c for c in df.columns if c.lower() in ['wh', 'warehouse_name', 'wh ']), 'WH ')
    
    # Determine warehouse from first row
    wh_name = str(df.iloc[0][wh_col]).split(' ')[0].upper() if wh_col in df.columns else 'JAIPUR'
    depot_lat, depot_lon = WAREHOUSES.get(wh_name, WAREHOUSES['DEFAULT'])

    # 2. Extract Route Letter
    df['Route_Group'] = df[group_col].astype(str).str.extract(r'([A-Za-z]+)')
    
    # 3. Setup the Map
    m = folium.Map(location=[depot_lat, depot_lon], zoom_start=9)
    folium.Marker(
        [depot_lat, depot_lon],
        popup=f"WAREHOUSE: {wh_name}",
        icon=folium.Icon(color='black', icon='home', prefix='fa')
    ).add_to(m)

    colors = ['red', 'blue', 'darkgreen', 'purple', 'orange', 'darkred', 'cadetblue', 'darkblue', 'pink', 'lightgreen']
    print(f"Plotting manual routes for {short_name} from {wh_name} Hub...")

    # 4. Loop and Plot
    unique_routes = df['Route_Group'].dropna().unique()

    for idx, route in enumerate(unique_routes):
        color = colors[idx % len(colors)]
        route_data = df[df['Route_Group'] == route]
        
        # Sort by the number in clubbing (e.g. A1, A2)
        route_data = route_data.copy()
        route_data['route_num'] = route_data[group_col].str.extract(r'(\d+)').astype(float)
        route_data = route_data.sort_values(by='route_num')

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
            
        if len(route_coords) > 1:
            folium.PolyLine(route_coords, color=color, weight=4, opacity=0.8).add_to(m)

    # 5. Save
    output_filepath = output_dir / f"{short_name}_manual.html"
    m.save(output_filepath)
    print(f"Success! Saved map as '{output_filepath}'.")

print(f"Success! Saved map as '{output_filepath}'.")