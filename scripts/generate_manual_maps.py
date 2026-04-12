import pandas as pd
import os
import re
import json
import glob
import shutil

# ──────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────
WH_COORDS = {
    "JAIPUR": (26.810486, 75.496696),
    "JODHPUR": (26.148422, 73.061378),
    "DEFAULT": (26.810486, 75.496696)
}

# The root directory for organized manual data
BASE_OUTPUT_DIR = "SiteRouting/manual data"
INPUT_DIR = "SiteRouting/Processed_Inputs"

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Manual Routes: {{BATCH_NAME}}</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
        #map { height: 100vh; width: 100%; border-radius: 12px; }
        .legend { 
            position: fixed; top: 20px; right: 20px; background: white; 
            padding: 15px; border-radius: 12px; z-index: 1000; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.1); 
            font-family: 'Outfit', sans-serif;
            min-width: 200px;
        }
        .route-entry { margin-bottom: 8px; display: flex; align-items: center; gap: 8px; font-size: 14px; }
        .color-box { width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; }
        h3 { margin: 0 0 10px 0; font-size: 16px; border-bottom: 1px solid #eee; padding-bottom: 8px; }
    </style>
</head>
<body>
    <div id="map"></div>
    <div class="legend" id="legend">
        <h3>Manual Routes ({{BATCH_NAME}})</h3>
    </div>
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
        const map = L.map('map');
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png').addTo(map);

        const routesArray = {{ROUTES_JSON}};
        const colors = ['#6366f1', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#f97316'];
        const allBounds = [];
        const legend = document.getElementById('legend');

        routesArray.forEach((routeData, rIdx) => {
            const color = colors[rIdx % colors.length];
            const pathPoints = [];

            // Add Legend Entry
            const entry = document.createElement('div');
            entry.className = 'route-entry';
            entry.innerHTML = `<div class="color-box" style="background: ${color}"></div> <span>Route ${routeData.name} (${routeData.legs.length} sites)</span>`;
            legend.appendChild(entry);

            // Add Warehouse
            const origin = routeData.origin;
            pathPoints.push([origin.lat, origin.lng]);
            allBounds.push([origin.lat, origin.lng]);
            
            L.circleMarker([origin.lat, origin.lng], {
                radius: 8, fillColor: "#000", color: "#fff", weight: 2, opacity: 1, fillOpacity: 1
            }).addTo(map).bindPopup(`<strong>Warehouse</strong> (Origin for Route ${routeData.name})`);

            // Add Sites
            routeData.legs.forEach(leg => {
                const pos = [leg.lat, leg.lng];
                pathPoints.push(pos);
                allBounds.push(pos);
                L.circleMarker(pos, {
                    radius: 6, fillColor: color, color: "#fff", weight: 2, opacity: 1, fillOpacity: 0.8
                }).addTo(map).bindPopup(`<strong>Manual Route ${routeData.name}</strong><br>Stop ${leg.seq}<br>Site: ${leg.id}`);
            });

            // Draw Line
            L.polyline(pathPoints, { 
                color: color, weight: 3, opacity: 0.6, dashArray: '10, 5' 
            }).addTo(map);
        });

        if (allBounds.length > 0) {
            map.fitBounds(L.latLngBounds(allBounds), { padding: [50, 50] });
        }
    </script>
</body>
</html>
"""

def get_sort_key(s):
    match = re.search(r'(\d+)$', str(s))
    return int(match.group(1)) if match else 0

def process_file(file_path):
    # Load Data
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return None

    # Extract date from filename: Routing_Input_2026-02-06.xlsx -> 2026-02-06
    match = re.search(r'(\d{4}-\d{2}-\d{2})', os.path.basename(file_path))
    date_str = match.group(1) if match else "Unknown_Date"
    
    # Create date-wise subfolder
    date_folder = os.path.join(BASE_OUTPUT_DIR, date_str)
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)

    # Identify Columns
    lat_col = next((c for c in df.columns if c.strip().lower() in ['latitude', 'lat', 'lat ']), None)
    lng_col = next((c for c in df.columns if c.strip().lower() in ['longitude', 'lng', 'lon', 'long', 'long ']), None)
    id_col = next((c for c in df.columns if c.strip().lower() in ['site id', 'site_id', 'siteid', 'enbsiteid']), 'SiteID')
    wh_col = next((c for c in df.columns if c.strip().lower() in ['wh', 'warehouse']), None)
    club_col = 'CLUBBING'

    if not lat_col or not lng_col:
        print(f"Skipping {file_path}: Lat/Lng columns not found")
        return None

    # Parse Manual Routes into a list of route objects
    route_groups = {}
    for idx, row in df.iterrows():
        m_val = str(row.get(club_col, '')).strip()
        if not m_val or m_val.lower() in ['nan', 'none', '']: continue
        
        prefix = re.sub(r'\d+$', '', m_val)
        wh_name = str(row[wh_col]).strip().upper() if wh_col else "DEFAULT"
        if wh_name == 'JLJH' or wh_name == 'JOD': wh_name = 'JODHPUR'
        
        if prefix not in route_groups: 
            route_groups[prefix] = {"sites": [], "wh": wh_name}
        
        route_groups[prefix]["sites"].append({
            "id": str(row.get(id_col, idx)),
            "lat": float(row[lat_col]),
            "lng": float(row[lng_col]),
            "seq": get_sort_key(m_val)
        })

    routes_json_data = []
    for prefix in sorted(route_groups.keys()):
        route_info = route_groups[prefix]
        sorted_sites = sorted(route_info["sites"], key=lambda x: x['seq'])
        wh_coords = WH_COORDS.get(route_info["wh"], WH_COORDS['DEFAULT'])
        
        routes_json_data.append({
            "name": prefix,
            "origin": {"lat": wh_coords[0], "lng": wh_coords[1]},
            "legs": [{"id": s['id'], "lat": s['lat'], "lng": s['lng'], "seq": i+1} for i, s in enumerate(sorted_sites)]
        })

    if not routes_json_data:
        print(f"Skipping {file_path}: No manual routes found in CLUBBING column")
        return None

    # Copy the Excel file to the subfolder
    target_excel = os.path.join(date_folder, os.path.basename(file_path))
    shutil.copy2(file_path, target_excel)

    # Generate HTML Map in the subfolder
    html_content = HTML_TEMPLATE.replace("{{BATCH_NAME}}", date_str)
    html_content = html_content.replace("{{ROUTES_JSON}}", json.dumps(routes_json_data))
    
    map_filename = f"Manual_Routes_{date_str}.html"
    output_path = os.path.join(date_folder, map_filename)
    
    with open(output_path, "w") as f:
        f.write(html_content)
    
    return date_folder

def main():
    if not os.path.exists(BASE_OUTPUT_DIR):
        os.makedirs(BASE_OUTPUT_DIR)
    
    files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    print(f"Found {len(files)} files to process in {INPUT_DIR}...")
    
    success_count = 0
    for f in sorted(files):
        folder = process_file(f)
        if folder:
            print(f"Organized: {folder}")
            success_count += 1
    
    print(f"\nCompleted! Organized {success_count} dates in '{BASE_OUTPUT_DIR}/'")

if __name__ == "__main__":
    main()
