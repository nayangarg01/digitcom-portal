import pandas as pd
import os
import re
import json

# ──────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────
WH_COORDS = {
    "JAIPUR": (26.810486, 75.496696),
    "JODHPUR": (26.148422, 73.061378),
    "DEFAULT": (26.810486, 75.496696)
}

OUTPUT_DIR = "SiteRouting/Manual_Route_Maps"

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Manual Route: {{ROUTE_NAME}}</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <style>
        #map { height: 100vh; width: 100%; }
        .legend { position: fixed; bottom: 20px; right: 20px; background: white; padding: 10px; border-radius: 8px; z-index: 1000; box-shadow: 0 0 10px rgba(0,0,0,0.2); }
    </style>
</head>
<body>
    <div id="map"></div>
    <div class="legend"><strong>Manual Route: {{ROUTE_NAME}}</strong></div>
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script>
        const map = L.map('map');
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png').addTo(map);

        const routeData = {{ROUTE_JSON}};
        const pathPoints = [];

        // Add Warehouse
        const origin = routeData.origin;
        pathPoints.push([origin.lat, origin.lng]);
        L.circleMarker([origin.lat, origin.lng], {
            radius: 10, fillColor: "#000", color: "#fff", weight: 2, opacity: 1, fillOpacity: 1
        }).addTo(map).bindPopup("<strong>Warehouse</strong>");

        // Add Sites
        routeData.legs.forEach(leg => {
            const pos = [leg.lat, leg.lng];
            pathPoints.push(pos);
            L.circleMarker(pos, {
                radius: 8, fillColor: "#ff4757", color: "#fff", weight: 2, opacity: 1, fillOpacity: 0.8
            }).addTo(map).bindPopup(`<strong>Stop ${leg.seq}</strong><br>Site: ${leg.id}`);
        });

        // Draw Line
        L.polyline(pathPoints, { color: "#ff4757", weight: 4, opacity: 0.7 }).addTo(map);

        const group = new L.featureGroup(pathPoints.map(p => L.marker(p)));
        map.fitBounds(group.getBounds(), { padding: [50, 50] });
    </script>
</body>
</html>
"""

def generate_maps(file_path):
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # Load Data
    df = pd.read_excel(file_path)
    
    # Identify Columns
    lat_col = next((c for c in df.columns if c.strip().lower() in ['latitude', 'lat', 'lat ']), None)
    lng_col = next((c for c in df.columns if c.strip().lower() in ['longitude', 'lng', 'lon', 'long', 'long ']), None)
    id_col = next((c for c in df.columns if c.strip().lower() in ['site id', 'site_id', 'siteid', 'enbsiteid']), 'SiteID')
    wh_col = next((c for c in df.columns if c.strip().lower() in ['wh', 'warehouse']), None)
    club_col = 'CLUBBING'

    if not lat_col or not lng_col:
        print("Error: Lat/Lng columns not found")
        return

    # Parse Manual Routes
    def get_sort_key(s):
        match = re.search(r'(\d+)$', str(s))
        return int(match.group(1)) if match else 0

    route_groups = {}
    for idx, row in df.iterrows():
        m_val = str(row.get(club_col, '')).strip()
        if not m_val or m_val.lower() in ['nan', 'none', '']: continue
        
        prefix = re.sub(r'\d+$', '', m_val)
        wh_name = str(row[wh_col]).strip().upper() if wh_col else "DEFAULT"
        if wh_name == 'JLJH' or wh_name == 'JOD': wh_name = 'JODHPUR'
        
        m_key = f"{prefix}_{wh_name}"
        if m_key not in route_groups: route_groups[m_key] = []
        
        route_groups[m_key].append({
            "id": str(row.get(id_col, idx)),
            "lat": float(row[lat_col]),
            "lng": float(row[lng_col]),
            "seq": get_sort_key(m_val)
        })

    # Generate HTML files
    for m_key, sites in route_groups.items():
        sorted_sites = sorted(sites, key=lambda x: x['seq'])
        prefix, wh_name = m_key.split('_')
        wh_coords = WH_COORDS.get(wh_name, WH_COORDS['DEFAULT'])
        
        route_info = {
            "name": prefix,
            "origin": {"lat": wh_coords[0], "lng": wh_coords[1]},
            "legs": [{"id": s['id'], "lat": s['lat'], "lng": s['lng'], "seq": i+1} for i, s in enumerate(sorted_sites)]
        }
        
        html_content = HTML_TEMPLATE.replace("{{ROUTE_NAME}}", prefix)
        html_content = html_content.replace("{{ROUTE_JSON}}", json.dumps(route_info))
        
        safe_prefix = "".join([c for c in prefix if c.isalnum() or c in (' ', '.', '_', '-')]).strip()
        filename = f"Manual_Route_{safe_prefix}.html"
        output_path = os.path.join(OUTPUT_DIR, filename)
        
        with open(output_path, "w") as f:
            f.write(html_content)
        print(f"Generated: {output_path}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        generate_maps(sys.argv[1])
    else:
        print("Usage: python3 generate_manual_maps.py [INPUT_EXCEL_FILE]")
