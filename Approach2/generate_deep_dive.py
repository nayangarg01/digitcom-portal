import pandas as pd
from pathlib import Path
import os
import folium
import math
import re

# --- CONFIG ---
KMS_FILE = Path("Routing/KMS.xlsx")
DPR_FILE = Path("Routing/DPR_common.xlsx")
DEEP_DIVE_DIR = Path("Approach2/DC012_2025-04-11")
TARGET_DC = "DC012"
TARGET_DATE = "2025-04-11"

WAREHOUSE_MAP = {
    'JAIPUR': (26.9124, 75.7873),
    'JODHPUR': (26.2389, 73.0243),
    'DEFAULT': (26.9124, 75.7873)
}

def haversine(p1, p2):
    R = 6371
    lat1, lon1 = math.radians(p1[0]), math.radians(p1[1])
    lat2, lon2 = math.radians(p2[0]), math.radians(p2[1])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

import itertools

def run_routing(warehouse_coords, cluster):
    unvisited = list(cluster)
    routes = []
    while unvisited:
        # 1. Start with the 'Seed Outlier' (furthest from warehouse)
        seed = max(unvisited, key=lambda s: haversine(warehouse_coords, s['coords']))
        unvisited.remove(seed)
        
        # 2. Find up to 2 closest neighbors for this specific outlier
        clump = [seed]
        for _ in range(2):
            if not unvisited: break
            nearest = min(unvisited, key=lambda s: haversine(seed['coords'], s['coords']))
            clump.append(nearest)
            unvisited.remove(nearest)
            
        # 3. Find the optimal visit order for this specific 3-site clump
        best_p = None; min_d = float('inf')
        for p in itertools.permutations(clump):
            d = haversine(warehouse_coords, p[0]['coords'])
            if len(p) > 1: d += haversine(p[0]['coords'], p[1]['coords'])
            if len(p) > 2: d += haversine(p[1]['coords'], p[2]['coords'])
            if d < min_d: min_d = d; best_p = p
            
        # 4. Generate the path legs
        route_legs = []
        cp = warehouse_coords
        for s in best_p:
            dist_leg = haversine(cp, s['coords'])
            route_legs.append({"site": s, "dist": round(dist_leg, 2)})
            cp = s['coords']
        routes.append(route_legs)
    return routes

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def plot_map(warehouse_coords, routes, filename, title):
    m = folium.Map(location=[warehouse_coords[0], warehouse_coords[1]], zoom_start=9)
    folium.Marker(warehouse_coords, popup="WAREHOUSE", icon=folium.Icon(color='black', icon='home')).add_to(m)
    colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue']
    for idx, r in enumerate(routes):
        c = colors[idx % len(colors)]
        coords = [warehouse_coords]
        for leg in r:
            coords.append(leg['site']['coords'])
            folium.Marker(leg['site']['coords'], popup=leg['site']['id'], icon=folium.Icon(color=c)).add_to(m)
        folium.PolyLine(coords, color=c, weight=4).add_to(m)
    m.save(filename)

def main():
    os.makedirs(DEEP_DIVE_DIR, exist_ok=True)
    df = pd.read_excel(KMS_FILE).fillna("")
    df.columns = df.columns.str.strip()
    
    # Filter for target
    mask = (df['BILLING FILE'] == TARGET_DC) & (df['MIN DATE'].astype(str).str.contains(TARGET_DATE))
    day_df = df[mask].copy()
    
    id_col, lat_col, lon_col, club_col, wh_col = 'eNBsiteID', 'LAT', 'LONG', 'CLUBBING', 'WAREHOUSE'
    
    # 1. Data Sheet (No AKTBC/Clubbing)
    data_cols = [id_col, 'BILLING FILE', 'MIN DATE', 'JC', lat_col, lon_col, wh_col]
    day_df[data_cols].to_excel(DEEP_DIVE_DIR / f"{TARGET_DC}_{TARGET_DATE}_data.xlsx", index=False)
    
    # 2. Routing Comparison
    wh_name = str(day_df.iloc[0][wh_col]).split(' ')[0].upper()
    wh_coords = WAREHOUSE_MAP.get(wh_name, WAREHOUSE_MAP['DEFAULT'])
    
    # Standard Hex Colors matching maps (roughly)
    # red, blue, green, purple, orange, darkred, cadetblue
    HEX_COLORS = ['#FFCCCC', '#CCCCFF', '#CCFFCC', '#E5CCFF', '#FFE5CC', '#FFB2B2', '#CCFFFF']

    # Reconstruct Manual
    manual_routes = []
    hist_groups = {}
    for _, row in day_df.iterrows():
        c = str(row[club_col]).strip().upper(); key = c[0] if c else "X"
        if key not in hist_groups: hist_groups[key] = []
        hist_groups[key].append({"id":row[id_col],"coords":(float(row[lat_col]),float(row[lon_col])),"club_old":row[club_col]})
    
    for k in sorted(hist_groups.keys()):
        s_list = hist_groups[k]
        s_list.sort(key=lambda x: natural_sort_key(x['club_old']))
        route = []; cp = wh_coords
        for s in s_list:
            d = haversine(cp, s['coords'])
            route.append({"site":s, "dist":round(d,2), "club":s['club_old']})
            cp = s['coords']
        manual_routes.append(route)

    # Automated
    sites = [{"id":r[id_col],"coords":(float(r[lat_col]),float(r[lon_col]))} for _,r in day_df.iterrows() if r[lat_col] != 0]
    auto_routes = run_routing(wh_coords, sites)
    for r_idx, r in enumerate(auto_routes):
        for s_idx, leg in enumerate(r):
            leg['club'] = f"R{r_idx+1}-S{s_idx+1}"

    # Write Routing Excel
    with pd.ExcelWriter(DEEP_DIVE_DIR / f"{TARGET_DC}_{TARGET_DATE}_routing.xlsx", engine='xlsxwriter') as writer:
        wb = writer.book
        title_fmt = wb.add_format({'bold':True,'align':'center','bg_color':'#2F5597','font_color':'white','border':1})
        
        for name, routes in [("Manual", manual_routes), ("Automated", auto_routes)]:
            ws = wb.add_worksheet(name)
            curr = 0
            for r_idx, r in enumerate(routes):
                # Apply color format for this route
                color = HEX_COLORS[r_idx % len(HEX_COLORS)]
                route_fmt = wb.add_format({'bg_color': color, 'border': 1, 'align': 'center'})
                header_fmt = wb.add_format({'bold': True, 'bg_color': color, 'border': 1, 'align': 'center'})
                
                ws.merge_range(curr, 0, curr, 3, f"ROUTE {r_idx+1}", title_fmt); curr += 1
                cols = ["SITE ID", "SEQ", "LEG DIST", "CLUB LABEL"]
                for c, h in enumerate(cols): ws.write(curr, c, h, header_fmt)
                curr += 1
                for s_idx, leg in enumerate(r):
                    ws.write(curr, 0, leg['site']['id'], route_fmt)
                    ws.write(curr, 1, s_idx + 1, route_fmt)
                    ws.write(curr, 2, leg['dist'], route_fmt)
                    ws.write(curr, 3, leg['club'], route_fmt)
                    curr += 1
                curr += 2
            ws.set_column(0,0,25); ws.set_column(1,3,12)

    plot_map(wh_coords, manual_routes, DEEP_DIVE_DIR / f"{TARGET_DATE}_manual.html", "Manual")
    plot_map(wh_coords, auto_routes, DEEP_DIVE_DIR / f"{TARGET_DATE}_automated.html", "Automated")
    print(f"Deep dive created in {DEEP_DIVE_DIR}")

if __name__ == "__main__":
    main()
