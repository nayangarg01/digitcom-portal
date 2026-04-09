import pandas as pd
import numpy as np
import googlemaps
import folium
import math
import itertools
import argparse
import sys
import os
import re

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
INPUT_FILE  = "grouped_by_date.xlsx"
OUTPUT_FILE = "routed_by_date.xlsx"
MAPS_DIR    = "Maps"

WH_COORDS_FALLBACK = {
    "JAIPUR":  (26.810486, 75.496696),
    "JODHPUR": (26.148422, 73.061378),
}

HEX_COLORS = ['#FFCCCC', '#CCCCFF', '#CCFFCC', '#E5CCFF', '#FFE5CC', '#FFB2B2', '#CCFFFF']
ROUTE_COLORS = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue']

# ──────────────────────────────────────────────
# ALGORITHM / DISTANCE LOGIC
# ──────────────────────────────────────────────
def haversine(p1, p2):
    R = 6371
    lat1, lon1 = math.radians(p1[0]), math.radians(p1[1])
    lat2, lon2 = math.radians(p2[0]), math.radians(p2[1])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

def get_api_driving_distance(gmaps, origin, dest):
    """
    Query Google Directions API for the shortest driving distance between two coordinates.
    """
    try:
        routes = gmaps.directions(origin, dest, mode='driving', alternatives=True)
        if routes:
            def total_dist(r):
                return sum(leg['distance']['value'] for leg in r['legs'])
            shortest_route = min(routes, key=total_dist)
            dist_m = total_dist(shortest_route)
            return round(dist_m / 1000.0, 2)
    except Exception as e:
        print(f"      [API Error] Could not calculate {origin} -> {dest}: {e}")
    
    # Fallback to straight line (haversine) if API fails
    return round(haversine(origin, dest), 2)

def optimize_segment(warehouse_coords, cluster):
    best_p = None; min_d = float('inf')
    for p in itertools.permutations(cluster):
        d = haversine(warehouse_coords, p[0]['coords'])
        if len(p) > 1: d += haversine(p[0]['coords'], p[1]['coords'])
        if len(p) > 2: d += haversine(p[1]['coords'], p[2]['coords'])
        if d < min_d: min_d = d; best_p = p
    return best_p

def segment_to_legs(warehouse_coords, segment):
    # This prepares the initial leg structure. Distance will be updated by API later.
    route_legs = []; cp = warehouse_coords
    for s in segment:
        d = haversine(cp, s['coords'])
        route_legs.append({"site": s, "haversine_dist": round(d, 2), "api_dist": 0.0})
        cp = s['coords']
    return route_legs

def run_routing(warehouse_coords, cluster):
    # 1. Group by JC
    jc_groups = {}
    for s in cluster:
        row = s.get('row', {})
        jc = str(row.get('JC NAME', '')).strip().upper()
        if jc not in jc_groups: jc_groups[jc] = []
        jc_groups[jc].append(s)
    
    # 2. Transition-Aware Metric: Find distance to nearest site in a DIFFERENT JC
    for s in cluster:
        my_jc = str(s.get('row', {}).get('JC NAME', '')).strip().upper()
        others = [o for o in cluster if str(o.get('row', {}).get('JC NAME', '')).strip().upper() != my_jc]
        if others:
            s['trans_dist'] = min(haversine(s['coords'], o['coords']) for o in others)
        else:
            s['trans_dist'] = 999999.0 # Effectively isolated
            
    final_routes = []; mixer_pool = []
    
    # Tier 1: Intra-Town Triplets (Prioritize isolated core sites)
    for jc, sites in jc_groups.items():
        # Sort DESC: isolated core sites first, bridges (small trans_dist) last
        unvisited = sorted(sites, key=lambda x: x['trans_dist'], reverse=True)
        while len(unvisited) >= 3:
            seed = unvisited.pop(0) # Pull most isolated core site
            clump = [seed]
            for _ in range(2):
                nearest = min(unvisited, key=lambda s: haversine(seed['coords'], s['coords']))
                clump.append(nearest)
                unvisited.remove(nearest)
            best_p = optimize_segment(warehouse_coords, clump)
            final_routes.append(segment_to_legs(warehouse_coords, best_p))
        mixer_pool.extend(unvisited) # Residues (the 'Bridges') go to Mixer
    
    # Tier 2: Global Mixer (Consolidate the bridges)
    unvisited_mixer = list(mixer_pool)
    while unvisited_mixer:
        # Mixer still uses Outlier-First logic (furthest outlier first)
        seed = max(unvisited_mixer, key=lambda s: haversine(warehouse_coords, s['coords']))
        unvisited_mixer.remove(seed)
        clump = [seed]
        for _ in range(2):
            if not unvisited_mixer: break
            nearest = min(unvisited_mixer, key=lambda s: haversine(seed['coords'], s['coords']))
            clump.append(nearest)
            unvisited_mixer.remove(nearest)
        best_p = optimize_segment(warehouse_coords, clump)
        final_routes.append(segment_to_legs(warehouse_coords, best_p))
        
    return final_routes

# ──────────────────────────────────────────────
# MAPPING LOGIC
# ──────────────────────────────────────────────
def plot_map(warehouse_coords, routes, output_file, wh_name, title):
    m = folium.Map(location=[warehouse_coords[0], warehouse_coords[1]], zoom_start=9)
    folium.Marker(warehouse_coords, popup=f"WAREHOUSE: {wh_name}", icon=folium.Icon(color='black', icon='home')).add_to(m)
    
    for r_idx, route in enumerate(routes):
        color = ROUTE_COLORS[r_idx % len(ROUTE_COLORS)]
        coords = [warehouse_coords]
        for leg in route:
            s = leg['site']; lat, lon = s['coords']
            coords.append((lat, lon))
            pop = f"Site: {s['id']}<br>Route: {leg.get('club_new')}<br>AKTBC (API): {leg.get('api_dist')} km"
            folium.Marker((lat, lon), popup=pop, icon=folium.Icon(color=color)).add_to(m)
        if len(coords) > 1:
            folium.PolyLine(coords, color=color, weight=4, opacity=0.8).add_to(m)
            
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    m.save(output_file)

# ──────────────────────────────────────────────
# MAIN PIPELINE
# ──────────────────────────────────────────────
def load_wh_coords(excel_path):
    try:
        # We try to read 'wh location' from km required 2 if needed, otherwise fallback.
        df = pd.read_excel("km required 2.xlsx", sheet_name='wh location', header=None)
        coords = {}
        for _, row in df.iterrows():
            name = str(row[0]).strip().upper()
            coords[name] = (float(row[1]), float(row[2]))
        return coords
    except:
        return WH_COORDS_FALLBACK

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def build_excel_formats(wb):
    return {
        'title': wb.add_format({'bold':True,'font_size':14,'align':'center','bg_color':'#2F5597','font_color':'white','border':1}),
        'head': wb.add_format({'bold':True,'align':'center','bg_color':'#D9E1F2','border':1}),
        'data': wb.add_format({'align':'center','border':1})
    }

def process_pipeline(api_key, input_file, output_file):
    if not os.path.exists(input_file):
        print(f"Error: Could not find {input_file}")
        sys.exit(1)
        
    print(f"Loading data from {input_file}...")
    xl = pd.ExcelFile(input_file)
    sheets = xl.sheet_names
    
    gmaps = googlemaps.Client(key=api_key, timeout=10)
    wh_coords_map = load_wh_coords("")
    
    os.makedirs(MAPS_DIR, exist_ok=True)
    
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        wb = writer.book
        fmts = build_excel_formats(wb)
        
        for date_str in sheets:
            print(f"\n[=>] Processing Date Sheet: {date_str}")
            df_date = pd.read_excel(xl, sheet_name=date_str).fillna("")
            if df_date.empty: continue
            
            is_b6 = "_A6+B6" in date_str
            
            all_auto_legs = []
            auto_for_map = []
            
            # Route logic per CMP (District)
            for cmp_name, cmp_df in df_date.groupby('CMP'):
                if cmp_name == "" or pd.isna(cmp_name): cmp_name = "UNKNOWN"
                
                # Assume all sites in CMP group use the same warehouse (as mapped in sheet)
                wh_name = str(cmp_df.iloc[0].get('WH', 'DEFAULT')).strip().upper()
                wh_coords = wh_coords_map.get(wh_name, wh_coords_map.get('JAIPUR'))
                
                # Extract valid sites
                sites_isolated = []
                sites_for_triplets = []
                
                for _, row in cmp_df.iterrows():
                    try:
                        lat, lon = float(row['LATITUDE']), float(row['LONGITUDE'])
                        if math.isnan(lat) or lat == 0: continue
                        
                        dist_val = row.get('Distance from WH (km)', 0)
                        try:
                            dist_wh = float(dist_val)
                        except:
                            dist_wh = haversine(wh_coords, (lat, lon))
                            
                        site_data = {"id": row['SITE ID'], "coords": (lat, lon), "row": row}
                        if is_b6:
                            # B6 sites strictly never cluster
                            sites_isolated.append(site_data)
                        else:
                            # A6 and MM Wave skip clustering if < 50
                            if dist_wh < 50:
                                sites_isolated.append(site_data)
                            else:
                                sites_for_triplets.append(site_data)
                    except:
                        continue
                        
                routes = []
                
                # 1. Direct Isolated Routes (B6 or <50km)
                for s in sites_isolated:
                    routes.append([{"site": s, "haversine_dist": haversine(wh_coords, s['coords']), "api_dist": 0.0}])
                    
                # 2. Main Triplets Engine (Only A6/MM >= 50km)
                if sites_for_triplets:
                    print(f"  -> Routing CMP triplet engine: {cmp_name} ({len(sites_for_triplets)} sites)")
                    routes.extend(run_routing(wh_coords, sites_for_triplets))
                
                if not routes: continue
                
                # 3. Fetch True Driving API Distances and Deduct
                for r_idx, route in enumerate(routes):
                    current_origin = wh_coords
                    for s_idx, leg in enumerate(route):
                        # Generate Labels
                        leg['club_new'] = f"{cmp_name}-R{r_idx+1}-S{s_idx+1}"
                        leg['r_id'] = f"{cmp_name}-R{r_idx+1}"
                        leg['wh_name'] = wh_name
                        leg['wh_coords'] = wh_coords
                        
                        # API Call
                        dest = leg['site']['coords']
                        api_dist = get_api_driving_distance(gmaps, current_origin, dest)
                        
                        # Apply Deduction rules strictly on FIRST leg
                        if s_idx == 0:
                            if is_b6:
                                api_dist = max(0.0, api_dist - 100.0) # B6 specific deduction
                            else:
                                api_dist = max(0.0, api_dist - 50.0)  # A6 and MMWave deduction
                                
                        leg['api_dist'] = round(api_dist, 2)
                        current_origin = dest
                        all_auto_legs.append(leg)
                        
                    auto_for_map.append(route)
            
            if not all_auto_legs:
                df_date.to_excel(writer, sheet_name=date_str, index=False)
                continue
                
            # Maps for this date
            # We assume one predominant warehouse for the sheet just to anchor the map view
            main_wh_name = all_auto_legs[0]['wh_name']
            main_wh_coords = all_auto_legs[0]['wh_coords']
            map_path = os.path.join(MAPS_DIR, f"{date_str}_routed.html")
            plot_map(main_wh_coords, auto_for_map, map_path, main_wh_name, f"Routes for {date_str}")
            
            # Map back to DataFrame
            id_to_auto = {leg['site']['id']: leg for leg in all_auto_legs}
            
            # Add new columns to df
            aktbc_new = []
            clubbing_new = []
            rid_list = []
            
            for _, row in df_date.iterrows():
                site_id = row['SITE ID']
                if site_id in id_to_auto:
                    leg = id_to_auto[site_id]
                    aktbc_new.append(leg['api_dist'])
                    clubbing_new.append(leg['club_new'])
                    rid_list.append(leg['r_id'])
                else:
                    aktbc_new.append("N/A")
                    clubbing_new.append("N/A")
                    rid_list.append("N/A")
                    
            df_date['AKTBC NEW'] = aktbc_new
            df_date['CLUBBING NEW'] = clubbing_new
            df_date['_rid'] = rid_list
            
            # Sort explicitly by CLUBBING NEW
            df_date.sort_values(by="CLUBBING NEW", key=lambda col: [natural_sort_key(c) for c in col], inplace=True)
            
            # Write sheet formatted
            ws = wb.add_worksheet(date_str)
            cols = [c for c in df_date.columns if c != '_rid']
            
            # Write Headings
            for c, h in enumerate(cols):
                ws.write(0, c, h, fmts['head'])
                
            unique_rids = list(dict.fromkeys([r for r in df_date['_rid'] if r != "N/A"]))
            rid_color_map = {rid: i for i, rid in enumerate(unique_rids)}
            
            # Write Rows
            for r_idx, (_, row) in enumerate(df_date.iterrows()):
                rid = row['_rid']
                if rid != "N/A":
                    # Color alternate routes
                    color_idx = rid_color_map[rid] % len(HEX_COLORS)
                    hex_c = HEX_COLORS[color_idx]
                    fmt = wb.add_format({'bg_color': hex_c, 'border': 1, 'align': 'center'})
                else:
                    fmt = fmts['data']
                    
                for c_idx, col in enumerate(cols):
                    ws.write(r_idx + 1, c_idx, row[col], fmt)
            
            # Set basic col widths
            for i in range(len(cols)):
                ws.set_column(i, i, 16)

    print(f"\n✅ All dates routed successfully. \nOutput file: {output_file}\nMaps folder: {MAPS_DIR}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Routing Logic Pipeline")
    parser.add_argument("api_key", help="Google Maps Directions API Key")
    parser.add_argument("--input", default=INPUT_FILE, help="Input grouped excel file")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Output routed excel file")
    
    args = parser.parse_args()
    process_pipeline(args.api_key, args.input, args.output)
