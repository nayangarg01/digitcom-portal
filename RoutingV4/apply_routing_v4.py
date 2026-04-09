import pandas as pd
import numpy as np
import googlemaps
import folium
import math
import itertools
import argparse
import sys
import os

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
MANUAL_FILE    = "BILLING PENDING SITES 2.xlsx"
OUTPUT_FILE    = "BILLING_PENDING_SITES_AutoValidated.xlsx"

WH_COORDS_FALLBACK = {
    "JAIPUR":  (26.810486, 75.496696),
    "JODHPUR": (26.148422, 73.061378),
}

# ──────────────────────────────────────────────
# ROUTING CORE MATH
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
    try:
        routes = gmaps.directions(origin, dest, mode='driving', alternatives=True)
        if routes:
            def total_dist(r):
                return sum(leg['distance']['value'] for leg in r['legs'])
            shortest_route = min(routes, key=total_dist)
            dist_m = total_dist(shortest_route)
            return round(dist_m / 1000.0, 2)
    except Exception as e:
        pass
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
    route_legs = []; cp = warehouse_coords
    for s in segment:
        d = haversine(cp, s['coords'])
        route_legs.append({"site": s, "haversine_dist": round(d, 2), "api_dist": 0.0})
        cp = s['coords']
    return route_legs

def run_routing(warehouse_coords, cluster):
    jc_groups = {}
    for s in cluster:
        row = s.get('row_data', {})
        jc = str(row.get('INJECTED_JC', '')).strip().upper()
        if jc not in jc_groups: jc_groups[jc] = []
        jc_groups[jc].append(s)
    
    for s in cluster:
        my_jc = str(s.get('row_data', {}).get('INJECTED_JC', '')).strip().upper()
        others = [o for o in cluster if str(o.get('row_data', {}).get('INJECTED_JC', '')).strip().upper() != my_jc]
        if others: s['trans_dist'] = min(haversine(s['coords'], o['coords']) for o in others)
        else: s['trans_dist'] = 999999.0
            
    final_routes = []; mixer_pool = []
    
    for jc, sites in jc_groups.items():
        unvisited = sorted(sites, key=lambda x: x['trans_dist'], reverse=True)
        while len(unvisited) >= 3:
            seed = unvisited.pop(0)
            clump = [seed]
            for _ in range(2):
                nearest = min(unvisited, key=lambda s: haversine(seed['coords'], s['coords']))
                clump.append(nearest)
                unvisited.remove(nearest)
            best_p = optimize_segment(warehouse_coords, clump)
            final_routes.append(segment_to_legs(warehouse_coords, best_p))
        mixer_pool.extend(unvisited)
    
    unvisited_mixer = list(mixer_pool)
    while unvisited_mixer:
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
# MAPPING ENGINE
# ──────────────────────────────────────────────
ROUTE_COLORS = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue']

def plot_map(warehouse_coords, routes, output_file, wh_name, title):
    m = folium.Map(location=[warehouse_coords[0], warehouse_coords[1]], zoom_start=9)
    folium.Marker(warehouse_coords, popup=f"WAREHOUSE: {wh_name}", icon=folium.Icon(color='black', icon='home')).add_to(m)
    
    for r_idx, route in enumerate(routes):
        color = ROUTE_COLORS[r_idx % len(ROUTE_COLORS)]
        coords = [warehouse_coords]
        for leg in route:
            s = leg['site']
            lat, lon = s['coords']
            coords.append((lat, lon))
            pop = f"Site: {s['id']}<br>Band: {s['BAND']}"
            folium.Marker((lat, lon), popup=pop, icon=folium.Icon(color=color)).add_to(m)
            
        if len(coords) > 1:
            folium.PolyLine(coords, color=color, weight=4, opacity=0.8).add_to(m)
            
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    m.save(output_file)

def guess_manual_route(wh_coords, sites_list):
    route = []
    curr = wh_coords
    unvisited = sites_list.copy()
    while unvisited:
        nxt = min(unvisited, key=lambda s: haversine(curr, s['coords']))
        route.append({'site': nxt})
        curr = nxt['coords']
        unvisited.remove(nxt)
    return route

# ──────────────────────────────────────────────
# MAIN PIPELINE
# ──────────────────────────────────────────────
def run_direct_validation(api_key):
    gmaps = googlemaps.Client(key=api_key, timeout=10)
    
    print(f"Loading direct manual dataset: {MANUAL_FILE}...")
    df_manual = pd.read_excel(MANUAL_FILE).fillna("")
    
    # Create tracking columns natively
    df_manual['AUTO ROUTE CLUBBING'] = ''
    df_manual['AUTO DISTANCE WH'] = ''
    df_manual['AUTO AKTBC'] = ''
    
    processing_sites = []
    
    # We only process A6 and 2026. Keep track of what we process.
    for idx, row in df_manual.iterrows():
        activity = str(row.get('Activity', 'UNKNOWN')).strip().upper()
        # Clean date for grouping
        date_raw = str(row.get('MIN DATE', '')).strip()
        date_obj = pd.to_datetime(date_raw, errors='coerce')
        
        is_a6 = "A6" in activity
        is_2026 = (date_obj is not pd.NaT and date_obj.year == 2026)
        
        # Only process A6 and 2026
        if not (is_a6 and is_2026):
            continue
            
        if ' ' in date_raw and ':' in date_raw: 
            date_raw = date_raw.split()[0]
        elif date_obj is not pd.NaT:
            date_raw = date_obj.strftime("%Y-%m-%d")
        
        # Pull Band mapping directly from 'Activity'
        band = activity
        
        # Use existing cmp column
        cmp_val = str(row.get('cmp', '')).strip().upper()
        if not cmp_val:
            cmp_val = str(row.get('CMP', '')).strip().upper() 
            
        jc = str(row.get('JC', '')).strip().upper()
        
        # Extract Coords
        try: 
            lat, lon = float(row.get('LAT ', 0)), float(row.get('LONG', 0))
        except: 
            lat, lon = 0, 0
        
        wh_name = str(row.get('WAREHOUSE', 'JAIPUR')).strip().upper()
        if wh_name == 'JLJH' or wh_name == 'JOD': 
            wh_name = 'JODHPUR'
        
        site_data = {
            'df_idx': idx,
            'id': str(row.get('eNBsiteID', idx)),
            'coords': (lat, lon),
            'BAND': band,
            'DATE': date_raw,
            'CMP': cmp_val,
            'WH_NAME': wh_name,
            'row_data': {'INJECTED_JC': jc}
        }
        processing_sites.append(site_data)
        
    df_processing = pd.DataFrame(processing_sites)
    
    if df_processing.empty:
        print("No A6 sites from 2026 found. Exiting.")
        return
        
    # Track final assignment 
    final_output = {} # idx -> {'club': '', 'dist': '', 'aktbc': ''}
    
    # ── 2. Route by Band -> Date -> CMP ──
    groups = df_processing.groupby(['BAND', 'DATE', 'CMP'])
    print(f"Divided manual spreadsheet into {len(groups)} logical isolation chunks (Band+Date+CMP).")
    
    for (band, date_val, cmp_name), group_df in groups:
        if date_val == "": continue
        
        # Pre-filter WH assignments based STRICTLY on the manual sheet inputs
        wh_name = group_df.iloc[0]['WH_NAME']
        wh_coords = WH_COORDS_FALLBACK.get(wh_name, WH_COORDS_FALLBACK['JAIPUR'])
        
        is_b6 = "B6" in band.upper()
        sites_isolated = []
        sites_triplet = []
        
        for _, s_row in group_df.iterrows():
            s_dict = s_row.to_dict()
            lat, lon = s_dict['coords']
            if lat == 0 or math.isnan(lat): continue
            
            dist_wh = haversine(wh_coords, (lat, lon))
            if is_b6:
                sites_isolated.append(s_dict)
            else:
                if dist_wh < 50: sites_isolated.append(s_dict)
                else: sites_triplet.append(s_dict)
                
        routes = []
        for s in sites_isolated:
            routes.append([{"site": s, "haversine_dist": haversine(wh_coords, s['coords']), "api_dist": 0.0}])
            
        if sites_triplet:
            print(f"  -> Routing Engine ({band}) | {date_val} | {cmp_name} | {len(sites_triplet)} sites via Triplets")
            routes.extend(run_routing(wh_coords, sites_triplet))
            
        # ── 3. Loop API calculations and apply strictly local deductions ──
        for r_idx, route in enumerate(routes):
            current_origin = wh_coords
            for s_idx, leg in enumerate(route):
                s_dict = leg['site']
                
                # API Call distance for current routing leg
                dest = s_dict['coords']
                try: api_dist = get_api_driving_distance(gmaps, current_origin, dest)
                except: api_dist = leg['haversine_dist']
                
                base_api_dist = api_dist
                
                # Independently Fetch True Warehouse Distance for the output column
                if s_idx == 0:
                    true_wh_dist = base_api_dist
                    if is_b6: api_dist = max(0.0, api_dist - 100.0)
                    else: api_dist = max(0.0, api_dist - 50.0)
                else:
                    try: true_wh_dist = get_api_driving_distance(gmaps, wh_coords, dest)
                    except: true_wh_dist = haversine(wh_coords, dest)
                        
                final_output[s_dict['df_idx']] = {
                    'club': f"{s_dict['CMP']}-R{r_idx+1}-S{s_idx+1}",
                    'dist': round(true_wh_dist, 2),
                    'aktbc': round(api_dist, 2)
                }
                
                current_origin = dest
                
        # ── MAP GENERATION ──
        # Fix filename safety
        dist_str = cmp_name.replace("/", "-")
        date_safe = date_val.replace("/", "-")
        bnd_safe = band.replace("/", "-")
        
        # Plot Automated Maps
        plot_map(wh_coords, routes, f"Maps/{date_safe}_{dist_str}_{bnd_safe}_AUTO.html", wh_name, "AUTO ROUTES")
        
        # Plot Manual Maps (Guessing visual lines from their Clubbing groups)
        manual_routes = []
        for club_name, club_df in group_df.groupby(lambda i: str(df_manual.at[i, 'CLUBBING'])):
            if club_name == "" or club_name == "nan": continue
            m_sites = []
            for _, sr in club_df.iterrows():
                m_sites.append(sr.to_dict())
            if m_sites:
                manual_routes.append(guess_manual_route(wh_coords, m_sites))
                
        if manual_routes:
            plot_map(wh_coords, manual_routes, f"Maps/{date_safe}_{dist_str}_{bnd_safe}_MANUAL.html", wh_name, "MANUAL ROUTES")
                
    # ── 4. Append mathematically back to rows ──
    for idx, calc in final_output.items():
        df_manual.at[idx, 'AUTO ROUTE CLUBBING'] = calc['club']
        df_manual.at[idx, 'AUTO DISTANCE WH'] = calc['dist']
        df_manual.at[idx, 'AUTO AKTBC'] = calc['aktbc']
        
    print(f"\n✅ Writing validated matrix directly to {OUTPUT_FILE}...")
    
    # Restructure Column layout side-by-side
    cols = list(df_manual.columns)
    cols.remove('AUTO ROUTE CLUBBING')
    cols.remove('AUTO DISTANCE WH')
    cols.remove('AUTO AKTBC')
    
    if 'CLUBBING' in cols: cols.insert(cols.index('CLUBBING') + 1, 'AUTO ROUTE CLUBBING')
    else: cols.append('AUTO ROUTE CLUBBING')
        
    if 'KM FROM WH TO SITE' in cols: cols.insert(cols.index('KM FROM WH TO SITE') + 1, 'AUTO DISTANCE WH')
    else: cols.append('AUTO DISTANCE WH')
        
    if 'AKTBC' in cols: cols.insert(cols.index('AKTBC') + 1, 'AUTO AKTBC')
    else: cols.append('AUTO AKTBC')
        
    df_manual = df_manual[cols]
    
    # Format DATE columns securely to avoid Excel numerical serial format (e.g. 46051)
    if 'MIN DATE' in df_manual.columns:
        df_manual['MIN DATE'] = pd.to_datetime(df_manual['MIN DATE'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
    if 'RFS DATE' in df_manual.columns:
        df_manual['RFS DATE'] = pd.to_datetime(df_manual['RFS DATE'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
        
    # We will just write it straight out seamlessly
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True, 'strings_to_urls': False}}) as writer:
        wb = writer.book
        ws = wb.add_worksheet("Automation Audit Matrix")
        
        # Define formats natively
        header_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#2F5597', 'font_color': 'white', 'border': 1})
        std_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        auto_fmt = wb.add_format({'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        auto_num_fmt = wb.add_format({'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
        
        # Write Headers
        for c_idx, col_name in enumerate(cols):
            ws.write(0, c_idx, str(col_name), header_fmt)
            
        # Write Rows
        for r_idx, row in df_manual.iterrows():
            row_dict = row.to_dict()
            for c_idx, col_name in enumerate(cols):
                val = row_dict.get(col_name, "")
                
                # Apply formats conditionally
                if col_name == 'AUTO ROUTE CLUBBING':
                    ws.write(r_idx + 1, c_idx, val, auto_fmt)
                elif col_name in ['AUTO DISTANCE WH', 'AUTO AKTBC']:
                    try: val = float(val)
                    except: pass
                    ws.write(r_idx + 1, c_idx, val, auto_num_fmt)
                else:
                    ws.write(r_idx + 1, c_idx, val, std_fmt)
                    
        # Column Widths
        for c_idx, col_name in enumerate(cols):
            max_len = max([len(str(r.get(col_name, ""))) for _, r in df_manual.iterrows()], default=0)
            max_len = max(max_len, len(str(col_name))) + 4
            max_len = min(max_len, 40) # Cap safety
            ws.set_column(c_idx, c_idx, max_len)
            
    print("Process Complete. Maps generated in Maps/ directory. Enjoy the side-by-side audit.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("api_key", help="Google Directions API Key")
    args = parser.parse_args()
    run_direct_validation(args.api_key)
