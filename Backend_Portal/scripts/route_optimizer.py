import sys
import os
import pandas as pd
import numpy as np
import googlemaps
import json
import math
import itertools

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
    # Group sites by JC for prioritization
    jc_groups = {}
    for s in cluster:
        row = s.get('row_data', {})
        jc = str(row.get('INJECTED_JC', '')).strip().upper()
        if jc not in jc_groups: jc_groups[jc] = []
        jc_groups[jc].append(s)
    
    # Pre-calculate distances to nearest neighbors outside own JC to find 'frontier' sites
    for s in cluster:
        my_jc = str(s.get('row_data', {}).get('INJECTED_JC', '')).strip().upper()
        others = [o for o in cluster if str(o.get('row_data', {}).get('INJECTED_JC', '')).strip().upper() != my_jc]
        if others: s['trans_dist'] = min(haversine(s['coords'], o['coords']) for o in others)
        else: s['trans_dist'] = 999999.0
            
    final_routes = []; mixer_pool = []
    
    # Phase 1: JC Integrity (Clustering 2 or 3 sites within the same JC)
    for jc, sites in jc_groups.items():
        # Sort by proximity to other districts to keep inner JC sites grouped together
        unvisited = sorted(sites, key=lambda x: x['trans_dist'], reverse=True)
        
        # 1. Take Triplets
        while len(unvisited) >= 3:
            seed = unvisited.pop(0)
            clump = [seed]
            for _ in range(2):
                nearest = min(unvisited, key=lambda s: haversine(seed['coords'], s['coords']))
                clump.append(nearest)
                unvisited.remove(nearest)
            best_p = optimize_segment(warehouse_coords, clump)
            final_routes.append(segment_to_legs(warehouse_coords, best_p))
            
        # 2. Take Pairs (New Logic: Prioritize local pairing over global mixing)
        while len(unvisited) >= 2:
            seed = unvisited.pop(0)
            nearest = min(unvisited, key=lambda s: haversine(seed['coords'], s['coords']))
            clump = [seed, nearest]
            unvisited.remove(nearest)
            best_p = optimize_segment(warehouse_coords, clump)
            final_routes.append(segment_to_legs(warehouse_coords, best_p))
            
        mixer_pool.extend(unvisited)
    
    # Phase 2: Geographical Mixer for leftovers
    # Only mix if sites are genuinely close (e.g., < 40km), else keep them isolated
    MAX_MIX_DIST = 40.0 
    unvisited_mixer = list(mixer_pool)
    while unvisited_mixer:
        seed = max(unvisited_mixer, key=lambda s: haversine(warehouse_coords, s['coords']))
        unvisited_mixer.remove(seed)
        clump = [seed]
        
        for _ in range(2):
            if not unvisited_mixer: break
            nearest = min(unvisited_mixer, key=lambda s: haversine(seed['coords'], s['coords']))
            # IF nearest is too far, stop clustering this route
            if haversine(seed['coords'], nearest['coords']) > MAX_MIX_DIST:
                break
            clump.append(nearest)
            unvisited_mixer.remove(nearest)
            
        best_p = optimize_segment(warehouse_coords, clump)
        final_routes.append(segment_to_legs(warehouse_coords, best_p))
        
    return final_routes


def main():
    if len(sys.argv) < 5:
        print(json.dumps({"error": "Missing arguments"}))
        return

    file_path, origin_lat, origin_lng, api_key, output_path = sys.argv[1:6]
    try:
        gmaps = googlemaps.Client(key=api_key, timeout=10)
    except:
        gmaps = None

    # Warehouse Fallbacks
    WH_COORDS_FALLBACK = {
        "JAIPUR": (26.810486, 75.496696),
        "JODHPUR": (26.148422, 73.061378),
        "DEFAULT": (float(origin_lat), float(origin_lng))
    }

    try:
        # Load File
        df = pd.read_excel(file_path) if file_path.endswith(('.xlsx', '.xls')) else pd.read_csv(file_path)
        
        # Determine Columns robustly
        lat_col = next((c for c in df.columns if c.strip().lower() in ['latitude', 'lat', 'lat ']), None)
        lng_col = next((c for c in df.columns if c.strip().lower() in ['longitude', 'lng', 'lon', 'long', 'long ']), None)
        id_col = next((c for c in df.columns if c.strip().lower() in ['site id', 'site_id', 'siteid', 'enbsiteid']), None)
        cmp_col = next((c for c in df.columns if c.strip().lower() in ['cmp', 'company']), None)
        wh_col = next((c for c in df.columns if c.strip().lower() in ['wh', 'warehouse_name', 'wh ', 'warehouse']), None)
        jc_col = next((c for c in df.columns if c.strip().lower() in ['jc', 'jio center', 'jio_center']), None)
        date_col = next((c for c in df.columns if c.strip().lower() in ['min date', 'date', 'rfs date']), None)
        band_col = next((c for c in df.columns if c.strip().lower() in ['activity', 'band', 'site type']), None)

        if not lat_col or not lng_col:
            print(json.dumps({"error": "Missing precise GPS Latitude/Longitude columns in the dataset"}))
            return
            
        # Parse into engine internal structure
        processing_sites = []
        for idx, row in df.iterrows():
            try:
                lat, lng = float(row[lat_col]), float(row[lng_col])
                if np.isnan(lat) or np.isnan(lng): continue
                
                band = str(row[band_col]).strip().upper() if band_col else "UNKNOWN"
                cmp_val = str(row[cmp_col]).strip().upper() if cmp_col else "DEFAULT"
                jc_val = str(row[jc_col]).strip().upper() if jc_col else ""
                date_raw = str(row[date_col]).strip() if date_col else "NO_DATE"
                date_obj = pd.to_datetime(date_raw, errors='coerce')
                
                if ' ' in date_raw and ':' in date_raw: 
                    date_raw = date_raw.split()[0]
                elif date_obj is not pd.NaT:
                    date_raw = date_obj.strftime("%Y-%m-%d")
                    
                wh_name = str(row[wh_col]).strip().upper() if wh_col else "DEFAULT"
                if wh_name == 'JLJH' or wh_name == 'JOD': wh_name = 'JODHPUR'
                
                site_data = {
                    'df_idx': idx,
                    'id': str(row[id_col]) if id_col else str(idx),
                    'coords': (lat, lng),
                    'BAND': band,
                    'DATE': date_raw,
                    'CMP': cmp_val,
                    'WH_NAME': wh_name,
                    'row_data': {'INJECTED_JC': jc_val}
                }
                processing_sites.append(site_data)
            except Exception as e:
                continue

        if not processing_sites:
            print(json.dumps({"error": "No valid sites extracted"}))
            return
            
        df_processing = pd.DataFrame(processing_sites)
        
        final_output = {}
        routes_json = []
        route_global_counter = 0

        groups = df_processing.groupby(['BAND', 'DATE', 'CMP'])
        
        for (band, date_val, cmp_name), group_df in groups:
            # 1. Fetch exact Warehouse Coordinates purely based on what is MENTIONED
            wh_name_for_group = group_df.iloc[0]['WH_NAME']
            wh_coords = WH_COORDS_FALLBACK.get(wh_name_for_group, WH_COORDS_FALLBACK.get('DEFAULT'))

            is_b6 = "B6" in band.upper()
            sites_isolated = []
            sites_triplet = []
            
            for _, s_row in group_df.iterrows():
                s_dict = s_row.to_dict()
                if is_b6:
                    sites_isolated.append(s_dict)
                else:
                    # Intelligence over hard-rules: Pass ALL normal sites to the clusterer
                    # The clusterer will naturally isolate a site if no neighbor is within 40km
                    sites_triplet.append(s_dict)
                    
            routes = []
            for s in sites_isolated:
                routes.append([{"site": s, "haversine_dist": haversine(wh_coords, s['coords']), "api_dist": 0.0}])
                
            if sites_triplet:
                routes.extend(run_routing(wh_coords, sites_triplet))
                
            for route_loop_idx, route in enumerate(routes):
                route_global_counter += 1
                current_origin = wh_coords
                
                # Setup JSON representation for this route
                route_label = f"R{route_global_counter}" # Simplified label logic for JSON UX
                route_obj = {
                    "routeNumber": route_global_counter, 
                    "label": route_label, 
                    "origin_coords": {"lat": current_origin[0], "lng": current_origin[1]},
                    "legs": []
                }
                
                for s_idx, leg in enumerate(route):
                    s_dict = leg['site']
                    dest = s_dict['coords']
                    
                    try: api_dist = get_api_driving_distance(gmaps, current_origin, dest)
                    except: api_dist = leg['haversine_dist']
                    
                    base_api_dist = api_dist
                    
                    if s_idx == 0:
                        true_wh_dist = base_api_dist
                        if is_b6: api_dist = max(0.0, api_dist - 100.0)
                        else: api_dist = max(0.0, api_dist - 50.0)
                    else:
                        try: true_wh_dist = get_api_driving_distance(gmaps, wh_coords, dest)
                        except: true_wh_dist = haversine(wh_coords, dest)
                            
                    club_str = f"{s_dict['CMP']}-R{route_global_counter}-S{s_idx+1}"
                    aktbc_val = max(0.0, api_dist)
                    
                    # Store to Final Matrix structure
                    final_output[s_dict['df_idx']] = {
                        'club': club_str,
                        'dist': round(true_wh_dist, 2),
                        'aktbc': round(aktbc_val, 2)
                    }
                    
                    # Store to JSON structure
                    route_obj["legs"].append({
                        "routeLabel": club_str,
                        "stopSequence": s_idx + 1,
                        "distanceKm": round(aktbc_val, 2),
                        "site": {"id": s_dict['id'], "lat": dest[0], "lng": dest[1]}
                    })
                    
                    current_origin = dest
                routes_json.append(route_obj)

        if 'CLUBBING' not in df.columns: df['CLUBBING'] = ""
        else: df['CLUBBING'] = df['CLUBBING'].astype(object)
        
        if 'KM FROM WH TO SITE' not in df.columns: df['KM FROM WH TO SITE'] = 0.0
        df['KM FROM WH TO SITE'] = df['KM FROM WH TO SITE'].astype(object)
        
        if 'AKTBC' not in df.columns: df['AKTBC'] = 0.0
        df['AKTBC'] = df['AKTBC'].astype(object)

        for idx, calc in final_output.items():
            df.at[idx, 'CLUBBING'] = calc['club']
            df.at[idx, 'KM FROM WH TO SITE'] = calc['dist']
            df.at[idx, 'AKTBC'] = calc['aktbc']

        # Format dates properly
        if date_col:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
        rfs_col = next((c for c in df.columns if c.strip().lower() in ['rfs date', 'rfs_date']), None)
        if rfs_col:
            df[rfs_col] = pd.to_datetime(df[rfs_col], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')

        # Final Formatting and Export
        with pd.ExcelWriter(output_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True, 'strings_to_urls': False}}) as writer:
            wb = writer.book
            ws = wb.add_worksheet("Optimized Routes")
            
            # Format classes
            header_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#2F5597', 'font_color': 'white', 'border': 1})
            std_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
            auto_fmt = wb.add_format({'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            num_fmt = wb.add_format({'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
            
            # Write columns seamlessly
            cols = list(df.columns)
            
            # Make sure CLUBBING, KM FROM WH TO SITE, AKTBC are together where possible or just re-insert them 
            # We will just write the columns in original order, but placing them if they were newly created at the end.
            for c_idx, col_name in enumerate(cols):
                ws.write(0, c_idx, str(col_name), header_fmt)
                
            for r_idx, row in df.iterrows():
                row_dict = row.to_dict()
                for c_idx, col_name in enumerate(cols):
                    val = row_dict.get(col_name, "")
                    if col_name in ['KM FROM WH TO SITE', 'AKTBC']:
                        try: val = float(val)
                        except: pass
                        ws.write(r_idx + 1, c_idx, val, num_fmt)
                    elif col_name == 'CLUBBING':
                        ws.write(r_idx + 1, c_idx, val, auto_fmt)
                    else:
                        ws.write(r_idx + 1, c_idx, val, std_fmt)
                        
            # Adjust column width
            for c_idx, col_name in enumerate(cols):
                max_len = min(max(len(str(col_name)), max(len(str(v)) for v in df[col_name].astype(str))), 40)
                ws.set_column(c_idx, c_idx, max_len + 2)

        print(json.dumps({"success": True, "num_routes": len(routes_json), "routes": routes_json}))

    except Exception as e:
        print(json.dumps({"error": str(e)}))

if __name__ == "__main__":
    main()
