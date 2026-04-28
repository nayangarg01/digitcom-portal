import sys
import os
import pandas as pd
import numpy as np
import googlemaps
import json
import math
import itertools
import requests

# ──────────────────────────────────────────────
# ROUTING CORE MATH
# ──────────────────────────────────────────────
import math

WH_COORDS = {
    "JAIPUR": (26.810486, 75.496696),
    "JODHPUR": (26.148422, 73.061378),
    "UP": (26.8993, 81.1041),
    "DEFAULT": (26.810486, 75.496696)
}

def get_wh_coords(wh_name):
    wh_name = str(wh_name).upper().strip()
    if 'JLJH' in wh_name or 'JOD' in wh_name: return WH_COORDS['JODHPUR']
    if 'JLKD' in wh_name or 'JAIPUR' in wh_name: return WH_COORDS['JAIPUR']
    if 'JLJQ' in wh_name or 'UP' in wh_name: return WH_COORDS['UP']
    return WH_COORDS['DEFAULT']

def haversine(coord1, coord2):
    R = 6371.0
    lat1, lon1 = math.radians(coord1[0]), math.radians(coord1[1])
    lat2, lon2 = math.radians(coord2[0]), math.radians(coord2[1])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def calculate_bearing(origin, target):
    """Calculates the bearing from origin (lat, lon) to target (lat, lon) in degrees."""
    lat1, lon1 = math.radians(origin[0]), math.radians(origin[1])
    lat2, lon2 = math.radians(target[0]), math.radians(target[1])
    d_lon = lon2 - lon1
    y = math.sin(d_lon) * math.cos(lat2)
    x = math.cos(lat1) * math.sin(lat2) - math.sin(lat1) * math.cos(lat2) * math.cos(d_lon)
    bearing = math.atan2(y, x)
    return (math.degrees(bearing) + 360) % 360

def angular_diff(a, b):
    """Calculates shortest angular difference between two bearings."""
    diff = abs(a - b) % 360
    return min(diff, 360 - diff)

def get_api_driving_distance(gmaps, origin, dest):
    """
    STRICT IMPLEMENTATION: Mirrors Google Maps UI by selecting the shortest among visible routes.
    Returns: Rounded Integer KM to match billing requirements.
    """
    if not gmaps or not gmaps.key:
        raise Exception("Maps API key is missing. Strict API mode is enabled.")

    api_key = gmaps.key
    url = "https://routes.googleapis.com/directions/v2:computeRoutes"
    
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": "routes.distanceMeters"
    }

    body = {
        "origin": { "location": { "latLng": { "latitude": origin[0], "longitude": origin[1] } } },
        "destination": { "location": { "latLng": { "latitude": dest[0], "longitude": dest[1] } } },
        "travelMode": "DRIVE",
        "routingPreference": "TRAFFIC_UNAWARE",
        "computeAlternativeRoutes": True,
        "requestedReferenceRoutes": []
    }

    try:
        response = requests.post(url, headers=headers, json=body, timeout=15)
        
        if response.status_code == 200:
            data = response.json()
            if 'routes' in data and len(data['routes']) > 0:
                # Mirror Logic: Pick the shortest distance among standard visible routes
                min_m = min(r.get('distanceMeters', float('inf')) for r in data['routes'])
                # Return as rounded integer KM
                return int(round(min_m / 1000.0))
            else:
                raise Exception(f"Routes API: No visible routes found between {origin} and {dest}")
        else:
            raise Exception(f"Routes API Error {response.status_code}: {response.text}")
    except Exception as e:
        sys.stderr.write(f"ERROR: Distance Calculation Failed: {str(e)}\n")
        raise e

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
    # ADATIVE REACH: If sites are very far from WH, allow a larger pairing radius
    # DIRECTIONAL BIAS: Prefer sites along the same bearing from WH to create "corridors"
    ANGULAR_SENSITIVITY = 1.2 # Weight for bearing mismatch (km per degree)
    
    unvisited_mixer = list(mixer_pool)
    while unvisited_mixer:
        seed = max(unvisited_mixer, key=lambda s: haversine(warehouse_coords, s['coords']))
        unvisited_mixer.remove(seed)
        clump = [seed]
        
        # Calculate seed bearing
        seed_bearing = calculate_bearing(warehouse_coords, seed['coords'])
        
        # Determine clustering threshold based on distance from base
        dist_from_wh = haversine(warehouse_coords, seed['coords'])
        # If site is > 120km out, we allow a very large pairing radius (200km) to capture distant neighbors
        current_max_dist = 40.0 if dist_from_wh < 120 else 200.0
        
        for _ in range(2):
            if not unvisited_mixer: break
            
            # Find nearest with angular penalty
            # Metric = PhysicsDistance + (AngleDiff * Penalty)
            def selection_score(candidate):
                dist = haversine(seed['coords'], candidate['coords'])
                bearing = calculate_bearing(warehouse_coords, candidate['coords'])
                angle_mismatch = angular_diff(seed_bearing, bearing)
                return dist + (angle_mismatch * ANGULAR_SENSITIVITY)
                
            nearest = min(unvisited_mixer, key=selection_score)
            
            # IF nearest is too far (Adaptive), stop clustering this route
            if haversine(seed['coords'], nearest['coords']) > current_max_dist:
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
                    
                wh_raw = str(row[wh_col]).strip().upper() if wh_col else "DEFAULT"
                if 'JLJH' in wh_raw or 'JOD' in wh_raw: wh_name = 'JODHPUR'
                elif 'JLKD' in wh_raw or 'JAP' in wh_raw: wh_name = 'JAIPUR'
                else: wh_name = wh_raw
                
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

        groups = df_processing.groupby(['BAND', 'DATE', 'CMP', 'WH_NAME'])
        
        for (band, date_val, cmp_name, wh_name_for_group), group_df in groups:
            # 1. Fetch exact Warehouse Coordinates purely based on what is MENTIONED
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
                    
                    api_dist = get_api_driving_distance(gmaps, current_origin, dest)
                    
                    base_api_dist = api_dist
                    
                    if s_idx == 0:
                        true_wh_dist = base_api_dist
                        if is_b6: api_dist = max(0.0, api_dist - 100.0)
                        else: api_dist = max(0.0, api_dist - 50.0)
                    else:
                        true_wh_dist = get_api_driving_distance(gmaps, wh_coords, dest)
                            
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
