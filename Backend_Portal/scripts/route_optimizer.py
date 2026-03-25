import sys
import os
import pandas as pd
import numpy as np
import requests
import json
from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp
import openpyxl
from math import ceil, atan2

def get_distance_matrix(locations, api_key):
    """
    Fetch full distance matrix from Google Maps API in parallel batches.
    """
    num_locations = len(locations)
    matrix = np.zeros((num_locations, num_locations))
    batch_size = 5
    
    from concurrent.futures import ThreadPoolExecutor

    batches = []
    for i in range(0, num_locations, batch_size):
        for j in range(0, num_locations, batch_size):
            batches.append((i, j))

    def fetch_batch(coords):
        start_i, start_j = coords
        origins_batch = locations[start_i : min(start_i + batch_size, num_locations)]
        dest_batch = locations[start_j : min(start_j + batch_size, num_locations)]
        
        origin_str = "|".join([f"{round(lat, 6)},{round(lng, 6)}" for lat, lng in origins_batch])
        dest_str = "|".join([f"{round(lat, 6)},{round(lng, 6)}" for lat, lng in dest_batch])
        
        url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={origin_str}&destinations={dest_str}&avoid=highways&key={api_key}"
        
        try:
            response = requests.get(url).json()
            if response['status'] == 'OK':
                return (start_i, start_j, response['rows'])
            else:
                return (start_i, start_j, None)
        except:
            return (start_i, start_j, None)

    with ThreadPoolExecutor(max_workers=10) as executor:
        results = list(executor.map(fetch_batch, batches))

    for start_i, start_j, rows in results:
        if rows is None: continue
        for row_idx, row in enumerate(rows):
            for col_idx, element in enumerate(row['elements']):
                if element['status'] == 'OK':
                    matrix[start_i + row_idx][start_j + col_idx] = element['distance']['value']
                else:
                    matrix[start_i + row_idx][start_j + col_idx] = 999999
    return matrix

def partition_sites(n):
    if n < 2: return [n]
    if n == 2: return [2]
    if n == 3: return [3]
    if n == 4: return [2, 2]
    if n % 3 == 0: return [3] * (n // 3)
    if n % 3 == 1: return [3] * ((n - 4) // 3) + [2, 2]
    return [3] * ((n - 2) // 3) + [2]

def solve_recursive_splitting(warehouse_coords, chunk, api_key):
    """
    Applies the recursive WH->B < A->B logic using real distances.
    """
    locations = [warehouse_coords] + [s['coords'] for s in chunk]
    dist_matrix = get_distance_matrix(locations, api_key)
    
    # Indices in matrix: 0=WH, 1=A, 2=B, 3=C
    p_sites = chunk.copy()
    final_routes = []
    
    while p_sites:
        if len(p_sites) == 1:
            site = p_sites.pop(0)
            dist_km = dist_matrix[0][chunk.index(site)+1] / 1000.0
            final_routes.append([{ "site": site, "dist_km": max(0, dist_km - 50) }])
        elif len(p_sites) == 2:
            A, B = p_sites[0], p_sites[1]
            idx_a, idx_b = chunk.index(A)+1, chunk.index(B)+1
            dist_wh_b = dist_matrix[0][idx_b]
            dist_a_b = dist_matrix[idx_a][idx_b]
            
            if dist_wh_b < dist_a_b:
                # Split
                dist_wh_a_km = dist_matrix[0][idx_a] / 1000.0
                final_routes.append([{ "site": p_sites.pop(0), "dist_km": max(0, dist_wh_a_km - 50) }])
            else:
                # Together
                dist_wh_a_km = dist_matrix[0][idx_a] / 1000.0
                dist_a_b_km = dist_matrix[idx_a][idx_b] / 1000.0
                final_routes.append([
                    { "site": p_sites.pop(0), "dist_km": max(0, dist_wh_a_km - 50) },
                    { "site": p_sites.pop(0), "dist_km": dist_a_b_km }
                ])
        else: # 3 sites
            A, B, C = p_sites[0], p_sites[1], p_sites[2]
            idx_a, idx_b, idx_c = chunk.index(A)+1, chunk.index(B)+1, chunk.index(C)+1
            
            if dist_matrix[0][idx_b] < dist_matrix[idx_a][idx_b]:
                # A is separate
                dist_wh_a_km = dist_matrix[0][idx_a] / 1000.0
                final_routes.append([{ "site": p_sites.pop(0), "dist_km": max(0, dist_wh_a_km - 50) }])
            elif dist_matrix[0][idx_c] < dist_matrix[idx_b][idx_c]:
                # A, B together, C separate
                dist_wh_a_km = dist_matrix[0][idx_a] / 1000.0
                dist_a_b_km = dist_matrix[idx_a][idx_b] / 1000.0
                final_routes.append([
                    { "site": p_sites.pop(0), "dist_km": max(0, dist_wh_a_km - 50) },
                    { "site": p_sites.pop(0), "dist_km": dist_a_b_km }
                ])
            else:
                # All 3 together
                dist_wh_a_km = dist_matrix[0][idx_a] / 1000.0
                dist_a_b_km = dist_matrix[idx_a][idx_b] / 1000.0
                dist_b_c_km = dist_matrix[idx_b][idx_c] / 1000.0
                final_routes.append([
                    { "site": p_sites.pop(0), "dist_km": max(0, dist_wh_a_km - 50) },
                    { "site": p_sites.pop(0), "dist_km": dist_a_b_km },
                    { "site": p_sites.pop(0), "dist_km": dist_b_c_km }
                ])
    return final_routes

def main():
    if len(sys.argv) < 5:
        print(json.dumps({"error": "Missing arguments"}))
        return

    file_path, origin_lat, origin_lng, api_key, output_path = sys.argv[1:6]
    
    # Warehouse Mapping (Jaipur - Bagru, Jodhpur - Mogra Khurd)
    WAREHOUSE_MAP = {
        'JAIPUR': (26.8139, 75.5450),
        'JODHPUR': (26.1245, 73.0543),
        'DEFAULT': (float(origin_lat), float(origin_lng))
    }

    try:
        # 1. Load Data
        df = pd.read_excel(file_path) if file_path.endswith(('.xlsx', '.xls')) else pd.read_csv(file_path)
        
        # 2. Extract Sites and Automatically Detect Warehouse Location
        lat_col = next((c for c in df.columns if c.strip().lower() in ['latitude', 'lat', 'lat ']), None)
        lng_col = next((c for c in df.columns if c.strip().lower() in ['longitude', 'lng', 'lon', 'long', 'long ']), None)
        id_col = next((c for c in df.columns if c.strip().lower() in ['site id', 'site_id', 'siteid', 'enbsiteid']), None)
        cmp_col = next((c for c in df.columns if c.strip().lower() in ['cmp', 'company']), None)
        wh_col = next((c for c in df.columns if c.strip().lower() in ['wh', 'warehouse_name', 'wh ', 'warehouse']), None)

        # Detect Warehouse from first row
        warehouse_coords = WAREHOUSE_MAP['DEFAULT']
        if wh_col and not df.empty:
            wh_val = str(df.iloc[0][wh_col]).upper().strip()
            if 'JAIPUR' in wh_val:
                warehouse_coords = WAREHOUSE_MAP['JAIPUR']
            elif 'JODHPUR' in wh_val:
                warehouse_coords = WAREHOUSE_MAP['JODHPUR']

        site_data = []
        for idx, row in df.iterrows():
            try:
                lat, lng = float(row[lat_col]), float(row[lng_col])
                if not np.isnan(lat) and not np.isnan(lng):
                    # Calculate distance and angle from warehouse
                    dist_to_wh = ((lat - warehouse_coords[0])**2 + (lng - warehouse_coords[1])**2)**0.5
                    angle = atan2(lat - warehouse_coords[0], lng - warehouse_coords[1])
                    site_data.append({
                        "id": str(row[id_col]) if id_col else str(idx),
                        "coords": (lat, lng),
                        "orig_idx": idx,
                        "dist_to_wh": dist_to_wh,
                        "angle": angle,
                        "cmp": str(row[cmp_col]).strip() if cmp_col else "Default"
                    })
            except: continue

        if not site_data:
            print(json.dumps({"error": "No valid sites found"}))
            return

        # 3. Phase 1: Group by CMP
        cmp_groups = {}
        for s in site_data:
            c = s['cmp']
            if c not in cmp_groups: cmp_groups[c] = []
            cmp_groups[c].append(s)

        final_all_routes = []
        for cmp_name, group in cmp_groups.items():
            # 3.1 Global Angular Sort (Sector-based)
            # This ensures sites in the same direction are grouped together.
            group.sort(key=lambda x: x['angle'])
            
            # 3.2 Partition into n%3 chunks
            sizes = partition_sites(len(group))
            curr = 0
            for size in sizes:
                chunk = group[curr : curr + size]
                
                # 3.3 Local Sort by Distance (Inner to Outer)
                # Helps ensure a logical progression within a sector.
                chunk.sort(key=lambda x: x['dist_to_wh'])
                
                # 3.4 Apply recursive splitting with real distances
                routes = solve_recursive_splitting(warehouse_coords, chunk, api_key)
                final_all_routes.extend(routes)
                curr += size

        # 4. Finalize Results
        df['CLUBBING'] = ""
        df['AKTBC'] = 0.0
        routes_json = []

        # Labels will be in the format: A1, A2... or with Date prefix if we had it, 
        # but here we follow the chr(65+i) style
        for r_idx, route in enumerate(final_all_routes):
            label = f"R{r_idx + 1}" # Using R1, R2... for clarity across many routes
            route_obj = {"routeNumber": r_idx + 1, "label": label, "legs": []}
            
            for s_idx, leg_data in enumerate(route):
                site = leg_data['site']
                dist_km = leg_data['dist_km']
                
                df.at[site['orig_idx'], 'CLUBBING'] = f"{label}-S{s_idx + 1}"
                df.at[site['orig_idx'], 'AKTBC'] = dist_km
                
                route_obj["legs"].append({
                    "routeLabel": label,
                    "stopSequence": s_idx + 1,
                    "distanceKm": dist_km,
                    "site": {"id": site['id'], "lat": site['coords'][0], "lng": site['coords'][1]}
                })
            routes_json.append(route_obj)

        df['sort_key'] = df['CLUBBING'].apply(lambda x: x if x else "ZZZ")
        df = df.sort_values(by='sort_key').drop(columns=['sort_key'])
        df.to_excel(output_path, index=False)
        
        print(json.dumps({"success": True, "num_routes": len(final_all_routes), "routes": routes_json}))

    except Exception as e:
        print(json.dumps({"error": str(e)}))

if __name__ == "__main__":
    main()
