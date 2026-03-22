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

def solve_tsp_for_cluster(warehouse_coords, cluster_sites, api_key):
    """
    Solve TSP for a small cluster (max 3 sites) + Warehouse.
    Always follows the "Closest to Furthest" sequence requested by the user.
    """
    # 1. Prepare locations: [Warehouse, Site1, Site2, Site3]
    locations = [warehouse_coords] + [s['coords'] for s in cluster_sites]
    num_nodes = len(locations)
    
    # 2. Get Distance Matrix for this small group
    dist_matrix = get_distance_matrix(locations, api_key)
    
    # 3. Solve for shortest sequence starting at Warehouse (Depot)
    manager = pywrapcp.RoutingIndexManager(num_nodes, 1, 0)
    routing = pywrapcp.RoutingModel(manager)

    def distance_callback(from_index, to_index):
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        return int(dist_matrix[from_node][to_node])

    transit_callback_index = routing.RegisterTransitCallback(distance_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = (
        routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC)
    
    solution = routing.SolveWithParameters(search_parameters)
    
    if not solution:
        return sorted(cluster_sites, key=lambda x: x['dist_to_wh'])

    # 4. Extract Route
    index = routing.Start(0)
    node_sequence = []
    while not routing.IsEnd(index):
        node_index = manager.IndexToNode(index)
        if node_index != 0:
            node_sequence.append(cluster_sites[node_index - 1])
        index = solution.Value(routing.NextVar(index))
        
    # 5. USER PREFERENCE: Monotonic Sequence (Closest to Furthest)
    node_sequence = sorted(node_sequence, key=lambda x: x['dist_to_wh'])
    
    # 6. Calculate legs for the final response
    legs = []
    prev_node_idx = 0
    for i, site in enumerate(node_sequence):
        # Index in dist_matrix is the original order of chunk + 1
        curr_node_idx = cluster_sites.index(site) + 1
        dist_m = dist_matrix[prev_node_idx][curr_node_idx]
        dist_km = dist_m / 1000.0
        
        if i == 0:
            dist_km = max(0, dist_km - 50)
            
        legs.append({
            "site": site,
            "dist_km": round(dist_km, 2)
        })
        prev_node_idx = curr_node_idx
        
    return legs

def main():
    if len(sys.argv) < 5:
        print(json.dumps({"error": "Missing arguments"}))
        return

    file_path, origin_lat, origin_lng, api_key, output_path = sys.argv[1:6]
    warehouse_coords = (float(origin_lat), float(origin_lng))

    try:
        # 1. Load Data
        df = pd.read_excel(file_path) if file_path.endswith(('.xlsx', '.xls')) else pd.read_csv(file_path)
        
        # 2. Extract Sites
        lat_col = next((c for c in df.columns if c.strip().lower() in ['latitude', 'lat', 'lat ']), None)
        lng_col = next((c for c in df.columns if c.strip().lower() in ['longitude', 'lng', 'lon', 'long']), None)
        id_col = next((c for c in df.columns if c.strip().lower() in ['site id', 'site_id', 'siteid', 'enbsiteid']), None)

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
                        "angle": angle
                    })
            except: continue

        if not site_data:
            print(json.dumps({"error": "No valid sites found"}))
            return

        # 3. Phase 1: Angular Chunking (Strict 3 Groups)
        # We sort by angle to Warehouse to create "pie-slice" sectors.
        # This ensures zonal integrity while guaranteeing exactly 3 sites per route 
        # (until the final remainder group).
        site_data.sort(key=lambda x: x['angle'])
        
        final_routes = []
        for i in range(0, len(site_data), 3):
            chunk = site_data[i : i + 3]
            route_legs = solve_tsp_for_cluster(warehouse_coords, chunk, api_key)
            final_routes.append(route_legs)

        # 4. Finalize Results
        df['CLUBBING'] = ""
        df['AKTBC'] = 0.0
        routes_json = []

        for r_idx, route in enumerate(final_routes):
            label = chr(65 + r_idx)
            route_obj = {"routeNumber": r_idx + 1, "label": label, "legs": []}
            
            for s_idx, leg in enumerate(route):
                site = leg['site']
                df.at[site['orig_idx'], 'CLUBBING'] = f"{label}{s_idx + 1}"
                df.at[site['orig_idx'], 'AKTBC'] = leg['dist_km']
                
                route_obj["legs"].append({
                    "routeLabel": label,
                    "stopSequence": s_idx + 1,
                    "distanceKm": leg['dist_km'],
                    "site": {"id": site['id'], "lat": site['coords'][0], "lng": site['coords'][1]}
                })
            routes_json.append(route_obj)

        df['sort_key'] = df['CLUBBING'].apply(lambda x: x if x else "ZZZ")
        df = df.sort_values(by='sort_key').drop(columns=['sort_key'])
        df.to_excel(output_path, index=False)
        
        print(json.dumps({"success": True, "num_routes": len(final_routes), "routes": routes_json}))

    except Exception as e:
        print(json.dumps({"error": str(e)}))

if __name__ == "__main__":
    main()
