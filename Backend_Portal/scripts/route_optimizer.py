import sys
import os
import pandas as pd
import numpy as np
import requests
import json
from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp
import openpyxl
from math import ceil

def get_distance_matrix(locations, api_key):
    """
    Fetch full distance matrix from Google Maps API in parallel batches.
    Uses 5x5 batches (25 elements) to stay under strict API limits.
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
        
        url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={origin_str}&destinations={dest_str}&key={api_key}"
        
        # Debug log
        sys.stderr.write(f"DEBUG: Parallel Fetching {start_i}x{start_j} ({len(origins_batch)}x{len(dest_batch)})\n")
        
        response = requests.get(url).json()
        if response['status'] != 'OK':
            raise Exception(f"Google Maps API Error: {response.get('error_message', response['status'])}")
        
        return (start_i, start_j, response['rows'])

    # Use 10 workers for speed
    with ThreadPoolExecutor(max_workers=10) as executor:
        results = list(executor.map(fetch_batch, batches))

    for start_i, start_j, rows in results:
        for row_idx, row in enumerate(rows):
            for col_idx, element in enumerate(row['elements']):
                if element['status'] == 'OK':
                    matrix[start_i + row_idx][start_j + col_idx] = element['distance']['value']
                else:
                    matrix[start_i + row_idx][start_j + col_idx] = 999999
    
    return matrix

def solve_cvrp(distance_matrix, num_sites, capacity=3):
    """
    Solve Capacitated Vehicle Routing Problem using OR-Tools.
    """
    # Depot is index 0 (Warehouse)
    data = {}
    data['distance_matrix'] = distance_matrix.astype(int).tolist()
    data['demands'] = [0] + [1] * num_sites
    data['num_vehicles'] = ceil(num_sites / capacity) + 2 # Add buffer vehicles
    data['vehicle_capacities'] = [capacity] * data['num_vehicles']
    data['depot'] = 0

    # Create the routing index manager.
    manager = pywrapcp.RoutingIndexManager(len(data['distance_matrix']),
                                           data['num_vehicles'], data['depot'])

    # Create Routing Model.
    routing = pywrapcp.RoutingModel(manager)

    # Create and register a transit callback.
    def distance_callback(from_index, to_index):
        """Returns the distance between the two nodes."""
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        return data['distance_matrix'][from_node][to_node]

    transit_callback_index = routing.RegisterTransitCallback(distance_callback)

    # Define cost of each arc.
    routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

    # Add Capacity constraint.
    def demand_callback(from_index):
        """Returns the demand of the node."""
        from_node = manager.IndexToNode(from_index)
        return data['demands'][from_node]

    demand_callback_index = routing.RegisterUnaryTransitCallback(demand_callback)
    routing.AddDimensionWithVehicleCapacity(
        demand_callback_index,
        0,  # null capacity slack
        data['vehicle_capacities'],  # vehicle maximum capacities
        True,  # start cumul to zero
        'Capacity')

    # Setting first solution heuristic.
    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = (
        routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC)
    search_parameters.local_search_metaheuristic = (
        routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH)
    search_parameters.time_limit.seconds = 10

    # Solve the problem.
    solution = routing.SolveWithParameters(search_parameters)

    if not solution:
        return None

    # Extract routes
    routes = []
    for vehicle_id in range(data['num_vehicles']):
        index = routing.Start(vehicle_id)
        route = []
        while not routing.IsEnd(index):
            node_index = manager.IndexToNode(index)
            if node_index != 0: # Skip depot in the middle
                route.append(node_index)
            index = solution.Value(routing.NextVar(index))
        if route:
            routes.append(route)
    return routes

def main():
    if len(sys.argv) < 5:
        print(json.dumps({"error": "Missing arguments"}))
        return

    file_path = sys.argv[1]
    origin_lat = float(sys.argv[2])
    origin_lng = float(sys.argv[3])
    api_key = sys.argv[4]
    output_path = sys.argv[5]

    try:
        # 1. Load Data
        if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_csv(file_path)

        # 2. Extract Sites with flexible header matching
        def get_col(candidates):
            for c in candidates:
                for col in df.columns:
                    if col.strip().lower() == c.lower():
                        return col
            return None

        lat_col = get_col(['latitude', 'lat', 'lat '])
        lng_col = get_col(['longitude', 'lng', 'lon', 'long'])
        
        if not lat_col or not lng_col:
            print(json.dumps({"error": "Latitude or Longitude columns missing"}))
            return

        # Prepare locations list: Warehouse at index 0, then sites
        locations = [(origin_lat, origin_lng)]
        site_indices = []
        for idx, row in df.iterrows():
            try:
                lat = float(row[lat_col])
                lng = float(row[lng_col])
                if not np.isnan(lat) and not np.isnan(lng):
                    locations.append((lat, lng))
                    site_indices.append(idx)
            except:
                continue

        num_sites = len(locations) - 1
        if num_sites == 0:
            print(json.dumps({"error": "No valid sites found"}))
            return

        # 3. Get Distances
        dist_matrix = get_distance_matrix(locations, api_key)

        # 4. Solve CVRP
        routes_plan = solve_cvrp(dist_matrix, num_sites)
        if not routes_plan:
            print(json.dumps({"error": "Could not find an optimal solution"}))
            return

        # 5. Map back to DataFrame and build JSON for frontend
        df['CLUBBING'] = ""
        df['AKTBC'] = 0.0
        
        id_col = get_col(['site id', 'site_id', 'siteid', 'enbsiteid'])
        
        routes_output = []

        for route_idx, site_ids in enumerate(routes_plan):
            route_label = chr(65 + route_idx) # A, B, C...
            prev_node_idx = 0 # Warehouse
            
            route_obj = {
                "routeNumber": route_idx + 1,
                "label": route_label,
                "legs": []
            }
            
            for seq_idx, node_idx in enumerate(site_ids):
                row_idx = site_indices[node_idx - 1]
                dist_meters = dist_matrix[prev_node_idx][node_idx]
                dist_km = dist_meters / 1000.0
                
                # Special 50km rule for first leg
                if seq_idx == 0:
                    dist_km = max(0, dist_km - 50)
                
                df.at[row_idx, 'CLUBBING'] = f"{route_label}{seq_idx + 1}"
                df.at[row_idx, 'AKTBC'] = round(dist_km, 2)
                
                # Build Leg object for frontend map
                lat, lng = locations[node_idx]
                site_id = str(df.at[row_idx, id_col]) if id_col else str(row_idx)
                
                route_obj["legs"].append({
                    "routeLabel": route_label,
                    "stopSequence": seq_idx + 1,
                    "distanceKm": round(dist_km, 2),
                    "site": {
                        "id": site_id,
                        "lat": lat,
                        "lng": lng
                    }
                })
                
                prev_node_idx = node_idx
            
            routes_output.append(route_obj)

        # Sort by CLUBBING
        df['sort_key'] = df['CLUBBING'].apply(lambda x: x if x else "ZZZ")
        df = df.sort_values(by='sort_key').drop(columns=['sort_key'])

        # 6. Save Excel
        df.to_excel(output_path, index=False)
        
        # 7. Final Output
        print(json.dumps({
            "success": True, 
            "num_routes": len(routes_plan),
            "routes": routes_output
        }))

    except Exception as e:
        print(json.dumps({"error": str(e)}))

if __name__ == "__main__":
    main()
