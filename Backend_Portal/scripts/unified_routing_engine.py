import os
import sys
import pandas as pd
import numpy as np
import googlemaps
import json
import math
import itertools
from datetime import datetime

# Ensure we can import from route_optimizer
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from route_optimizer import (
    get_wh_coords, 
    get_api_driving_distance, 
    run_routing, 
    haversine,
    optimize_segment,
    segment_to_legs
)

def process_billing(file_path, api_key, output_path):
    sys.stderr.write(f"Starting Unified Routing Engine for: {file_path}\n")
    
    try:
        gmaps = googlemaps.Client(key=api_key, timeout=10)
    except:
        gmaps = None
        sys.stderr.write("Warning: Running with Haversine fallback (no API key)\n")

    # 1. Load Data
    try:
        df = pd.read_excel(file_path).fillna("")
        # Clean MIN DATE column immediately (Remove timestamps 00:00:00)
        date_col = next((c for c in df.columns if 'date' in c.lower()), 'MIN DATE')
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%-d-%b-%y').fillna('')
    except Exception as e:
        return {"error": f"Failed to load Excel: {str(e)}"}

    # Identify Standard Columns (Robust Mapping)
    col_map = {
        'site_id': next((c for c in df.columns if 'site' in c.lower()), 'ENB SITE ID'),
        'jc': next((c for c in df.columns if 'jc' in c.lower()), 'JC'),
        'cmp': next((c for c in df.columns if 'cmp' in c.lower()), 'CMP'),
        'wh': next((c for c in df.columns if 'wh' in c.lower()), 'WH'),
        'date': next((c for c in df.columns if 'date' in c.lower()), 'MIN DATE'),
        'lat': next((c for c in df.columns if 'lat' in c.lower()), 'LAT '),
        'lng': next((c for c in df.columns if 'long' in c.lower() or 'lng' in c.lower()), 'LONG'),
        'mrn': next((c for c in df.columns if 'mrn' in c.lower()), 'MRN REQD OR NOT'),
        'cap': next((c for c in df.columns if 'cap' in c.lower()), 'KM CAP'),
        'club': next((c for c in df.columns if 'club' in c.lower()), 'CLUBBING'),
        'wh_to_site': 'KM FROM WH TO SITE',
        'site_to_wh': 'KM FROM SITE TO WH',
        'chargeable': 'CHARGEBLE KMS'
    }

    # Ensure output columns exist
    for col in [col_map['wh_to_site'], col_map['site_to_wh'], col_map['chargeable']]:
        if col not in df.columns:
            df[col] = 0.0

    # Local Cache to save API calls for identical WH -> Site pairs
    dist_cache = {}

    def get_cached_dist(origin, dest):
        key = (origin, dest)
        if key not in dist_cache:
            dist_cache[key] = get_api_driving_distance(gmaps, origin, dest)
        return dist_cache[key]

    # Pre-parse types
    df[col_map['mrn']] = df[col_map['mrn']].astype(str).str.strip().str.upper()
    df[col_map['club']] = df[col_map['club']].astype(str).str.strip().str.upper()
    df[col_map['cap']] = pd.to_numeric(df[col_map['cap']], errors='coerce').fillna(50.0)

    # 2. Logic Branching
    # We process in 3 passes: MRN/NR (Direct), Manual Sequences, and Auto-Clustering

    # Pass A: Identify what needs Auto-Clustering
    # Check if single date or multiple dates
    unique_dates = [d for d in df[col_map['date']].unique() if d != ""]
    is_single_date = len(unique_dates) <= 1
    
    # Sites requiring Clustering
    auto_cluster_pool = []
    
    # Process Rows
    for idx, row in df.iterrows():
        mrn_val = row[col_map['mrn']]
        club_val = row[col_map['club']]
        lat, lng = float(row[col_map['lat']]), float(row[col_map['lng']])
        wh_name = row[col_map['wh']]
        wh_coords = get_wh_coords(wh_name)
        site_coords = (lat, lng)
        km_cap = float(row[col_map['cap']])

        # FETCH BASE DISTANCE (WH -> SITE)
        dist_wh = get_cached_dist(wh_coords, site_coords)
        df.at[idx, col_map['wh_to_site']] = dist_wh
        df.at[idx, col_map['site_to_wh']] = dist_wh # For return trip visualization

        # --- LOGIC BRANCHING ---
        
        # BRANCH 1: MRN REQD = YES (Absolute Priority)
        if mrn_val == 'YES':
            df.at[idx, col_map['chargeable']] = max(0, (dist_wh * 2) - km_cap)
            continue

        # BRANCH 2: NO RETURN (NR)
        if club_val == 'NR':
            df.at[idx, col_map['chargeable']] = max(0, dist_wh - km_cap)
            continue

        # BRANCH 3: Pre-filled Clubhouse (A1, A2...)
        if club_val != "" and club_val != "NAN" and club_val != "NONE":
            # This is handled later by sequence logic
            continue

        # BRANCH 4: AUTO-CLUSTERING (Only if NO NR/MRN and CLUBBING is empty)
        if not is_single_date:
            # "If MIN date column comes with more than one date then all sites will simply be wh routed"
            df.at[idx, col_map['chargeable']] = max(0, dist_wh - km_cap)
            df.at[idx, col_map['club']] = "NR"
        else:
            # Single date + Empty clubbing -> Part of the auto-routing pool
            auto_cluster_pool.append({
                'df_idx': idx,
                'coords': site_coords,
                'WH_NAME': wh_name,
                'wh_coords': wh_coords,
                'km_cap': km_cap,
                'row_data': {'INJECTED_JC': row[col_map['jc']], 'CMP': row[col_map['cmp']]}
            })

    # Pass B: Process Auto-Clustering Pool
    if auto_cluster_pool:
        # Group by WH to run optimizer
        # Actually our route_optimizer run_routing expects coords and row_data
        wh_groups = {}
        for s in auto_cluster_pool:
            wn = s['WH_NAME']
            if wn not in wh_groups: wh_groups[wn] = []
            wh_groups[wn].append(s)
            
        for wn, sites in wh_groups.items():
            wh_coords = sites[0]['wh_coords']
            routes = run_routing(wh_coords, sites)
            
            for r_idx, route in enumerate(routes):
                # Convert index to Letter (0 -> A, 1 -> B, etc.)
                route_letter = chr(65 + r_idx) if r_idx < 26 else f"A{chr(65 + r_idx - 26)}"
                
                prev_coords = wh_coords
                for s_idx, leg in enumerate(route):
                    s_idx_1based = s_idx + 1
                    s_data = leg['site']
                    idx = s_data['df_idx']
                    curr_coords = s_data['coords']
                    cap = s_data['km_cap']
                    
                    # New Format: A1, A2, B1...
                    df.at[idx, col_map['club']] = f"{route_letter}{s_idx_1based}"
                    
                    if s_idx_1based == 1:
                        # First stop: WH to Site - CAP
                        df.at[idx, col_map['chargeable']] = max(0, get_cached_dist(wh_coords, curr_coords) - cap)
                    else:
                        # Subsequent stops: Raw distance from previous site
                        df.at[idx, col_map['chargeable']] = get_cached_dist(prev_coords, curr_coords)
                    
                    prev_coords = curr_coords

    # Pass C: Process Manual Sequences (Pre-filled A1, A2...)
    # We group by date and CMP to find sequences
    manual_rows = df[df[col_map['club']].str.match(r'^[A-Z][0-9]+$', na=False)]
    if not manual_rows.empty:
        # Parse Prefix and Sequence
        def parse_club(val):
            import re
            m = re.match(r'([A-Z])([0-9]+)', val)
            if m: return m.group(1), int(m.group(2))
            return val, 0

        for idx, row in manual_rows.iterrows():
            if df.at[idx, col_map['chargeable']] != 0.0: continue # Already handled if MRN/NR somehow overlapped
            
            prefix, seq = parse_club(row[col_map['club']])
            wh_coords = get_wh_coords(row[col_map['wh']])
            curr_coords = (float(row[col_map['lat']]), float(row[col_map['lng']]))
            cap = float(row[col_map['cap']])

            if seq <= 1:
                df.at[idx, col_map['chargeable']] = max(0, get_cached_dist(wh_coords, curr_coords) - cap)
            else:
                # Find previous site in the sequence
                # We look in the same sheet for the same CMP and prefix with seq-1
                prev_club = f"{prefix}{seq-1}"
                cmp_val = row[col_map['cmp']]
                prev_row = df[(df[col_map['cmp']] == cmp_val) & (df[col_map['club']].str.strip() == prev_club)]
                if not prev_row.empty:
                    p_lat, p_lng = float(prev_row.iloc[0][col_map['lat']]), float(prev_row.iloc[0][col_map['lng']])
                    df.at[idx, col_map['chargeable']] = get_cached_dist((p_lat, p_lng), curr_coords)
                else:
                    # Fallback to WH - CAP if prev not found
                    df.at[idx, col_map['chargeable']] = max(0, get_cached_dist(wh_coords, curr_coords) - cap)

    # 3. Export with XlsxWriter Formatting
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='KMS_Billing')
        workbook = writer.book
        worksheet = writer.sheets['KMS_Billing']

        # Formats
        header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#CFE2F3', 'border': 1})
        num_fmt = workbook.add_format({'num_format': '0', 'border': 1, 'align': 'center'})
        std_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        
        # Apply headers
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)

        # Apply data formatting
        for row_num in range(1, len(df) + 1):
            for col_num, col_name in enumerate(df.columns):
                val = df.iloc[row_num-1, col_num]
                # Columns to format as numbers
                if col_name in [col_map['wh_to_site'], col_map['site_to_wh'], col_map['chargeable']]:
                    try: 
                        val = int(round(float(val)))
                    except: 
                        val = 0
                    worksheet.write(row_num, col_num, val, num_fmt)
                else:
                    worksheet.write(row_num, col_num, str(val), std_fmt)

        # Dynamic Column Widths
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 4
            worksheet.set_column(i, i, min(max_len, 40))

    # 2.9 Final Sorting for Excel Order
    def sorting_key(val):
        if val in ["", "NR", "NAN", "NONE"]: return ("ZZZ", 0)
        import re
        m = re.match(r'([A-Z]+)([0-9]+)', val)
        if m:
            return (m.group(1), int(m.group(2)))
        return (val, 0)
    
    df['sort_temp'] = df[col_map['club']].apply(sorting_key)
    df = df.sort_values('sort_temp').drop('sort_temp', axis=1)

    # 3. Final JSON Reconstruction for Website Preview
    routes_json = []
    
    # Filter out NR and empty strings
    plot_df = df[~df[col_map['club']].isin(["", "NR", "NAN", "NONE"])]
    
    if not plot_df.empty:
        # Sort by Route Letter and Sequence
        def get_route_letter(val):
            import re
            m = re.match(r'([A-Z]+)', val)
            return m.group(1) if m else val
            
        def get_seq(val):
            import re
            m = re.search(r'([0-9]+)$', val)
            return int(m.group(1)) if m else 0

        route_letters = plot_df[col_map['club']].apply(get_route_letter).unique()
        
        for r_idx, r_letter in enumerate(route_letters):
            r_group = plot_df[plot_df[col_map['club']].str.startswith(r_letter)].copy()
            r_group['seq_num'] = r_group[col_map['club']].apply(get_seq)
            r_group = r_group.sort_values('seq_num')
            
            first_row = r_group.iloc[0]
            wh_coords = get_wh_coords(first_row[col_map['wh']])
            
            route_obj = {
                "routeNumber": r_idx + 1,
                "label": f"Route {r_letter}",
                "origin_coords": {"lat": wh_coords[0], "lng": wh_coords[1]},
                "legs": []
            }
            
            for _, leg_row in r_group.iterrows():
                route_obj["legs"].append({
                    "routeLabel": leg_row[col_map['club']],
                    "stopSequence": int(leg_row['seq_num']),
                    "distanceKm": int(round(float(leg_row[col_map['chargeable']]))),
                    "site": {
                        "id": str(leg_row[col_map['site_id']]),
                        "lat": float(leg_row[col_map['lat']]),
                        "lng": float(leg_row[col_map['lng']])
                    }
                })
            routes_json.append(route_obj)

    return {"success": True, "output": output_path, "num_routes": len(routes_json), "routes": routes_json}

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: unified_billing_engine.py <input_file> <api_key> [output_path]"}))
        sys.exit(1)

    inp = sys.argv[1]
    key = sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else f"Billing_Result_{datetime.now().strftime('%H%M%S')}.xlsx"
    
    try:
        res = process_billing(inp, key, out)
        # Ensure filenames are relative for the backend to handle
        if "output" in res:
             res["filename"] = os.path.basename(res["output"])
        print(json.dumps(res))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
