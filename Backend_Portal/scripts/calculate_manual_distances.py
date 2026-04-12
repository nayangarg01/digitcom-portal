import pandas as pd
import numpy as np
import sys
import os
import math
import googlemaps
import json
import re

# ──────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────
WH_COORDS = {
    "JAIPUR": (26.810486, 75.496696),
    "JODHPUR": (26.148422, 73.061378),
    "DEFAULT": (26.810486, 75.496696)
}

def haversine(coord1, coord2):
    R = 6371.0
    lat1, lon1 = math.radians(coord1[0]), math.radians(coord1[1])
    lat2, lon2 = math.radians(coord2[0]), math.radians(coord2[1])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def get_road_distance(gmaps, origin, destination):
    if gmaps is None:
        return round(haversine(origin, destination), 2)
    try:
        res = gmaps.distance_matrix(origin, destination, mode='driving')
        if res['status'] == 'OK' and res['rows'][0]['elements'][0]['status'] == 'OK':
            return round(res['rows'][0]['elements'][0]['distance']['value'] / 1000.0, 2)
        return round(haversine(origin, destination), 2)
    except Exception as e:
        return round(haversine(origin, destination), 2)

def main():
    if len(sys.argv) < 3:
        print(json.dumps({"success": False, "error": "Missing arguments. Usage: python script.py file_path api_key"}))
        return

    file_path = sys.argv[1]
    api_key = sys.argv[2]
    
    try:
        gmaps = googlemaps.Client(key=api_key)
    except:
        gmaps = None

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}))
        return

    # Column identification
    id_col = next((c for c in df.columns if 'site id' in c.lower() or 'enbsiteid' in c.lower()), 'SiteID')
    lat_col = next((c for c in df.columns if 'lat' in c.lower()), 'LAT ')
    lng_col = next((c for c in df.columns if 'long' in c.lower() or 'lng' in c.lower()), 'LONG')
    wh_col = next((c for c in df.columns if 'wh' in c.lower() and 'name' not in c.lower()), 'WH')
    club_col = 'CLUBBING'
    act_col = next((c for c in df.columns if 'activity' in c.lower()), 'Activity')

    # Target Columns
    col1 = 'KM FROM WH TO SITE'
    col2 = 'KM-50(for a6+b6-100)'
    col3 = 'CHRG EXTRA TRANSPORT. > 50 KM (PICKUP'

    # Ensure columns exist
    for c in [col1, col2, col3]:
        df[c] = np.nan

    # Process by Route
    # We group by the prefix of the clubbing value (e.g., A, B, C)
    def parse_clubbing(val):
        val = str(val).strip()
        if not val or val.lower() == 'nan': return None, 0
        match = re.search(r'([A-Za-z]+)(\d*)', val)
        if match:
            prefix = match.group(1)
            num = int(match.group(2)) if match.group(2) else 1
            return prefix, num
        return val, 1

    # 1. Map Hub Coordinates
    def get_wh_coords(wh_val):
        wh_val = str(wh_val).upper().strip()
        if 'JLJH' in wh_val or 'JOD' in wh_val: return WH_COORDS['JODHPUR']
        if 'JLKD' in wh_val or 'JAP' in wh_val: return WH_COORDS['JAIPUR']
        return WH_COORDS['DEFAULT']

    # We need to process each site
    for idx, row in df.iterrows():
        club_val = row.get(club_col)
        if pd.isna(club_val) or str(club_val).strip() == '': continue
        
        prefix, seq = parse_clubbing(club_val)
        site_coords = (float(row[lat_col]), float(row[lng_col]))
        wh_coords = get_wh_coords(row.get(wh_col))
        
        # Column 1: WH to Site
        dist_wh = get_road_distance(gmaps, wh_coords, site_coords)
        df.at[idx, col1] = dist_wh
        
        # Column 2: Subtract 50 (A6) or 100 (B6)
        deduction = 100 if 'B6' in str(row.get(act_col, '')).upper() else 50
        df.at[idx, col2] = max(0, dist_wh - deduction)
        
        # Column 3: Route distance
        if seq == 1:
            # First or single site
            df.at[idx, col3] = dist_wh
        else:
            # Seek previous site in the same group (CMP + Prefix + seq-1)
            cmp_val = row.get('CMP')
            prev_seq_val = f"{prefix}{seq-1}"
            prev_row = df[(df['CMP'] == cmp_val) & (df[club_col].astype(str).str.strip() == prev_seq_val)]
            
            if not prev_row.empty:
                prev_coords = (float(prev_row.iloc[0][lat_col]), float(prev_row.iloc[0][lng_col]))
                dist_prev = get_road_distance(gmaps, prev_coords, site_coords)
                df.at[idx, col3] = dist_prev
            else:
                # Fallback to WH if precursor not found
                df.at[idx, col3] = dist_wh

    # Save output
    output_path = f"Manual_Distance_Result_{os.path.basename(file_path)}"
    output_full_path = os.path.join(os.path.dirname(file_path), output_path)
    df.to_excel(output_full_path, index=False)

    print(json.dumps({
        "success": True, 
        "filename": output_path,
        "message": f"Calculated distances for {len(df[df[club_col].notna()])} manual sites."
    }))

if __name__ == "__main__":
    main()
