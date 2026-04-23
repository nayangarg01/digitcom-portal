import pandas as pd
import numpy as np
import sys
import os
import math
import googlemaps
import json
import re
import requests
from route_optimizer import get_api_driving_distance

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
    """
    STRICT IMPLEMENTATION: Uses the standardized Routes API (v2) mirror logic.
    """
    if gmaps is None:
        # Fallback to haversine ONLY if gmaps client passed as None (safety check)
        return int(round(haversine(origin, destination)))
        
    try:
        # Calls the unified shortest-visible-path logic
        return get_api_driving_distance(gmaps, origin, destination)
    except Exception as e:
        sys.stderr.write(f"Manual Distance Error: {str(e)}\n")
        # In strict mode, we propagate or return 0, user wants parity.
        return 0

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

def parse_clubbing(val):
    val = str(val).strip()
    if not val or val.lower() == 'nan': return None, 0
    match = re.search(r'([A-Za-z]+)(\d*)', val)
    if match:
        prefix = match.group(1)
        num = int(match.group(2)) if match.group(2) else 1
        return prefix, num
    return val, 1

def get_wh_coords(wh_val):
    wh_val = str(wh_val).upper().strip()
    if 'JLJH' in wh_val or 'JOD' in wh_val: return WH_COORDS['JODHPUR']
    if 'JLKD' in wh_val or 'JAP' in wh_val: return WH_COORDS['JAIPUR']
    return WH_COORDS['DEFAULT']

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
        col2_val = max(0, dist_wh - deduction)
        df.at[idx, col2] = col2_val
        
        # Column 3: Route distance
        if seq <= 1:
            # First or single site: Use DEDUCTED distance for Seq 1
            df.at[idx, col3] = col2_val
        else:
            # Multi-stop: RAW shortest road distance from the previous site (no deduction)
            cmp_val = row.get('CMP')
            prev_seq_val = f"{prefix}{seq-1}"
            prev_row = df[(df['CMP'] == cmp_val) & (df[club_col].astype(str).str.strip() == prev_seq_val)]
            
            if not prev_row.empty:
                prev_site_coords = (float(prev_row.iloc[0][lat_col]), float(prev_row.iloc[0][lng_col]))
                dist_prev = get_road_distance(gmaps, prev_site_coords, site_coords)
                df.at[idx, col3] = dist_prev
            else:
                df.at[idx, col3] = col2_val

    # Save output with Formatting
    output_path = f"Manual_Distance_Result_{os.path.basename(file_path)}"
    output_full_path = os.path.join(os.path.dirname(file_path), output_path)
    
    writer = pd.ExcelWriter(output_full_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Formats
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    
    cell_format = workbook.add_format({'border': 1})
    
    # Handle NaN values for XlsxWriter
    df_clean = df.replace([np.nan, np.inf, -np.inf], '')
    
    # Calculate column widths and write data
    for i, col in enumerate(df.columns):
        series = df[col].astype(str).map(len)
        max_len = max(series.max() if not series.empty else 0, len(col)) + 2
        worksheet.set_column(i, i, min(max_len, 50)) # Cap at 50
        worksheet.write(0, i, col, header_format)
        
        # Apply border to all rows
        for row_num in range(1, len(df) + 1):
             val = df_clean.iloc[row_num-1, i]
             worksheet.write(row_num, i, val, cell_format)

    writer.close()

    print(json.dumps({
        "success": True, 
        "filename": output_path,
        "message": f"Calculated distances for {len(df[df[club_col].notna()])} manual sites."
    }))

if __name__ == "__main__":
    main()
