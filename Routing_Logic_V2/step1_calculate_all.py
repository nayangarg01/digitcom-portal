import pandas as pd
import googlemaps
import os
import sys
import argparse

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
INPUT_FILE  = "km_required_2_enriched.xlsx"
OUTPUT_FILE = "km_calculated_all.xlsx"

BAND_A6 = "A6 I&C"
SHEET_A6 = "A6 I&C - Distances"

BAND_MM = "MM WAVE"
SHEET_MM = "MM WAVE - Distances"

BAND_B6 = "A6+B6 I&C"
SHEET_B6 = "A6+B6 I&C - Distances"

WH_COORDS_FALLBACK = {
    "JAIPUR":  (26.810486, 75.496696),
    "JODHPUR": (26.148422, 73.061378),
}

# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────
def load_wh_coords(excel_path):
    try:
        df = pd.read_excel(excel_path, sheet_name='wh location', header=None)
        coords = {}
        for _, row in df.iterrows():
            name = str(row[0]).strip().upper()
            lat  = float(row[1])
            lng  = float(row[2])
            coords[name] = (lat, lng)
        print(f"Loaded WH coordinates: {coords}")
        return coords
    except Exception as e:
        print(f"Warning: could not read 'wh location' sheet ({e}). Using fallback coords.")
        return WH_COORDS_FALLBACK

def calculate_group_distances(gmaps, df_filtered, wh_coords, group_name):
    print(f"\n--- Processing {group_name} ({len(df_filtered)} sites) ---")
    distances = []
    total = len(df_filtered)

    for idx, row in df_filtered.iterrows():
        site_id = row.get('SITE ID', f"Row {idx}")
        lat     = row.get('LATITUDE')
        lng     = row.get('LONGITUDE')
        wh_name = str(row.get('WH', '')).strip().upper()

        if pd.isna(lat) or pd.isna(lng):
            print(f"  [{idx+1}/{total}] SKIP {site_id} — missing coordinates")
            distances.append("N/A")
            continue

        if wh_name not in wh_coords:
            print(f"  [{idx+1}/{total}] SKIP {site_id} — unknown WH '{wh_name}'")
            distances.append("N/A")
            continue

        origin = wh_coords[wh_name]
        dest   = (lat, lng)

        import math
        def fallback_haversine(p1, p2):
            R = 6371
            lat1, lon1 = math.radians(p1[0]), math.radians(p1[1])
            lat2, lon2 = math.radians(p2[0]), math.radians(p2[1])
            dlat, dlon = lat2 - lat1, lon2 - lon1
            a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2)**2
            c = 2 * math.asin(math.sqrt(a))
            return R * c

        try:
            routes = gmaps.directions(
                origin, dest, mode='driving', alternatives=True
            )
            if routes:
                def total_dist(route):
                    return sum(leg['distance']['value'] for leg in route['legs'])
                
                shortest_route = min(routes, key=total_dist)
                dist_m  = total_dist(shortest_route)
                # Round to nearest whole integer
                dist_km = int(round(dist_m / 1000.0))
                n_routes = len(routes)
                print(f"  [{idx+1}/{total}] {site_id} | WH: {wh_name} | {n_routes} route(s) → shortest = {dist_km} km")
            else:
                dist_km = int(round(fallback_haversine(origin, dest)))
                print(f"  [{idx+1}/{total}] {site_id} | No API routes. Fallback Haversine = {dist_km} km")
        except Exception as e:
            dist_km = int(round(fallback_haversine(origin, dest)))
            print(f"  [{idx+1}/{total}] {site_id} | API error fallback Haversine = {dist_km} km")

        distances.append(dist_km)

    df_filtered['Distance from WH (km)'] = distances
    cols = list(df_filtered.columns)
    if 'LONGITUDE' in cols:
        cols.remove('Distance from WH (km)')
        lng_idx = cols.index('LONGITUDE')
        cols.insert(lng_idx + 1, 'Distance from WH (km)')
        df_filtered = df_filtered[cols]

    return df_filtered

def write_sheet(writer, df_filtered, sheet_name):
    df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)
    wb = writer.book
    ws = writer.sheets[sheet_name]

    header_fmt = wb.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'fg_color': '#D7E4BC', 'border': 1
    })
    cell_fmt = wb.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })
    num_fmt = wb.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1,
        'num_format': '0.00'
    })
    dist_col_fmt = wb.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'border': 1, 'num_format': '0',
        'fg_color': '#FFF2CC'
    })

    for i, col in enumerate(df_filtered.columns):
        max_len = max(df_filtered[col].astype(str).map(len).max(), len(str(col))) + 4
        max_len = max(min(max_len, 50), 15)

        if col == 'Distance from WH (km)':
            ws.set_column(i, i, max_len, dist_col_fmt)
        elif 'km' in str(col).lower():
            ws.set_column(i, i, max_len, num_fmt)
        else:
            ws.set_column(i, i, max_len, cell_fmt)

    # Rewrite headers
    for col_num, value in enumerate(df_filtered.columns.values):
        ws.write(0, col_num, value, header_fmt)


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────
def main(api_key, input_file, output_file):
    print(f"Loading data from: {input_file}")
    try:
        df = pd.read_excel(input_file, sheet_name='Sheet1')
    except Exception as e:
        print(f"Error reading Excel: {e}")
        sys.exit(1)

    df['_band_clean'] = df['BAND'].astype(str).str.strip()
    wh_coords = load_wh_coords(input_file)

    # Format dates strings
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]) or 'DATE' in str(col).upper() or 'ATP' in str(col).upper():
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d').replace('NaT', '')

    # Initialize Gmaps
    gmaps = googlemaps.Client(key=api_key, timeout=10, retry_over_query_limit=True)

    # Band 1: A6
    df_a6 = df[df['_band_clean'] == BAND_A6].copy().reset_index(drop=True)
    df_a6.drop(columns=['_band_clean'], inplace=True, errors='ignore')
    if not df_a6.empty: df_a6_processed = calculate_group_distances(gmaps, df_a6, wh_coords, "A6 I&C")
    else: df_a6_processed = pd.DataFrame()

    # Band 2: MM WAVE
    df_mm = df[df['_band_clean'] == BAND_MM].copy().reset_index(drop=True)
    df_mm.drop(columns=['_band_clean'], inplace=True, errors='ignore')
    if not df_mm.empty: df_mm_processed = calculate_group_distances(gmaps, df_mm, wh_coords, "MM WAVE")
    else: df_mm_processed = pd.DataFrame()

    # Band 3: B6
    df_b6 = df[df['_band_clean'] == BAND_B6].copy().reset_index(drop=True)
    df_b6.drop(columns=['_band_clean'], inplace=True, errors='ignore')
    if not df_b6.empty: df_b6_processed = calculate_group_distances(gmaps, df_b6, wh_coords, "A6+B6 I&C")
    else: df_b6_processed = pd.DataFrame()

    print(f"\nSaving combined output to: {output_file}")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        if not df_a6_processed.empty: write_sheet(writer, df_a6_processed, SHEET_A6)
        if not df_mm_processed.empty: write_sheet(writer, df_mm_processed, SHEET_MM)
        if not df_b6_processed.empty: write_sheet(writer, df_b6_processed, SHEET_B6)

    print(f"✅ Run complete! Exported isolated sheets to '{output_file}'")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Calculate road distance for Group1 & Group2 into one workbook.")
    parser.add_argument("api_key", help="Google Maps Distance Matrix API Key")
    parser.add_argument("--input",  default=INPUT_FILE,  help="Input Excel file")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Output Excel file")
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' not found.")
        sys.exit(1)

    main(args.api_key, args.input, args.output)
