import pandas as pd
import googlemaps
import os
import sys
import argparse

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
INPUT_FILE  = "km required 2.xlsx"
OUTPUT_FILE = "km_calculated_B6.xlsx"
OUTPUT_SHEET = "B6 - Distances"

# Bands to keep (exact match, case-insensitive strip)
TARGET_BANDS = {"A6+B6 I&C"}

# WH name → coordinates  (read from 'wh location' sheet, but also hardcoded as fallback)
WH_COORDS_FALLBACK = {
    "JAIPUR":  (26.810486, 75.496696),
    "JODHPUR": (26.148422, 73.061378),
}

# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def load_wh_coords(excel_path):
    """Read warehouse coordinates from the 'wh location' sheet."""
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


def calculate_distances(api_key, input_file, output_file):
    print(f"\nLoading data from: {input_file}")

    # ── 1. Load Sheet1 ──
    try:
        df = pd.read_excel(input_file, sheet_name='Sheet1')
        print(f"Total rows in sheet: {len(df)}")
    except Exception as e:
        print(f"Error reading Excel: {e}")
        sys.exit(1)

    # ── 2. Filter bands ──
    df['_band_clean'] = df['BAND'].astype(str).str.strip()
    mask = df['_band_clean'].isin(TARGET_BANDS)
    df_filtered = df[mask].copy().reset_index(drop=True)
    df_filtered.drop(columns=['_band_clean'], inplace=True)
    print(f"Rows after filtering (B6 Sites): {len(df_filtered)}")
    print(f"  BAND breakdown:\n{df_filtered['BAND'].value_counts().to_string()}")

    # ── 3. Load WH coordinates ──
    wh_coords = load_wh_coords(input_file)

    # ── 4. Format date columns as readable strings ──
    for col in df_filtered.columns:
        if pd.api.types.is_datetime64_any_dtype(df_filtered[col]) \
                or 'DATE' in str(col).upper() \
                or 'ATP' in str(col).upper():
            df_filtered[col] = pd.to_datetime(df_filtered[col], errors='coerce') \
                                  .dt.strftime('%d/%m/%Y') \
                                  .replace('NaT', '')

    # ── 5. Google Maps client ──
    gmaps = googlemaps.Client(key=api_key, timeout=10, retry_over_query_limit=True)

    # ── 6. Calculate distance for each row ──
    distances = []
    total = len(df_filtered)

    for idx, row in df_filtered.iterrows():
        site_id = row.get('SITE ID', f"Row {idx}")
        lat     = row.get('LATITUDE')
        lng     = row.get('LONGITUDE')
        wh_name = str(row.get('WH', '')).strip().upper()

        # Missing coordinates
        if pd.isna(lat) or pd.isna(lng):
            print(f"  [{idx+1}/{total}] SKIP {site_id} — missing coordinates")
            distances.append("N/A")
            continue

        # Resolve WH coords
        if wh_name not in wh_coords:
            print(f"  [{idx+1}/{total}] SKIP {site_id} — unknown WH '{wh_name}'")
            distances.append("N/A")
            continue

        origin = wh_coords[wh_name]
        dest   = (lat, lng)

        try:
            # Use Directions API with alternatives=True to get ALL possible routes,
            # then pick the one with the MINIMUM total distance (not time).
            routes = gmaps.directions(
                origin,
                dest,
                mode='driving',
                alternatives=True   # request all available routes
            )

            if routes:
                # Sum up all leg distances for each route and pick the shortest
                def total_dist(route):
                    return sum(leg['distance']['value'] for leg in route['legs'])

                shortest_route = min(routes, key=total_dist)
                dist_m  = total_dist(shortest_route)
                # Round to the nearest whole integer
                dist_km = int(round(dist_m / 1000.0))
                n_routes = len(routes)
                print(f"  [{idx+1}/{total}] {site_id} | WH: {wh_name} | "
                      f"{n_routes} route(s) found → shortest = {dist_km} km")
            else:
                dist_km = "N/A"
                print(f"  [{idx+1}/{total}] {site_id} | No routes returned by API")

        except Exception as e:
            dist_km = "N/A"
            print(f"  [{idx+1}/{total}] {site_id} | API error: {e}")

        distances.append(dist_km)

    # ── 7. Insert 'Distance from WH (km)' after LONGITUDE ──
    df_filtered['Distance from WH (km)'] = distances

    cols = list(df_filtered.columns)
    if 'LONGITUDE' in cols:
        cols.remove('Distance from WH (km)')
        lng_idx = cols.index('LONGITUDE')
        cols.insert(lng_idx + 1, 'Distance from WH (km)')
        df_filtered = df_filtered[cols]

    # ── 8. Write to Excel with formatting ──
    print(f"\nSaving output to: {output_file}")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_filtered.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)

        wb  = writer.book
        ws  = writer.sheets[OUTPUT_SHEET]

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
            'border': 1, 'num_format': '0.00',
            'fg_color': '#FFF2CC'  # light yellow highlight for the new column
        })

        for i, col in enumerate(df_filtered.columns):
            max_len = max(
                df_filtered[col].astype(str).map(len).max(),
                len(str(col))
            ) + 4
            max_len = max(min(max_len, 50), 15)

            if col == 'Distance from WH (km)':
                ws.set_column(i, i, max_len, dist_col_fmt)
            elif 'km' in str(col).lower():
                ws.set_column(i, i, max_len, num_fmt)
            else:
                ws.set_column(i, i, max_len, cell_fmt)

        # Rewrite headers with header format
        for col_num, value in enumerate(df_filtered.columns.values):
            ws.write(0, col_num, value, header_fmt)

    print(f"\n✅ Done! {len(df_filtered)} sites processed → '{output_file}' (sheet: '{OUTPUT_SHEET}')")


# ──────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Calculate driving distance from each site's assigned WH (B6 sites only)."
    )
    parser.add_argument("api_key", help="Google Maps Distance Matrix API Key")
    parser.add_argument("--input",  default=INPUT_FILE,  help="Input Excel file")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Output Excel file")

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' not found.")
        sys.exit(1)

    calculate_distances(args.api_key, args.input, args.output)
