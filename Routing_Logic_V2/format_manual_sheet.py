import pandas as pd
import os
import sys

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
MANUAL_FILE     = "KM A6 MMWAVE.xlsx"
ENRICHED_FILE   = "km_required_2_enriched.xlsx"
OUTPUT_FILE     = "comparison_manual_A6.xlsx"
SHEET_NAME      = "Manual Routing - A6"
TARGET_BAND     = "A6 I&C"

HEX_COLORS = ['#F2F2F2', '#FFFFFF'] # Alternating colors for date blocks

def format_manual_routing():
    if not os.path.exists(MANUAL_FILE):
        print(f"Error: {MANUAL_FILE} not found.")
        sys.exit(1)
        
    if not os.path.exists(ENRICHED_FILE):
        print(f"Error: {ENRICHED_FILE} not found. We need this to verify BAND mapping.")
        sys.exit(1)

    print(f"Loading Base Enriched File to build Band Dictionary: {ENRICHED_FILE}...")
    df_enriched = pd.read_excel(ENRICHED_FILE).fillna("")
    
    # Build dictionary mapping: SITE ID -> BAND
    band_map = {}
    for _, row in df_enriched.iterrows():
        site = str(row.get('SITE ID', '')).strip()
        band = str(row.get('BAND', '')).strip()
        if site:
            band_map[site] = band

    print(f"Loaded {len(band_map)} sites from Base Enriched Dictionary.")

    print(f"Loading Employee Raw Routing Sheet: {MANUAL_FILE}...")
    df_manual = pd.read_excel(MANUAL_FILE).fillna("")
    total_raw = len(df_manual)
    
    # ── 1. Map Columns Safely ──
    # The employee sheet is messy, lets map exact matches to our framework.
    column_mapping = {
        'eNBsiteID': 'SITE ID',
        'MIN DATE': 'DATE',
        'JC': 'JC NAME',
        'LAT ': 'LATITUDE',
        'LONG': 'LONGITUDE',
        'CLUBBING': 'CLUBBING NEW',
        'KM FROM WH TO SITE': 'Distance from WH (km)',
        'WAREHOUSE': 'WH',
        'AKTBC': 'AKTBC NEW',
        'PMP ID': 'PMP ID',
        'NO OF SECTOR': 'NO OF\\n SECTOR',
        'Activity': 'Current Status',
        'COMPANY': 'CMP' # Not in the manual sheet usually, but we inject it
    }
    
    manual_data = []

    for _, row in df_manual.iterrows():
        site_id = str(row.get('eNBsiteID', '')).strip()
        if not site_id: continue
        
        # Filter by Band explicitly
        site_band = band_map.get(site_id, "UNKNOWN BAND")
        if site_band != TARGET_BAND:
            continue
            
        # Map values
        mapped_row = {}
        for old_col, new_col in column_mapping.items():
            mapped_row[new_col] = row.get(old_col, "")
            
        # Inject the CMP property correctly using the band map if possible? Or fallback
        # Wait, km_required_2_enriched has CMP. Let's pull it directly.
        cmp_val = str(df_enriched.loc[df_enriched['SITE ID'] == site_id, 'CMP'].values[0]) if site_id in df_enriched['SITE ID'].values else "UNKNOWN CMP"
        mapped_row['CMP'] = cmp_val
        mapped_row['BAND'] = site_band
        
        # Calculate Date Total AKTBC numerically correctly
        def safe_float(val):
            try: return float(val)
            except: return 0.0
            
        mapped_row['AKTBC_NUM'] = safe_float(row.get('AKTBC', 0))
        
        # Fix Date
        date_val = str(row.get('MIN DATE', '')).strip()
        if ' ' in date_val and ':' in date_val:
            date_val = date_val.split()[0]  # strip timestamps
        mapped_row['DATE'] = date_val
        
        manual_data.append(mapped_row)
        
    if not manual_data:
        print(f"Error: No valid {TARGET_BAND} sites found in the manual spreadsheet.")
        sys.exit(1)
        
    # ── 2. Group by Date chronologically ──
    print(f"Found {len(manual_data)} / {total_raw} sites matching {TARGET_BAND}.")
    df_filtered = pd.DataFrame(manual_data)
    
    # Safe date sorting
    df_filtered['DATE'] = pd.to_datetime(df_filtered['DATE'], errors='coerce')
    # fallback na to lowest value to group at top
    df_filtered = df_filtered.sort_values(by=['DATE', 'CLUBBING NEW'], ascending=[True, True])
    # Convert dates back to YYYY-MM-DD
    df_filtered['DATE'] = df_filtered['DATE'].dt.strftime('%Y-%m-%d').fillna("Missing Date")
    
    # Build exact matching structure as Step 4
    all_data = []
    date_blocks = []
    grand_total_aktbc = 0.0
    current_row = 1
    
    for date_str, group in df_filtered.groupby('DATE', sort=False):
        date_total = group['AKTBC_NUM'].sum()
        grand_total_aktbc += date_total
        
        start_r = current_row
        end_r = current_row + len(group) - 1
        
        date_blocks.append({
            'date': date_str,
            'start_r': start_r,
            'end_r': end_r,
            'total': round(date_total, 2)
        })
        
        for _, row in group.iterrows():
            all_data.append(row.to_dict())
            
        current_row += len(group)
        
    # ── 3. Write Excel with Merged Cells ──
    print(f"Generating matched comparison sheet into: {OUTPUT_FILE}...")
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        wb = writer.book
        ws = wb.add_worksheet(SHEET_NAME)
        
        # Exact same formats as final_consolidation.xlsx
        header_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#2F5597', 'font_color': 'white', 'border': 1})
        merged_date_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2', 'rotation': 90})
        merged_total_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#E2EFDA', 'num_format': '0.00'})
        grand_total_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#C6E0B4', 'border': 1, 'font_size': 14})
        grand_total_lbl_fmt = wb.add_format({'bold': True, 'align': 'right', 'valign': 'vcenter', 'bg_color': '#C6E0B4', 'border': 1, 'font_size': 14})
        row_fmts = [
            wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': HEX_COLORS[0]}),
            wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': HEX_COLORS[1]})
        ]
        
        final_cols = [
            'DATE', 'SITE ID', 'PMP ID', 'NO OF\\n SECTOR', 'BAND', 'JC NAME', 'CMP', 
            'Current Status', 'LATITUDE', 'LONGITUDE', 'CLUBBING NEW',
            'Distance from WH (km)', 'WH', 'AKTBC NEW', 'DATE TOTAL AKTBC (km)'
        ]
        
        # 1. Headers
        for c_idx, col_name in enumerate(final_cols):
            ws.write(0, c_idx, str(col_name).replace('\\n', '\n'), header_fmt)
            
        # 2. Rows & Merged cells
        for block_idx, block in enumerate(date_blocks):
            s_row, e_row = block['start_r'], block['end_r']
            fmt = row_fmts[block_idx % 2]
            
            # Merge DATE
            if s_row == e_row: ws.write(s_row, 0, block['date'], merged_date_fmt)
            else: ws.merge_range(s_row, 0, e_row, 0, block['date'], merged_date_fmt)
                
            # Merge TOTAL
            last_col_idx = len(final_cols) - 1
            if s_row == e_row: ws.write(s_row, last_col_idx, block['total'], merged_total_fmt)
            else: ws.merge_range(s_row, last_col_idx, e_row, last_col_idx, block['total'], merged_total_fmt)

            for i in range(s_row - 1, e_row):
                row_data = all_data[i]
                for c_idx, col_name in enumerate(final_cols[1:-1]):
                    val = row_data.get(col_name, "")
                    if col_name == 'AKTBC NEW': val = row_data.get('AKTBC_NUM', val)
                    ws.write(i + 1, c_idx + 1, val, fmt)

        # 3. Grand Total
        grand_total_row = current_row
        ws.merge_range(grand_total_row, 0, grand_total_row, len(final_cols) - 2, "GRAND TOTAL AKTBC ALL DATES (km) [MANUAL EMPLOYEE]", grand_total_lbl_fmt)
        ws.write(grand_total_row, len(final_cols) - 1, grand_total_aktbc, grand_total_fmt)

        # 4. Column Widths
        ws.set_column(0, 0, 12)
        for c_idx, col_name in enumerate(final_cols[1:-1]):
            max_len = max([len(str(r.get(col_name, ""))) for r in all_data], default=0)
            max_len = max(max_len, len(str(col_name))) + 6
            max_len = min(max_len, 40)
            ws.set_column(c_idx + 1, c_idx + 1, max_len)
        ws.set_column(len(final_cols) - 1, len(final_cols) - 1, 24)

    print(f"✅ Executed perfectly. Grand Total Employee AKTBC for A6 I&C: {round(grand_total_aktbc, 2)} km")

if __name__ == "__main__":
    format_manual_routing()
