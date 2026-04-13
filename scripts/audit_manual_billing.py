import os
import pandas as pd
import numpy as np
import sys
import json
import glob

# Ensure we can import from route_optimizer or calculate_manual_distances
sys.path.append(os.path.join(os.getcwd(), 'Backend_Portal', 'scripts'))
from calculate_manual_distances import get_road_distance, get_wh_coords, parse_clubbing

import googlemaps

# API Key - Try to get from environment or use a dummy for local haversine if not available
API_KEY = os.environ.get('Maps_API_KEY', 'dummy_key')
try:
    gmaps = googlemaps.Client(key=API_KEY)
except:
    gmaps = None

def audit_file(file_path):
    print(f"Auditing: {file_path}")
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        return {"file": file_path, "status": f"Error loading: {str(e)}", "mismatches": 0}

    # Identifying cols
    wh_col = next((c for c in df.columns if 'wh' in c.lower() and 'name' not in c.lower()), 'WH')
    lat_col = next((c for c in df.columns if 'lat' in c.lower()), 'LAT ')
    lng_col = next((c for c in df.columns if 'long' in c.lower() or 'lng' in c.lower()), 'LONG')
    club_col = 'CLUBBING'
    
    # Original Manual Distance Cols
    m_wh_col = 'KM FROM WH TO SITE'
    m_extra_col = 'CHRG EXTRA TRANSPORT. > 50 KM (PICKUP'
    
    # New Auto Cols
    a_wh_col = 'AUTO_KM_WH'
    a_extra_col = 'AUTO_EXTRA_TRANSPORT'
    
    results = []
    mismatch_count = 0
    total_sites = 0

    for idx, row in df.iterrows():
        club_val = str(row.get(club_col, '')).strip()
        if not club_val or club_val.lower() == 'nan': continue
        
        total_sites += 1
        prefix, seq = parse_clubbing(club_val)
        site_coords = (float(row[lat_col]), float(row[lng_col]))
        wh_coords = get_wh_coords(row.get(wh_col))
        
        # Deduction logic
        deduction = 100 if 'B6' in str(row.get('Activity', '')).upper() else 50
        
        # 1. Calc Auto WH
        auto_wh = get_road_distance(gmaps, wh_coords, site_coords)
        df.at[idx, a_wh_col] = auto_wh
        
        # 2. Calc Auto Extra
        if seq <= 1:
            auto_extra = max(0, auto_wh - deduction)
        else:
            prev_seq_val = f"{prefix}{seq-1}"
            cmp_val = row.get('CMP')
            prev_row = df[(df['CMP'] == cmp_val) & (df[club_col].astype(str).str.strip() == prev_seq_val)]
            if not prev_row.empty:
                prev_site_coords = (float(prev_row.iloc[0][lat_col]), float(prev_row.iloc[0][lng_col]))
                auto_extra = get_road_distance(gmaps, prev_site_coords, site_coords)
            else:
                auto_extra = max(0, auto_wh - deduction)
        df.at[idx, a_extra_col] = auto_extra

        # Compare
        manual_wh = float(row.get(m_wh_col, 0))
        manual_extra = float(row.get(m_extra_col, 0))
        
        diff_wh = abs(manual_wh - auto_wh)
        diff_extra = abs(manual_extra - auto_extra)
        
        df.at[idx, 'DIFF_WH'] = round(diff_wh, 2)
        df.at[idx, 'DIFF_EXTRA'] = round(diff_extra, 2)
        
        if diff_wh > 5.0 or diff_extra > 5.0:
            df.at[idx, 'AUDIT_RESULT'] = 'REVIEW (Mismatch > 5km)'
            mismatch_count += 1
        else:
            df.at[idx, 'AUDIT_RESULT'] = 'OK (Match)'

    # Re-order columns for Side-by-Side Comparison
    base_cols = ['eNBsiteID', 'Activity', 'JC', 'CMP', 'WH', 'CLUBBING']
    
    # Comparison Sets
    wh_set = ['KM FROM WH TO SITE', a_wh_col, 'DIFF_WH']
    extra_set = ['CHRG EXTRA TRANSPORT. > 50 KM (PICKUP', a_extra_col, 'DIFF_EXTRA']
    status_set = ['AUDIT_RESULT']
    
    # Final column list
    ordered_cols = base_cols + wh_set + extra_set + status_set
    # Add any other missing columns at the end
    other_cols = [c for c in df.columns if c not in ordered_cols]
    df = df[ordered_cols + other_cols]

    # Save Results
    res_path = os.path.join(os.path.dirname(file_path), f"Comparison_Report_{os.path.basename(file_path)}")
    
    writer = pd.ExcelWriter(res_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Comparison')
    workbook = writer.book
    worksheet = writer.sheets['Comparison']
    
    # Formats
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})
    green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
    cell_format = workbook.add_format({'border': 1})
    
    # Auto-width and styling
    df_fmt = df.replace([np.nan, np.inf, -np.inf], '')
    for i, col in enumerate(df.columns):
        series = df[col].astype(str).map(len)
        max_len = max(series.max() if not series.empty else 0, len(col)) + 2
        worksheet.set_column(i, i, min(max_len, 40))
        worksheet.write(0, i, col, header_format)
        
        # Apply data and conditional formatting
        for row_num in range(1, len(df) + 1):
            val = df_fmt.iloc[row_num-1, i]
            result = df.at[row_num-1, 'AUDIT_RESULT']
            
            if col == 'AUDIT_RESULT':
                fmt = red_format if result == 'REVIEW (Mismatch > 5km)' else green_format
                worksheet.write(row_num, i, val, fmt)
            else:
                worksheet.write(row_num, i, val, cell_format)
            
    writer.close()
    
    return {
        "date": os.path.basename(os.path.dirname(file_path)),
        "total_sites": total_sites,
        "mismatches": mismatch_count,
        "success_rate": round(((total_sites - mismatch_count) / total_sites * 100), 2) if total_sites > 0 else 100
    }

def main():
    manual_dir = os.path.join(os.getcwd(), 'SiteRouting', 'manual data')
    folders = [f for f in glob.glob(os.path.join(manual_dir, '*')) if os.path.isdir(f)]
    
    summary_data = []
    for folder in sorted(folders):
        xls_files = glob.glob(os.path.join(folder, 'Routing_Input_*.xlsx'))
        # Avoid auditing already audited files
        xls_files = [f for f in xls_files if 'Audit_Results' not in f]
        for f in xls_files:
            summary_data.append(audit_file(f))
            
    # Write Final Summary
    summary_df = pd.DataFrame(summary_data)
    summary_path = os.path.join(manual_dir, "Audit_Summary_Report.xlsx")
    summary_df.to_excel(summary_path, index=False)
    
    print("\n--- AUDIT SUMMARY ---")
    print(summary_df.to_string())
    print(f"\nOverall Summary saved to: {summary_path}")

if __name__ == "__main__":
    main()
