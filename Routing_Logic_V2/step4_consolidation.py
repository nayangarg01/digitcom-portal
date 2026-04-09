import pandas as pd
import sys
import os
import argparse

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
INPUT_FILE  = "routed_by_date.xlsx"
OUTPUT_FILE = "final_consolidation.xlsx"
HEX_COLORS = ['#F2F2F2', '#FFFFFF']

def generate_consolidation(input_file, output_file):
    if not os.path.exists(input_file):
        print(f"Error: {input_file} not found.")
        sys.exit(1)
        
    print(f"Loading scheduled routes from: {input_file}...")
    xl = pd.ExcelFile(input_file)
    sheets = xl.sheet_names
    
    # We will build distinct structures for each band found in the sheet names
    band_collections = {} # keys: 'A6', 'MM', 'A6+B6', values: { 'data': [], 'blocks': [], 'grand_total': 0.0, 'current_row': 1 }
    
    columns = []

    for sheet_name in sheets:
        if "_" not in sheet_name: continue
        date_str, band_str = sheet_name.rsplit('_', 1)
        
        df_date = pd.read_excel(xl, sheet_name=sheet_name).fillna("")
        if df_date.empty:
            continue
            
        if 'AKTBC NEW' not in df_date.columns:
            print(f"Skipping unrouted sheet: {sheet_name}")
            continue
            
        if not columns:
            columns = list(df_date.columns)
            
        if band_str not in band_collections:
            band_collections[band_str] = {'data': [], 'blocks': [], 'grand_total': 0.0, 'current_row': 1}
            
        coll = band_collections[band_str]
        
        def safe_float(val):
            try: return float(val)
            except: return 0.0
                
        df_date['AKTBC_NUM'] = df_date['AKTBC NEW'].apply(safe_float)
        date_total = df_date['AKTBC_NUM'].sum()
        coll['grand_total'] += date_total
        
        start_r = coll['current_row']
        end_r = coll['current_row'] + len(df_date) - 1
        
        coll['blocks'].append({
            'date': date_str,
            'start_r': start_r,
            'end_r': end_r,
            'total': round(date_total, 2)
        })
        
        for _, row in df_date.iterrows():
            row_dict = {c: row[c] for c in columns}
            row_dict['AKTBC_NUM'] = row['AKTBC_NUM']
            coll['data'].append(row_dict)
            
        coll['current_row'] += len(df_date)

    if not band_collections:
        print("No valid data found to consolidate.")
        sys.exit(0)

    print(f"Generating consolidated sheets into: {output_file}...")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        wb = writer.book
        
        # Formats
        header_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#2F5597', 'font_color': 'white', 'border': 1})
        merged_date_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2', 'rotation': 90})
        merged_total_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#E2EFDA', 'num_format': '0.00'})
        grand_total_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#C6E0B4', 'border': 1, 'font_size': 14})
        grand_total_lbl_fmt = wb.add_format({'bold': True, 'align': 'right', 'valign': 'vcenter', 'bg_color': '#C6E0B4', 'border': 1, 'font_size': 14})
        row_fmts = [
            wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': HEX_COLORS[0]}),
            wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': HEX_COLORS[1]})
        ]
        
        desired_cols = [
            'SITE ID', 'PMP ID', 'NO OF\\n SECTOR', 'BAND', 'JC NAME', 'CMP', 
            'Current Status', 'LATITUDE', 'LONGITUDE', 'CLUBBING NEW',
            'Distance from WH (km)', 'WH', 'AKTBC NEW'
        ]
        
        final_cols = ["DATE"]
        for c in desired_cols:
            if c in columns: final_cols.append(c)
            elif c.replace('\\n', '\n') in columns: final_cols.append(c.replace('\\n', '\n'))
        final_cols.append("DATE TOTAL AKTBC (km)")

        for band_str, coll in band_collections.items():
            sheet_name = f"Master Summary - {band_str}"
            ws = wb.add_worksheet(sheet_name)
            
            for c_idx, col_name in enumerate(final_cols):
                ws.write(0, c_idx, str(col_name).replace('\n', ' '), header_fmt)
                
            for block_idx, block in enumerate(coll['blocks']):
                s_row, e_row = block['start_r'], block['end_r']
                fmt = row_fmts[block_idx % 2]
                
                if s_row == e_row: ws.write(s_row, 0, block['date'], merged_date_fmt)
                else: ws.merge_range(s_row, 0, e_row, 0, block['date'], merged_date_fmt)
                    
                last_col_idx = len(final_cols) - 1
                if s_row == e_row: ws.write(s_row, last_col_idx, block['total'], merged_total_fmt)
                else: ws.merge_range(s_row, last_col_idx, e_row, last_col_idx, block['total'], merged_total_fmt)

                for i in range(s_row - 1, e_row):
                    row_data = coll['data'][i]
                    for c_idx, col_name in enumerate(final_cols[1:-1]):
                        val = row_data.get(col_name, "")
                        if col_name == 'AKTBC NEW': val = row_data.get('AKTBC_NUM', val)
                        ws.write(i + 1, c_idx + 1, val, fmt)

            grand_total_row = coll['current_row']
            ws.merge_range(grand_total_row, 0, grand_total_row, len(final_cols) - 2, "GRAND TOTAL AKTBC ALL DATES (km)", grand_total_lbl_fmt)
            ws.write(grand_total_row, len(final_cols) - 1, coll['grand_total'], grand_total_fmt)

            ws.set_column(0, 0, 12)
            for c_idx, col_name in enumerate(final_cols[1:-1]):
                max_len = max([len(str(r.get(col_name, ""))) for r in coll['data']], default=0)
                max_len = max(max_len, len(str(col_name))) + 6
                max_len = min(max_len, 40)
                ws.set_column(c_idx + 1, c_idx + 1, max_len)
            ws.set_column(len(final_cols) - 1, len(final_cols) - 1, 24)
            
            print(f"✅ Generated {sheet_name} | Subtotal: {round(coll['grand_total'], 2)} km")
            
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Consolidate routed excel into one master sheet.")
    parser.add_argument("--input", default=INPUT_FILE, help="Input routed_by_date.xlsx file")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Output consolidated Excel file")
    
    args = parser.parse_args()
    generate_consolidation(args.input, args.output)
