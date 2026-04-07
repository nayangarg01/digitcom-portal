import pandas as pd
import openpyxl
import sys
import argparse
import os
from copy import copy
import datetime

from openpyxl.utils import get_column_letter

def load_master_data(master_path, target_billing_file):
    """Loads the KM SHEET from Master DPR and filters strictly for the target DC code."""
    try:
        df_master = pd.read_excel(master_path, sheet_name='KM SHEET', header=1)
    except Exception as e:
        print(f"Error loading master file {master_path}: {e}")
        return None

    df_master.columns = df_master.columns.astype(str).str.strip()
    
    # Identify BILLING FILE column safely
    bill_col = next((c for c in df_master.columns if 'BILLING FILE' in c.upper()), None)
    if not bill_col:
        print("Could not find 'BILLING FILE' column in master sheet.")
        return None
        
    df_filtered = df_master[df_master[bill_col].astype(str).str.strip().str.upper() == target_billing_file.upper()].copy()
    if df_filtered.empty:
        print(f"Warning: No sites found mapped to {target_billing_file}!")
        return None

    return df_filtered

def generate_wcc_sheet(df_sites, wb):
    """Injects the filtered sites specifically into the WCC Template sheet."""
    if 'WCC' not in wb.sheetnames:
        print("Error: WCC sheet not found in template.")
        return False
        
    ws = wb['WCC']
    
    # Target exactly the explicit string the user confirmed
    aktbc_col = next((c for c in df_sites.columns if 'CHRG EXTRA TRANSPORT' in str(c).upper() or 'AKTBC' == str(c).strip().upper()), None)
    
    # Locate headers in WCC
    header_row_idx = None
    cols_map = {}
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row and 'GIS SECTOR_ID' in str(row):
            header_row_idx = r_idx
            for c_idx, val in enumerate(row, start=1):
                if val:
                    cols_map[str(val).strip()] = c_idx
            break

    if not header_row_idx:
        print("Error: Could not locate visual headers in the WCC template sheet.")
        return False

    start_row = header_row_idx + 1
    
    # Pre-calculate base styling
    base_styles = {}
    for c_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=start_row, column=c_idx)
        base_styles[c_idx] = {
            'font': copy(cell.font),
            'border': copy(cell.border),
            'fill': copy(cell.fill),
            'number_format': cell.number_format,
            'alignment': copy(cell.alignment)
        }

    # Management of rows
    if len(df_sites) > 22:
        ws.insert_rows(start_row + 22, amount=(len(df_sites) - 22))
    elif len(df_sites) < 22:
        amount_to_delete = 22 - len(df_sites)
        ws.delete_rows(start_row + len(df_sites), amount=amount_to_delete)
    
    # Map and Write rows
    for i, (_, row) in enumerate(df_sites.iterrows()):
        curr_row = start_row + i
        
        def get_val(matcher):
            c_name = next((c for c in df_sites.columns if matcher.upper() in c.upper()), None)
            return row[c_name] if c_name else ""
            
        def get_exact(name):
            return row[name] if name in df_sites.columns else ""

        mapping = [
            ('Sr. No', i + 1),
            ('ENB SITE ID', get_val('ENBSITEID') or get_val('SAP ID') or get_exact('Unnamed: 0')),
            ('PMP SAP ID', get_val('PMP ID')),
            ('GIS SECTOR_ID', get_val('GIS SECTOR')),
            ('No of Sectors', get_val('NO OF SECTOR')),
            ('Tower type', get_val('Tower type')),
            ('JC', get_val('JC')),
            ('WH', get_val('WH')),
            ('VEHICLE NO', get_val('VEHICLE NO')),
            ('MIN  NO', get_val('MIN NO')),
            ('MIN Date', get_val('MIN DATE')),
            ('Completion Date', get_val('Completion Date')),
            ('REMARKS', "RFS DONE" if pd.notna(get_val('Completion Date')) and get_val('Completion Date') != "" else ""),
            ('ACTUAL KM', float(row[aktbc_col]) if aktbc_col and pd.notna(row[aktbc_col]) else 0.0),
            ('KM IN WO', float(get_val('KM IN WO')) if pd.notna(get_val('KM IN WO')) and get_val('KM IN WO') != "" else 0.0)
        ]

        # Calculate Gap and Used KM
        actual_km = next((val for k, val in mapping if k == 'ACTUAL KM'), 0.0)
        km_in_wo = next((val for k, val in mapping if k == 'KM IN WO'), 0.0)
        mapping.append(('GAP', actual_km - km_in_wo))
        mapping.append(('USED KM IN WCC', actual_km if actual_km <= km_in_wo else km_in_wo))

        # Styling and value injection
        for c_idx in range(1, ws.max_column + 1):
            c = ws.cell(row=curr_row, column=c_idx)
            styles = base_styles.get(c_idx)
            if styles:
                c.font = copy(styles['font'])
                c.border = copy(styles['border'])
                c.fill = copy(styles['fill'])
                c.number_format = styles['number_format']
                c.alignment = copy(styles['alignment'])

        for col_name, val in mapping:
            c_target = next((cols_map[k] for k in cols_map if col_name.upper().strip() in k.upper().strip()), None)
            if c_target:
                target_cell = ws.cell(row=curr_row, column=c_target)
                if isinstance(val, pd.Timestamp):
                    target_cell.value = val.to_pydatetime()
                else:
                    target_cell.value = val

    # Formulas
    last_row = start_row + len(df_sites) - 1
    summary_row = last_row + 1
    for sum_col in ['ACTUAL KM', 'USED KM IN WCC']:
        c_idx = next((cols_map[k] for k in cols_map if sum_col.upper().strip() in k.upper().strip()), None)
        if c_idx:
            col_let = get_column_letter(c_idx)
            ws.cell(row=summary_row, column=c_idx).value = f"=SUM({col_let}{start_row}:{col_let}{last_row})"

    return True

def generate_jms_sheet(df_sites, wb):
    """Injects the filtered sites into the horizontal matrix JMS Template sheet."""
    if 'JMS' not in wb.sheetnames:
        print("Error: JMS sheet not found in template.")
        return False
        
    ws = wb['JMS']
    
    # 1. Identify key horizontal and vertical tracking indices
    # Row 10: Site IDs (Col 4 onwards)
    SITE_ROW = 10
    SITE_TYPE_ROW = 11
    SECTOR_ROW = 12
    START_COL = 4 # 'D'
    
    # Site mapping
    towers = [str(t).strip() for t in df_sites['Tower type']] if 'Tower type' in df_sites.columns else ["GBT"]*len(df_sites)
    sectors = [v for v in df_sites['NO OF SECTOR']] if 'NO OF SECTOR' in df_sites.columns else [1]*len(df_sites)
    site_ids = [str(s).strip() for s in (df_sites['PMP ID'] if 'PMP ID' in df_sites.columns else df_sites.iloc[:, 0])]

    # Column Management
    if len(df_sites) > 22:
        ws.insert_cols(START_COL + 22, amount=(len(df_sites) - 22))
    elif len(df_sites) < 22:
        amount_to_delete = 22 - len(df_sites)
        ws.delete_cols(START_COL + len(df_sites), amount=amount_to_delete)

    # Pre-copy styles from Column D (Start column)
    col_styles = {}
    for r_idx in range(1, ws.max_row + 1):
        cell = ws.cell(row=r_idx, column=START_COL)
        col_styles[r_idx] = {
            'font': copy(cell.font), 'border': copy(cell.border), 'fill': copy(cell.fill),
            'number_format': cell.number_format, 'alignment': copy(cell.alignment)
        }

    # Write Site Headers and Apply styles across columns
    for i in range(len(df_sites)):
        curr_col = START_COL + i
        # Write headers
        ws.cell(row=SITE_ROW, column=curr_col).value = site_ids[i]
        ws.cell(row=SITE_TYPE_ROW, column=curr_col).value = towers[i]
        ws.cell(row=SECTOR_ROW, column=curr_col).value = sectors[i]
        
        # Apply style to whole column
        for r_idx, s in col_styles.items():
            if r_idx < 10: continue # Skip top header fluff
            c = ws.cell(row=r_idx, column=curr_col)
            c.font, c.border, c.fill = copy(s['font']), copy(s['border']), copy(s['fill'])
            c.number_format, c.alignment = s['number_format'], copy(s['alignment'])

    # Item Quantity Injection
    # Scan Column B for descriptions and match with df_sites columns
    for r_idx in range(14, ws.max_row + 1):
        desc = str(ws.cell(row=r_idx, column=2).value).strip() if ws.cell(row=r_idx, column=2).value else ""
        if not desc or "TOTAL" in desc.upper(): continue
        
        # Fuzzy match description with MASTERDPR columns
        match_col = next((c for c in df_sites.columns if desc.upper() in str(c).upper() or str(c).upper() in desc.upper()), None)
        if match_col:
            quantities = df_sites[match_col].values
            for i, qty in enumerate(quantities):
                ws.cell(row=r_idx, column=START_COL + i).value = float(qty) if pd.notna(qty) else 0.0

    # Summary Column Logic (Total Qty, Amount)
    summary_start_col = None
    for c_idx in range(START_COL + len(df_sites), ws.max_column + 1):
        if "Total Quantity" in str(ws.cell(row=SITE_TYPE_ROW + 1, column=c_idx).value):
            summary_start_col = c_idx
            break
            
    if summary_start_col:
        qty_col = summary_start_col
        rate_col = qty_col + 1
        amt_col = qty_col + 2
        
        first_let = get_column_letter(START_COL)
        last_let = get_column_letter(START_COL + len(df_sites) - 1)
        
        for r_idx in range(16, ws.max_row):
            desc_val = ws.cell(row=r_idx, column=2).value
            if not desc_val or "TOTAL" in str(desc_val).upper():
                if "TOTAL" in str(desc_val).upper():
                    # Update Grand Total
                    sum_formula = f"=SUM({get_column_letter(amt_col)}16:{get_column_letter(amt_col)}{r_idx-1})"
                    ws.cell(row=r_idx, column=amt_col).value = sum_formula
                continue
                
            # Total Quantity = SUM(Sites)
            ws.cell(row=r_idx, column=qty_col).value = f"=SUM({first_let}{r_idx}:{last_let}{r_idx})"
            # Amount = Qty * Rate
            ws.cell(row=r_idx, column=amt_col).value = f"={get_column_letter(qty_col)}{r_idx}*{get_column_letter(rate_col)}{r_idx}"

    return True

def main():
    parser = argparse.ArgumentParser(description="Generate JIO Billing Formats from MASTERDPR.")
    parser.add_argument("master_file", help="Path to the MASTERDPR.xlsx file")
    parser.add_argument("billing_target", help="Target Billing File code (e.g., DC0105)")
    args = parser.parse_args()

    template_path = 'Billing/DC0105.xlsx' 
    output_path = f'Billing/{args.billing_target.upper()}_Automated.xlsx'

    print(f"--- Processing {args.billing_target.upper()} ---")
    df_sites = load_master_data(args.master_file, args.billing_target)
    
    if df_sites is not None:
        print(f"Loaded {len(df_sites)} sites. Opening template...")
        try:
            wb = openpyxl.load_workbook(template_path)
            
            print("Processing WCC Sheet...")
            generate_wcc_sheet(df_sites, wb)
            
            print("Processing JMS Sheet...")
            generate_jms_sheet(df_sites, wb)
            
            wb.save(output_path)
            print(f"Success! Finalized file: {output_path}")
        except Exception as e:
            print(f"Generation failed: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("Generation aborted due to missing data.")

if __name__ == '__main__':
    main()
