import pandas as pd
import xlsxwriter
import sys
import argparse
import os

def safe_float(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().upper() in ["NR", "NA", "N.A", "-"]:
        return 0.0
    try:
        return float(val)
    except:
        return 0.0

def load_master_data(master_path, dc_number):
    try:
        df_full = pd.read_excel(master_path, header=None)
        if df_full.empty: return None, None
            
        dc_col_idx = 0
        raw_row1 = df_full.iloc[1].tolist() if len(df_full) > 1 else []
        for i, h in enumerate(raw_row1):
            h_str = str(h).upper().strip()
            if "BILLING FILE" in h_str or "DC NUMBER" in h_str:
                dc_col_idx = i
                break

        df_sites = df_full[df_full.iloc[:, dc_col_idx].astype(str).str.strip().str.upper() == dc_number.upper()].copy()
        if df_sites.empty: return None, None
            
        raw_headers = df_full.iloc[1].tolist()
        df_sites.columns = [str(h).strip() for h in raw_headers]
        
        code_to_col_idx = {}
        row0 = df_full.iloc[0].tolist()
        row1 = df_full.iloc[1].tolist()
        for i in range(len(row0)):
            sap = str(row0[i]).split('.')[0].strip()
            if sap and sap != 'nan': code_to_col_idx[sap] = i
            desc = str(row1[i]).strip()
            if desc and desc != 'nan': code_to_col_idx[desc] = i
                
        return df_sites, code_to_col_idx
    except Exception as e:
        print(f"Error loading Master: {e}")
        return None, None

# Template structure for Matrix
TEMPLATE_ITEMS = [
  {"sap": "3367489", "desc": "CHRG BASE RADIO INST&COMS AZIMUTH", "uom": "EA", "rate": 200},
  {"sap": "3367548", "desc": "CHRG LAYING-MULTIMODE FIBER CABLE", "uom": "M", "rate": 5},
  {"sap": "3137158", "desc": "LAYING-2CX2.5SQMM POWER CABLE", "uom": "M", "rate": 12},
  {"sap": "3397253", "desc": "CHRG EXTRA TRANSPORT. > 50 KM (PICKUP", "uom": "KM", "rate": 20},
  {"sap": "3397271", "desc": "CHRG TRANSPORT-UPTO 50KM-SCV-2-3SITE", "uom": "EA", "rate": 1500},
  {"sap": "3367713", "desc": "CHRG ATP-11C - WIFI A6 DEPLOYMENT", "uom": "EA", "rate": 700},
  {"sap": "3367739", "desc": "CHRG ATP-11A - WIFI A6 DEPLOYMENT", "uom": "EA", "rate": 700},
  {"sap": "3317347", "desc": "SITC 1CX6 SQMM CU CBL YY FRLSH", "uom": "M", "rate": 68},
  {"sap": "3383067", "desc": "CHRG APPLY PUFF SEALENT AT CABLE ENTRY", "uom": "EA", "rate": 330},
  {"sap": "3269867", "desc": "TERMINATION OF CABLE 1CX6 SQMM Y", "uom": "M", "rate": 40},
  {"sap": "EXTRA VISIT", "desc": "EXTRA VISIT", "uom": "EA", "rate": 1000},
  {"sap": "POLE MOUNT", "desc": "POLE MOUNT", "uom": "EA", "rate": 500}
]

def write_main_wcc(wb, df_sites, dc_number, formats):
    ws = wb.add_worksheet('Main WCC')
    ws.set_column('A:Z', 15)
    
    ws.merge_range('C4:G6', 'Work Completion Certificate', formats['title'])
    ws.write('D10', 'DC Number:', formats['bold_right'])
    ws.write('E10', dc_number, formats['bold_left'])
    
    num_sites = len(df_sites)
    date_col = 'Completion Date ' if 'Completion Date ' in df_sites.columns else 'Completion Date'
    date_range = "N/A"
    if date_col in df_sites.columns:
        dates = pd.to_datetime(df_sites[date_col], errors='coerce')
        min_date = dates.min().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"
        max_date = dates.max().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"
        date_range = f"{min_date} TO {max_date}"

    ws.write('D12', 'No of Sites:', formats['bold_right'])
    ws.write('E12', f"{num_sites} SITES", formats['bold_left'])
    
    ws.write('D14', 'Date Range:', formats['bold_right'])
    ws.write('E14', date_range, formats['bold_left'])

def write_wcc(wb, df_sites, dc_number, formats):
    ws = wb.add_worksheet('WCC')
    headers = [
        'Sr. No', 'ENB SITE ID', 'PMP SAP ID', 'GIS SECTOR_ID', 'No of Sectors', 
        'Tower type', 'JC', 'WH', 'VEHICLE NO', 'MIN NO', 'MIN Date', 
        'Completion Date', 'REMARKS', 'ACTUAL KM', 'KM-50', 'KM IN WO', 
        'A6 in wo', 'cpri in wo', 'power in wo', 'puff sealant in wo', 
        'termination in wo', 'EXTRA VISIT IN WO', 'Polemount in wo', 'GAP', 'USED KM IN WCC'
    ]
    
    ws.merge_range(0, 0, 0, len(headers)-1, f"WCC - {dc_number}", formats['title'])
    for col_num, header in enumerate(headers):
        ws.write(1, col_num, header, formats['header'])
        ws.set_column(col_num, col_num, 15)

    def get_val(row, matcher):
        c_name = next((c for c in df_sites.columns if matcher.upper() in c.upper()), None)
        return row[c_name] if c_name else ""
    
    aktbc_col = next((c for c in df_sites.columns if 'CHRG EXTRA TRANSPORT' in c.upper() or 'AKTBC' == c.upper()), None)

    for i, (_, row) in enumerate(df_sites.iterrows()):
        r_idx = i + 2
        act_km = safe_float(row[aktbc_col]) if aktbc_col else 0.0
        wo_km = safe_float(get_val(row, 'KM IN WO'))
        
        vals = [
            i + 1, get_val(row, 'ENBSITEID'), get_val(row, 'PMP ID'), get_val(row, 'GIS SECTOR'),
            get_val(row, 'NO OF SECTOR'), get_val(row, 'Tower type'), get_val(row, 'JC'),
            get_val(row, 'WH'), get_val(row, 'VEHICLE NO'), get_val(row, 'MIN NO'),
            get_val(row, 'MIN DATE'), get_val(row, 'Completion Date'), 
            "RFS DONE" if pd.notna(get_val(row, 'Completion Date')) and str(get_val(row, 'Completion Date')) != "" else "",
            act_km, safe_float(get_val(row, 'KM-50(for a6+b6-100)')), wo_km,
            safe_float(get_val(row, 'A6 in wo')), safe_float(get_val(row, 'cpri in wo')),
            safe_float(get_val(row, 'power in wo')), safe_float(get_val(row, 'puff sealant in wo')),
            safe_float(get_val(row, 'termination in wo')), safe_float(get_val(row, 'EXTRA VISIT IN WO')),
            safe_float(get_val(row, 'Polemount in wo')), act_km - wo_km, act_km if act_km <= wo_km else wo_km
        ]
        
        for col_num, val in enumerate(vals):
            # Special formatting for dates
            if isinstance(val, pd.Timestamp):
                ws.write_datetime(r_idx, col_num, val, formats['date'])
            elif isinstance(val, (int, float)):
                ws.write_number(r_idx, col_num, val, formats['number'])
            else:
                ws.write(r_idx, col_num, str(val), formats['cell'])

def write_matrix_sheet(wb, sheet_name, df_sites, code_to_col_idx, dc_number, formats, include_amounts=True):
    ws = wb.add_worksheet(sheet_name)
    num_sites = len(df_sites)
    
    ws.merge_range(0, 0, 0, num_sites + 6, f"{sheet_name} - {dc_number}", formats['title'])
    
    ws.write(2, 0, "SAP Code", formats['header'])
    ws.write(2, 1, "Material Description", formats['header'])
    ws.write(2, 2, "UOM", formats['header'])
    ws.set_column(0, 0, 15)
    ws.set_column(1, 1, 40)
    ws.set_column(2, 2, 8)
    
    # Write Site Headers
    for i, (_, row) in enumerate(df_sites.iterrows()):
        col = 3 + i
        ws.write(2, col, str(row.get('PMP ID', '')).strip(), formats['header_vertical'])
        ws.set_column(col, col, 8)
        
    tot_col = 3 + num_sites
    ws.write(2, tot_col, "Total Qty", formats['header'])
    ws.set_column(tot_col, tot_col, 12)
    
    if include_amounts:
        ws.write(2, tot_col + 1, "Rate", formats['header'])
        ws.write(2, tot_col + 2, "Amount", formats['header'])
        ws.set_column(tot_col + 1, tot_col + 2, 12)

    r_idx = 3
    for item in TEMPLATE_ITEMS:
        ws.write(r_idx, 0, item['sap'], formats['cell'])
        ws.write(r_idx, 1, item['desc'], formats['cell_left'])
        ws.write(r_idx, 2, item['uom'], formats['cell'])
        
        row_sum = 0
        for i, (_, site_row) in enumerate(df_sites.iterrows()):
            col = 3 + i
            sap_code = item['sap']
            val = site_row.iloc[code_to_col_idx[sap_code]] if sap_code in code_to_col_idx else 0.0
            
            # Use raw extraction or 0.0. Exception for Extra Visit / Pole Mount
            if sap_code == "EXTRA VISIT":
                val = safe_float(site_row.get('EXTRA VISIT IN WO', 0.0))
            elif sap_code == "POLE MOUNT":
                val = safe_float(site_row.get('Polemount in wo', 0.0))
            else:
                val = safe_float(val)
                
            ws.write(r_idx, col, val, formats['number'])
            row_sum += val
            
        ws.write_formula(r_idx, tot_col, f"=SUM({xlsxwriter.utility.xl_col_to_name(3)}{r_idx+1}:{xlsxwriter.utility.xl_col_to_name(tot_col-1)}{r_idx+1})", formats['number_bold'], row_sum)
        
        if include_amounts:
            rate = safe_float(item['rate'])
            ws.write(r_idx, tot_col + 1, rate, formats['number'])
            ws.write_formula(r_idx, tot_col + 2, f"={xlsxwriter.utility.xl_col_to_name(tot_col)}{r_idx+1}*{xlsxwriter.utility.xl_col_to_name(tot_col+1)}{r_idx+1}", formats['number_bold'], row_sum * rate)
            
        r_idx += 1
        
    if include_amounts:
        ws.write(r_idx, tot_col + 1, "GRAND TOTAL", formats['bold_right'])
        ws.write_formula(r_idx, tot_col + 2, f"=SUM({xlsxwriter.utility.xl_col_to_name(tot_col+2)}4:{xlsxwriter.utility.xl_col_to_name(tot_col+2)}{r_idx})", formats['number_bold'])

def write_declaration(wb, df_sites, dc_number, formats):
    ws = wb.add_worksheet('Declaration')
    ws.set_column('B:E', 25)
    ws.write('B4', f"DECLARATION FOR {len(df_sites)} SITES", formats['title'])
    ws.write('B6', f"DC NUMBER: {dc_number}", formats['bold_left'])

def write_annexure_and_reco(wb, df_sites, dc_number, formats, mindump_path):
    ws_ann = wb.add_worksheet('Annexture')
    ws_rec = wb.add_worksheet('Reco')
    
    pmp_ids = df_sites['PMP ID'].astype(str).str.strip().tolist()
    
    try:
        if mindump_path and os.path.exists(mindump_path):
            df_dump = pd.read_excel(mindump_path)
        else:
            df_dump = pd.DataFrame()
            
        if not df_dump.empty:
            df_dump['Site ID'] = df_dump['Site ID'].astype(str).str.strip()
            df_all_snapshots = []
            for pid in pmp_ids:
                df_site = df_dump[df_dump['Site ID'] == pid]
                if not df_site.empty:
                    latest_date = df_site['Date'].max()
                    df_all_snapshots.append(df_site[df_site['Date'] == latest_date])
            
            if df_all_snapshots:
                df_filtered = pd.concat(df_all_snapshots)
            else:
                df_filtered = pd.DataFrame()
        else:
            df_filtered = pd.DataFrame()
    except Exception as e:
        print(f"MINDUMP Error: {e}")
        df_filtered = pd.DataFrame()

    if df_filtered.empty:
        pivot = pd.DataFrame(columns=pmp_ids)
    else:
        pivot = df_filtered.pivot_table(index=['SAP Code', 'Material Description'], 
                                       columns='Site ID', values='No. Of Qty', aggfunc='sum').fillna(0)
                                       
    # ANNEXURE
    ws_ann.merge_range(0, 0, 0, len(pmp_ids) + 2, f"Annexture - {dc_number}", formats['title'])
    ws_ann.write(2, 0, "SAP Code", formats['header'])
    ws_ann.write(2, 1, "Material Description", formats['header'])
    ws_ann.set_column(0, 0, 15)
    ws_ann.set_column(1, 1, 50)
    
    for i, pmp_id in enumerate(pmp_ids):
        col = 2 + i
        ws_ann.write(2, col, pmp_id, formats['header_vertical'])
        ws_ann.set_column(col, col, 8)
        
    tot_col = 2 + len(pmp_ids)
    ws_ann.write(2, tot_col, "GRAND TOTAL", formats['header'])
    ws_ann.set_column(tot_col, tot_col, 15)
    
    r_idx = 3
    reco_items = []
    for (sap_code, mat_desc), row_vals in pivot.iterrows():
        ws_ann.write(r_idx, 0, str(sap_code), formats['cell'])
        ws_ann.write(r_idx, 1, str(mat_desc), formats['cell_left'])
        
        for i, pmp_id in enumerate(pmp_ids):
            col = 2 + i
            q = float(row_vals.get(pmp_id, 0))
            ws_ann.write(r_idx, col, q, formats['number'])
            
        ws_ann.write_formula(r_idx, tot_col, f"=SUM({xlsxwriter.utility.xl_col_to_name(2)}{r_idx+1}:{xlsxwriter.utility.xl_col_to_name(tot_col-1)}{r_idx+1})", formats['number_bold'])
        reco_items.append({"sap": str(sap_code), "desc": str(mat_desc), "annexure_row": r_idx + 1})
        r_idx += 1

    # RECO
    ws_rec.merge_range('A1:D2', f"Reconciliation Sheet - {dc_number}", formats['title'])
    ws_rec.write(4, 0, "Material Description", formats['bold_left'])
    ws_rec.write(5, 0, "SAP Code", formats['bold_left'])
    ws_rec.write(6, 0, "Consumption as per Annexure", formats['bold_left'])
    ws_rec.set_column(0, 0, 40)
    
    for i, item in enumerate(reco_items):
        col = 1 + i
        cl = xlsxwriter.utility.xl_col_to_name(col)
        ws_rec.write(4, col, item['desc'], formats['cell'])
        ws_rec.write(5, col, item['sap'], formats['cell'])
        ws_rec.write_formula(6, col, f"=Annexture!{xlsxwriter.utility.xl_col_to_name(tot_col)}{item['annexure_row']}", formats['number'])
        ws_rec.set_column(col, col, 15)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("master_path")
    parser.add_argument("dc_number")
    parser.add_argument("--output", default=None)
    parser.add_argument("--mindump", default=None)
    # Ignored legacy flags
    parser.add_argument("--template", default=None) 
    args = parser.parse_args()

    dc_number = args.dc_number
    output_path = args.output if args.output else f"Billing/{dc_number}_Unified_Billing.xlsx"
    
    print(f"--- Generating Clean Billing Workbook for {dc_number} ---")
    df_sites, code_to_col_idx = load_master_data(args.master_path, dc_number)
    
    if df_sites is not None and not df_sites.empty:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        
        with xlsxwriter.Workbook(output_path, {'nan_inf_to_errors': True}) as wb:
            formats = {
                'title': wb.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#2F5597', 'font_color': 'white', 'border': 1}),
                'header': wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9E1F2', 'border': 1}),
                'header_vertical': wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'rotation': 90, 'bg_color': '#D9E1F2', 'border': 1}),
                'cell': wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
                'cell_left': wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1}),
                'number': wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0.00'}),
                'number_bold': wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0.00', 'bg_color': '#F2F2F2'}),
                'bold_right': wb.add_format({'bold': True, 'align': 'right', 'valign': 'vcenter'}),
                'bold_left': wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter'}),
                'date': wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'dd-mmm-yy'})
            }
            
            write_main_wcc(wb, df_sites, dc_number, formats)
            write_wcc(wb, df_sites, dc_number, formats)
            write_matrix_sheet(wb, 'JMS', df_sites, code_to_col_idx, dc_number, formats, include_amounts=True)
            write_matrix_sheet(wb, 'Abstract', df_sites, code_to_col_idx, dc_number, formats, include_amounts=True)
            write_matrix_sheet(wb, 'BOQ', df_sites, code_to_col_idx, dc_number, formats, include_amounts=True)
            write_declaration(wb, df_sites, dc_number, formats)
            write_annexure_and_reco(wb, df_sites, dc_number, formats, args.mindump)

        print(f"COMPLETE: {output_path}")
    else:
        print("ERROR: No valid data found for DC Number.")

if __name__ == "__main__":
    main()
