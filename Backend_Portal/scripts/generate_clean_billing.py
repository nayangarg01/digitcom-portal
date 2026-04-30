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

def format_date(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    try:
        dt = pd.to_datetime(val)
        return dt.strftime('%d-%b-%y')
    except:
        return str(val)

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
  {"sap": "3397248", "desc": "EXTRA VISIT", "uom": "EA", "rate": 1000},
  {"sap": "3268025", "desc": "INSTALLATION OF POLE MOUNT ON TOWER", "uom": "EA", "rate": 500}
]

def write_main_wcc(wb, df_sites, dc_number, formats):
    ws = wb.add_worksheet('Main WCC')
    ws.set_column('A:A', 5)
    ws.set_column('B:B', 20)
    ws.set_column('C:G', 15)
    ws.set_column('H:H', 30)

    # Calculate dates and sites
    num_sites = len(df_sites)
    date_col = 'Completion Date ' if 'Completion Date ' in df_sites.columns else 'Completion Date'
    date_range = "N/A"
    if date_col in df_sites.columns:
        dates = pd.to_datetime(df_sites[date_col], errors='coerce')
        min_date = dates.min().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"
        max_date = dates.max().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"
        date_range = f"{min_date} TO {max_date}"

    # Draw the static form (Row 1 to 28, Col B to H)
    f_title = wb.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'border': 2})
    f_bold = wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1})
    f_norm = wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    f_center = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})

    ws.merge_range('B2:H3', 'Work Completion Certificate', f_title)

    # Row 4
    ws.write('B5', 'State', f_bold)
    ws.merge_range('C5:D5', 'RAJASTHAN', f_center)
    ws.merge_range('E5:G5', 'Maintenance Point', f_bold)
    ws.write('H5', 'Jaipur', f_center)

    # Row 5 (Project Type)
    ws.write('B7', 'Project Type', f_bold)
    ws.write('C7', '93 K', f_center)
    ws.write('D7', 'Infill', f_center)
    ws.write('E7', 'Growth', f_center)
    ws.merge_range('F7:G7', 'Other (Specify) _________________', f_bold)
    ws.write('H7', 'Air Fiber Installation', f_center)

    # Row 7 (Site Type)
    ws.write('B9', 'Site Type', f_bold)
    ws.write('C9', 'Own Built', f_center)
    ws.write('D9', 'IP Colo', f_center)
    ws.write('E9', 'RP1', f_center)
    ws.write('F9', 'BSNL', f_center)
    ws.write('G9', 'MAG1 NLD AG1', f_center)
    ws.write('H9', 'ZXZ', f_center)

    # Row 9 (Tower type)
    ws.write('B11', 'Tower type', f_bold)
    ws.write('C11', 'GBT', f_center)
    ws.write('D11', 'RTT', f_center)
    ws.write('E11', 'RTP', f_center)
    ws.write('F11', 'GBM', f_center)
    ws.write('G11', 'NBT Other(Specify) ____________', f_bold)
    ws.write('H11', 'Air Fiber Installation', f_center)

    # Row 11 & 12 (Cert text)
    ws.merge_range('B13:H13', 'This is to certify that work has been completed as per specification given in workorder on the sites mentioned', f_center)
    ws.merge_range('B14:H14', 'The required ITP / Checklists are available and verified in system', f_bold)

    # Row 14
    ws.write('B16', 'Site Name', f_bold)
    ws.merge_range('C16:E16', 'As per Annexture', f_center)
    ws.merge_range('F16:G16', 'SAP ID', f_bold)
    ws.write('H16', 'As per Annexture', f_center)

    # Row 16
    ws.write('B18', 'W.O.Number', f_bold)
    ws.merge_range('C18:E18', 'P14/630330726', f_center)
    ws.merge_range('F18:G18', 'Vendor Name  M/S.', f_bold)
    ws.write('H18', 'DIGITCOM INDIA TECHNOLOGIES', f_center)

    # Row 18
    ws.write('B20', 'No of Sites', f_bold)
    ws.merge_range('C20:E20', f'{num_sites} SITES', f_center)
    ws.merge_range('F20:G20', 'Completion Date', f_bold)
    ws.write('H20', date_range, f_center)

    # Row 20
    ws.merge_range('B22:E22', 'Vendor Representative', f_bold)
    ws.merge_range('F22:H22', 'RJIL Representative', f_bold)

    # Row 22
    ws.write('B24', 'Name', f_bold)
    ws.merge_range('C24:E24', 'ANKUSH SRIVASTAVA', f_center)
    ws.merge_range('F24:G24', 'Name', f_bold)
    ws.write('H24', 'MR. Manish Nahar', f_center)

    # Row 24
    ws.write('B26', 'Sign', f_bold)
    ws.merge_range('C26:E26', '', f_norm)
    ws.merge_range('F26:G26', 'Sign', f_bold)
    ws.write('H26', '', f_norm)

    # Row 26
    ws.write('B28', 'Date', f_bold)
    ws.merge_range('C28:E28', '', f_norm)
    ws.merge_range('F28:G28', 'Date', f_bold)
    ws.write('H28', '', f_norm)

    # Row 28
    ws.merge_range('B30:H30', 'Remarks, if any:', f_bold)
    
    # Row 30 (Note)
    ws.write('B32', 'Note :', f_bold)
    ws.merge_range('C32:H32', 'In case of Multiple sites, please attach applicable site details with this certificate', f_norm)

def write_wcc(wb, df_sites, dc_number, formats):
    ws = wb.add_worksheet('WCC')
    
    headers_1 = [
        'Sr. No', 'ENB SITE ID', 'PMP SAP ID', 'GIS SECTOR_ID', 'No of Sectors', 
        'Tower type', 'JC', 'WH', 'VEHICLE NO', 'MIN NO', 'MIN Date', 
        'Completion Date', 'REMARKS'
    ]
    headers_2 = [
        'ACTUAL KM', 'KM IN WO', 'GAP', 'USED KM IN WCC'
    ]
    
    # 1. Main Title
    ws.merge_range('C3:O4', 'Work Completion Certificate', formats['title'])
    
    # 2. Certification Text
    cert_text = "This is to certify that below sites pertaining to WO/WCO No.P14/630330726 Dated in 03-10-2025 respect of Digitcom India Technologies  has  been successfully completed in all respect."
    ws.merge_range('C6:O6', cert_text, formats['cert_text'])
    
    # 3. Write Headers
    r_head = 8
    col_idx = 2  # Start at column C (index 2)
    for h in headers_1:
        ws.write(r_head, col_idx, h, formats['header_blue'])
        if 'ID' in h or 'SECTOR' in h:
            ws.set_column(col_idx, col_idx, 22)
        elif 'Date' in h or 'REMARKS' in h or 'VEHICLE' in h:
            ws.set_column(col_idx, col_idx, 15)
        else:
            ws.set_column(col_idx, col_idx, 10)
        col_idx += 1
        
    col_idx = 16  # Start yellow table at column Q (index 16)
    for h in headers_2:
        ws.write(r_head, col_idx, h, formats['header_yellow'])
        ws.set_column(col_idx, col_idx, 12)
        col_idx += 1

    def get_val(row, matcher):
        c_name = next((c for c in df_sites.columns if matcher.upper() in c.upper()), None)
        return row[c_name] if c_name else ""
    
    aktbc_col = next((c for c in df_sites.columns if 'CHRG EXTRA TRANSPORT' in c.upper() or 'AKTBC' == c.upper()), None)

    r_idx = 9
    total_act = 0
    total_used = 0

    for i, (_, row) in enumerate(df_sites.iterrows()):
        act_km = safe_float(row[aktbc_col]) if aktbc_col else 0.0
        wo_km = safe_float(get_val(row, 'KM IN WO'))
        used_km = act_km if act_km <= wo_km else wo_km
        
        total_act += act_km
        total_used += used_km
        
        vals_1 = [
            i + 1, get_val(row, 'ENBSITEID'), get_val(row, 'PMP ID'), get_val(row, 'GIS SECTOR'),
            safe_float(get_val(row, 'NO OF SECTOR')), get_val(row, 'Tower type'), get_val(row, 'JC'),
            get_val(row, 'WH'), get_val(row, 'VEHICLE NO'), get_val(row, 'MIN NO'),
            format_date(get_val(row, 'MIN DATE')), format_date(get_val(row, 'Completion Date')), 
            "RFS DONE" if pd.notna(get_val(row, 'Completion Date')) and str(get_val(row, 'Completion Date')) != "" else ""
        ]
        
        vals_2 = [
            act_km, wo_km, act_km - wo_km, used_km
        ]
        
        for c, val in enumerate(vals_1):
            c_pos = 2 + c
            if isinstance(val, pd.Timestamp):
                ws.write_datetime(r_idx, c_pos, val, formats['date'])
            elif isinstance(val, (int, float)):
                ws.write_number(r_idx, c_pos, val, formats['number'])
            else:
                ws.write(r_idx, c_pos, str(val), formats['cell'])
                
        for c, val in enumerate(vals_2):
            c_pos = 16 + c
            ws.write_number(r_idx, c_pos, val, formats['number'])
            
        r_idx += 1

    # Totals Row for Yellow Table
    ws.write(r_idx, 16, total_act, formats['header_yellow'])
    ws.write(r_idx, 17, "", formats['header_yellow'])
    ws.write(r_idx, 18, "", formats['header_yellow'])
    ws.write(r_idx, 19, total_used, formats['header_yellow'])

    r_sig = r_idx + 2
    ws.write(r_sig, 3, "SIGN:", formats['bold_left'])
    ws.write(r_sig+1, 3, "PROJECT-IN-CHARGE", formats['bold_left'])
    ws.write(r_sig+2, 3, "MR. YUNUS KHAN", formats['bold_left'])
    ws.write(r_sig+3, 3, "DATE:", formats['bold_left'])
    
    ws.write(r_sig, 12, "SIGN:", formats['bold_left'])
    ws.write(r_sig+1, 12, "DEPLOYMENT HEAD", formats['bold_left'])
    ws.write(r_sig+2, 12, "MR. MANISH NAHAR", formats['bold_left'])
    ws.write(r_sig+3, 12, "DATE:", formats['bold_left'])

def write_matrix_sheet(wb, sheet_name, df_sites, code_to_col_idx, dc_number, formats, include_amounts=True):
    ws = wb.add_worksheet(sheet_name)
    num_sites = len(df_sites)
    
    # Calculate dates
    date_col = 'Completion Date ' if 'Completion Date ' in df_sites.columns else 'Completion Date'
    min_date_str = "N/A"
    max_date_str = "N/A"
    if date_col in df_sites.columns:
        dates = pd.to_datetime(df_sites[date_col], errors='coerce')
        min_date_str = dates.min().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"
        max_date_str = dates.max().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"

    tot_col = 3 + num_sites
    last_col = tot_col + 2 if include_amounts else tot_col
    
    ws.set_column(0, 0, 15)
    ws.set_column(1, 1, 40)
    ws.set_column(2, 2, 8)
    for col in range(3, tot_col):
        ws.set_column(col, col, 5)
    ws.set_column(tot_col, tot_col, 15)
    if include_amounts:
        ws.set_column(tot_col + 1, tot_col + 1, 15)
        ws.set_column(tot_col + 2, tot_col + 2, 15)

    f_title = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9D9D9', 'border': 1})
    f_center = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    f_head = wb.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#DCE6F1', 'border': 1})
    f_head_vert = wb.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'rotation': 90, 'bg_color': '#DCE6F1', 'border': 1})
    
    # Title Block
    ws.merge_range(0, 0, 0, last_col, sheet_name, f_title)
    ws.merge_range(1, 0, 1, last_col, 'Work Order No P14/630330726', f_center)
    ws.merge_range(2, 0, 2, last_col, 'Contractor Name: DIGITCOM INDIA TECHNOLOGIES       Work Order Dated: 03-10-2025', f_center)
    ws.merge_range(3, 0, 3, last_col, 'WO for Airspan A6 and C6 Radios for Airfiber', f_center)
    ws.merge_range(4, 0, 4, last_col, f'Service Done From Date: {min_date_str}', f_center)
    ws.merge_range(5, 0, 5, last_col, f'Service Done To Date: {max_date_str}', f_center)
    
    # Site Headers
    ws.write(7, 1, 'Count -', formats['bold_right'])
    ws.write(7, 2, '', formats['bold_right'])
    ws.write(8, 0, 'Code', f_head)
    ws.write(8, 1, 'Site ID --', f_head)
    ws.write(8, 2, '', f_head)
    
    ws.set_row(8, 150)  # Increase row height for rotated text
    
    for i, (_, row) in enumerate(df_sites.iterrows()):
        col = 3 + i
        ws.write(7, col, i + 1, formats['number_bold'])
        pmp_id = str(row.get('PMP ID', '')).strip()
        ws.write(8, col, pmp_id, f_head_vert)
        
    ws.write(8, tot_col, 'Total Quantity', f_head)
    if include_amounts:
        ws.write(8, tot_col + 1, 'RATE AS PER SOW', f_head)
        ws.write(8, tot_col + 2, 'AMOUNT', f_head)
        
    # Site Type & Sectors
    ws.write(10, 1, 'Site Type', formats['bold_right'])
    ws.write(10, 2, '', formats['bold_right'])
    for i, (_, row) in enumerate(df_sites.iterrows()):
        tt = str(row.get('Tower type', '')).strip()
        ws.write(10, 3 + i, tt, formats['number'])
        
    ws.write(12, 1, 'Sectors', formats['bold_right'])
    ws.write(12, 2, '', formats['bold_right'])
    total_sectors = 0
    for i, (_, row) in enumerate(df_sites.iterrows()):
        sec = safe_float(row.get('NO OF SECTOR'))
        total_sectors += sec
        ws.write(12, 3 + i, sec, formats['number'])
    ws.write(12, tot_col, total_sectors, formats['number_bold'])
    
    # Data Table Headers
    ws.write(13, 0, 'Item code', f_head)
    ws.write(13, 1, 'Description of Item', f_head)
    ws.write(13, 2, 'UOM', f_head)
    
    r_idx = 14
    for item in TEMPLATE_ITEMS:
        ws.write(r_idx, 0, item['sap'], formats['cell'])
        ws.write(r_idx, 1, item['desc'], formats['cell_left'])
        ws.write(r_idx, 2, item['uom'], formats['cell'])
        
        row_sum = 0
        for i, (_, site_row) in enumerate(df_sites.iterrows()):
            col = 3 + i
            sap_code = item['sap']
            val = site_row.iloc[code_to_col_idx[sap_code]] if sap_code in code_to_col_idx else 0.0
            
            if sap_code == "3397248":
                val = safe_float(site_row.get('EXTRA VISIT IN WO', 0.0))
            elif sap_code == "3268025":
                val = safe_float(site_row.get('Polemount in wo', 0.0))
            else:
                val = safe_float(val)
                
            ws.write(r_idx, col, val, formats['number'])
            row_sum += val
            
        ws.write_formula(r_idx, tot_col, f"=SUM({xlsxwriter.utility.xl_col_to_name(3)}{r_idx+1}:{xlsxwriter.utility.xl_col_to_name(tot_col-1)}{r_idx+1})", formats['number_bold'])
        
        if include_amounts:
            rate = safe_float(item['rate'])
            ws.write(r_idx, tot_col + 1, rate, formats['number'])
            ws.write_formula(r_idx, tot_col + 2, f"={xlsxwriter.utility.xl_col_to_name(tot_col)}{r_idx+1}*{xlsxwriter.utility.xl_col_to_name(tot_col+1)}{r_idx+1}", formats['number_bold'], row_sum * rate)
            
        r_idx += 1
        
    if include_amounts:
        ws.merge_range(r_idx, 0, r_idx, tot_col + 1, "TOTAL", formats['bold_right'])
        ws.write_formula(r_idx, tot_col + 2, f"=SUM({xlsxwriter.utility.xl_col_to_name(tot_col+2)}15:{xlsxwriter.utility.xl_col_to_name(tot_col+2)}{r_idx})", formats['number_bold'])

    r_sig = r_idx + 2
    left_col = 1
    right_col = tot_col + 2 if include_amounts else tot_col
    
    ws.write(r_sig, left_col, "SIGN:", formats['bold_left'])
    ws.write(r_sig+1, left_col, "PROJECT-IN-CHARGE", formats['bold_left'])
    ws.write(r_sig+2, left_col, "MR. YUNUS KHAN", formats['bold_left'])
    ws.write(r_sig+3, left_col, "DATE:", formats['bold_left'])
    
    if sheet_name == 'BOQ':
        mid_col = (left_col + right_col) // 2
        ws.write(r_sig, mid_col, "SIGN:", formats['bold_left'])
        ws.write(r_sig+1, mid_col, "DEPLOYMENT HEAD", formats['bold_left'])
        ws.write(r_sig+2, mid_col, "MR. MANISH NAHAR", formats['bold_left'])
        ws.write(r_sig+3, mid_col, "DATE:", formats['bold_left'])
        
        ws.write(r_sig, right_col, "SIGN:", formats['bold_left'])
        ws.write(r_sig+1, right_col, "RJIO CTO", formats['bold_left'])
        ws.write(r_sig+2, right_col, "MR.RAJEEV KUMAR GUPTA", formats['bold_left'])
        ws.write(r_sig+3, right_col, "DATE:", formats['bold_left'])
    else:
        ws.write(r_sig, right_col, "SIGN:", formats['bold_left'])
        ws.write(r_sig+1, right_col, "DEPLOYMENT HEAD", formats['bold_left'])
        ws.write(r_sig+2, right_col, "MR. MANISH NAHAR", formats['bold_left'])
        ws.write(r_sig+3, right_col, "DATE:", formats['bold_left'])

def write_declaration(wb, df_sites, dc_number, formats):
    ws = wb.add_worksheet('Declaration')
    # Increased widths to prevent Apple Numbers from visually stretching the merged cells
    ws.set_column('A:A', 35)
    ws.set_column('B:C', 30)
    ws.set_column('D:D', 40)
    
    f_title = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 2})
    f_bold = wb.add_format({'bold': True, 'font_size': 10, 'align': 'left', 'valign': 'vcenter', 'border': 1})
    f_bold_center = wb.add_format({'bold': True, 'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    f_center = wb.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    f_text = wb.add_format({'font_size': 10, 'align': 'left', 'valign': 'top', 'text_wrap': True, 'border': 1})
    f_empty = wb.add_format({'border': 1})
    
    # Row 1
    ws.merge_range('A1:D1', 'DECLARATION STATEMENT', f_title)
    
    # Row 2
    ws.write('A2', 'Name of Contractor', f_bold)
    ws.merge_range('B2:C2', 'DIGITCOM INDIA TECHNOLOGIES', f_bold_center)
    ws.merge_range('D2:D6', '', f_empty) # Logo placeholder
    
    # Row 3 & 4
    ws.merge_range('A3:A4', 'Authorised Signatory', f_bold)
    ws.merge_range('B3:C4', '', f_empty)
    
    # Row 5
    ws.write('A5', 'Vendor Code', f_bold)
    ws.merge_range('B5:C5', '3267708', f_bold_center)
    
    # Row 6
    ws.write('A6', 'Work Order No:', f_bold)
    ws.merge_range('B6:C6', 'P14/630330726', f_bold_center)
    
    # Row 7
    ws.write('A7', 'SAP ID/WBS :', f_bold)
    ws.merge_range('B7:D7', 'As per Annexture', f_bold_center)
    
    # Row 8
    ws.write('A8', 'Warehouse Location', f_bold)
    ws.merge_range('B8:D8', 'JODHPUR', f_bold_center)
    
    # Row 9
    ws.merge_range('A9:D9', 'Declaration', f_bold_center)
    
    # Get max date
    date_col = 'Completion Date ' if 'Completion Date ' in df_sites.columns else 'Completion Date'
    max_date_str = "30.03.2026"
    if date_col in df_sites.columns:
        dates = pd.to_datetime(df_sites[date_col], errors='coerce')
        if not dates.isna().all():
            max_date_str = dates.max().strftime('%d.%m.%Y')
            
    cert_text = f"We hereby certify that this Material Reconcilation Statement as on {max_date_str} attached herein is certified and justified by Bills submited by Contractor for given Work Done on given site as per Work Order Issued."
    
    # Row 10-12
    ws.merge_range('A10:D12', cert_text, f_text)
    ws.set_row(9, 30)
    ws.set_row(10, 30)
    ws.set_row(11, 30)
    
    # Row 13
    ws.merge_range('A13:D13', f'{len(df_sites)} SITES(A6)', f_bold_center)
    
    # Row 14 (blank narrow with vertical split)
    ws.write('A14', '', f_empty)
    ws.merge_range('B14:D14', '', f_empty)
    ws.set_row(13, 8)
    
    # Row 15 (Headers)
    ws.write('A15', 'Warehouse Incharge', f_bold_center)
    ws.write('B15', 'Circle Project Head', f_bold_center)
    ws.write('C15', 'Contractor', f_bold_center)
    ws.write('D15', 'Remark', f_bold_center)
    
    # Row 16-19 (Signatures row 1)
    ws.merge_range('A16:A19', '', f_empty)
    ws.merge_range('B16:B19', '', f_empty)
    ws.merge_range('C16:C19', '', f_empty)
    ws.merge_range('D16:D23', '', f_empty) # Remarks spans down
    
    # Row 20
    ws.write('A20', 'Checked By Material Coordinator', f_bold)
    ws.merge_range('B20:C20', 'Through FC&A', f_bold_center)
    
    # Row 21-23 (Signatures row 2)
    ws.merge_range('A21:A23', '', f_empty)
    ws.merge_range('B21:C23', '', f_empty)
    
    # Give signature rows height
    ws.set_row(15, 20)
    ws.set_row(16, 20)
    ws.set_row(17, 20)
    ws.set_row(18, 20)
    ws.set_row(20, 20)
    ws.set_row(21, 20)
    ws.set_row(22, 20)

def write_annexure_and_reco(wb, df_sites, dc_number, formats, mindump_path):
    ws_ann = wb.add_worksheet('Annexture')
    ws_rec = wb.add_worksheet('Reco')
    
    pmp_ids = df_sites['PMP ID'].astype(str).str.strip().tolist()
    num_sites = len(pmp_ids)
    tot_col = num_sites + 1
    desc_col = tot_col + 1
    
    ws_ann.set_column(0, 0, 15)
    for col in range(1, tot_col):
        ws_ann.set_column(col, col, 5)
    ws_ann.set_column(tot_col, tot_col, 10)
    ws_ann.set_column(desc_col, desc_col, 40)
    
    f_title = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    f_head = wb.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'bg_color': '#DCE6F1'})
    f_head_vert = wb.add_format({'bold': True, 'font_size': 8, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'rotation': 90, 'border': 1, 'bg_color': '#DCE6F1'})
    f_cell = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    f_cell_bold = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    f_sum_row = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1'})
    f_sum_row_bold = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1'})
    f_empty = wb.add_format({'border': 1})
    
    ws_ann.merge_range(0, 0, 0, tot_col, 'Annexture', f_title)
    ws_ann.write(0, desc_col, '', f_empty)
    
    ws_ann.set_row(1, 120)
    ws_ann.write(1, 0, 'Row Labels', f_head)
    for i, pid in enumerate(pmp_ids):
        ws_ann.write(1, i + 1, pid, f_head_vert)
    ws_ann.write(1, tot_col, 'Grand Total', f_head)
    ws_ann.write(1, desc_col, '', f_empty)
    
    # Process mindump
    try:
        df_mindump = pd.read_excel(mindump_path)
        
        def match_site(row):
            wbs = str(row.get('WBS ID', ''))
            site_id = str(row.get('Site ID', ''))
            for pid in pmp_ids:
                if pid in wbs or pid in site_id:
                    return pid
            return None
            
        df_mindump['Matched_PMP'] = df_mindump.apply(match_site, axis=1)
        df_filtered = df_mindump[df_mindump['Matched_PMP'].notna()]
        
        if not df_filtered.empty:
            pt = pd.pivot_table(df_filtered, values='No. Of Qty', index=['SAP Code', 'Material Description'], columns='Matched_PMP', aggfunc='sum', fill_value=0)
            for pid in pmp_ids:
                if pid not in pt.columns:
                    pt[pid] = 0
            pt = pt[pmp_ids]
            pt = pt.sort_index()
        else:
            pt = pd.DataFrame(columns=pmp_ids)
            
    except Exception as e:
        print(f"Warning: Could not process MINDUMP for Annexture: {e}")
        pt = pd.DataFrame(columns=pmp_ids)
        
    r_idx = 2
    col_sums = {pid: 0 for pid in pmp_ids}
    grand_total_sum = 0
    reco_items = []
    
    for (sap_code, desc), row in pt.iterrows():
        ws_ann.write(r_idx, 0, str(sap_code), f_cell_bold)
        row_total = 0
        for i, pid in enumerate(pmp_ids):
            val = safe_float(row[pid])
            ws_ann.write(r_idx, i + 1, val, f_cell)
            col_sums[pid] += val
            row_total += val
            
        ws_ann.write(r_idx, tot_col, row_total, f_cell)
        grand_total_sum += row_total
        ws_ann.write(r_idx, desc_col, str(desc), f_cell_bold)
        reco_items.append({"sap": str(sap_code), "desc": str(desc), "annexure_row": r_idx + 1})
        r_idx += 1
        
    # Bottom Grand Total Row
    ws_ann.write(r_idx, 0, 'Grand Total', f_sum_row_bold)
    for i, pid in enumerate(pmp_ids):
        ws_ann.write(r_idx, i + 1, col_sums[pid], f_sum_row)
    ws_ann.write(r_idx, tot_col, grand_total_sum, f_sum_row)
    ws_ann.write(r_idx, desc_col, '', f_empty)
    
    # --- RECO SHEET ---
    # Calculate dates
    date_col = 'Completion Date ' if 'Completion Date ' in df_sites.columns else 'Completion Date'
    min_date_str = "N/A"
    max_date_str = "N/A"
    if date_col in df_sites.columns:
        dates = pd.to_datetime(df_sites[date_col], errors='coerce')
        min_date_str = dates.min().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"
        max_date_str = dates.max().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"

    num_items = len(reco_items)
    last_col = 1 + num_items
    
    # Setup columns
    ws_rec.set_column(0, 0, 5)   # Col A: Index
    ws_rec.set_column(1, 1, 45)  # Col B: Description Label
    for i in range(num_items):
        ws_rec.set_column(2 + i, 2 + i, 12)
        
    f_reco_header_label = wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1})
    f_reco_header_val = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    f_r4g_box = wb.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'border': 2})
    f_reco_title = wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DCE6F1'})
    
    # Standard Table Formats
    f_reco_item_head = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_size': 9, 'bg_color': '#D9D9D9'})
    f_reco_cell = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0'})
    f_reco_cell_bold = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0'})
    f_reco_label = wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    f_reco_label_bold = wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    f_reco_index = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    
    # Specific Color Formats
    f_blue_head = wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': '#BDD7EE', 'text_wrap': True})
    f_blue_index = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#BDD7EE'})
    
    f_yellow_label = wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00', 'text_wrap': True})
    f_yellow_cell = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00', 'num_format': '0'})
    f_yellow_index = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00'})
    
    f_empty = wb.add_format({'border': 1})
    
    # Metadata Header Block (Rows 0-6)
    meta_labels = [
        ('Name of Contractor', 'DIGITCOM INDIA TECHNOLOGIES'),
        ('WO No', 'P14/630330726'),
        ('Site ID', 'As per attached Detail'),
        ('WBS Code', 'As per attached Detail'),
        ('Site Name', 'As per attached Detail'),
        ('Warehouse Location', 'JODHPUR'),
        ('Time Period', f'{min_date_str} TO {max_date_str}')
    ]
    
    for r, (label, val) in enumerate(meta_labels):
        ws_rec.write(r, 0, label, f_reco_header_label)
        ws_rec.merge_range(r, 1, r, 5, val, f_reco_header_label)
        
    # R4G Project Box (Merged Rows 0-6, starting from Col 6 to last_col)
    if last_col >= 6:
        ws_rec.merge_range(0, 6, 6, last_col, 'R4G Project', f_r4g_box)
    else:
        ws_rec.write(0, last_col + 1, 'R4G Project', f_r4g_box)
        
    # Main Title Row (Row 8)
    ws_rec.merge_range(8, 0, 8, last_col, 'Material Reconciliation Statement', f_reco_title)
    
    # Items Headers (Rows 10-11)
    ws_rec.set_row(10, 60)
    ws_rec.write(10, 1, 'Material Description', f_reco_item_head)
    ws_rec.write(11, 1, 'Material Code', f_reco_item_head)
    ws_rec.write(10, 0, '', f_reco_item_head)
    ws_rec.write(11, 0, '', f_reco_item_head)
    
    for i, item in enumerate(reco_items):
        col = 2 + i
        # Keep numeric item headers standard (White) instead of Gray/Blue
        ws_rec.write(10, col, item['desc'], f_reco_cell_bold)
        ws_rec.write(11, col, item['sap'], f_reco_cell_bold)
        
    # Section A
    ws_rec.write(12, 0, 'A.', f_blue_index)
    ws_rec.write(12, 1, 'Receipt details as confirmed by Contractor:', f_blue_head)
    for i in range(num_items): ws_rec.write(12, 2 + i, '', f_empty)
    
    ws_rec.write(14, 0, '1', f_reco_index)
    ws_rec.write(14, 1, 'Received directly from Warehouse', f_yellow_label)
    for i, item in enumerate(reco_items):
        col = 2 + i
        annex_col_str = xlsxwriter.utility.xl_col_to_name(tot_col)
        ws_rec.write_formula(14, col, f"=Annexture!{annex_col_str}{item['annexure_row']}", f_reco_cell)
        
    ws_rec.write(16, 0, '2', f_reco_index)
    ws_rec.write(16, 1, 'Material received from other contractors', f_reco_label)
    for i in range(num_items): ws_rec.write(16, 2 + i, 0, f_reco_cell)
        
    ws_rec.write(18, 0, '', f_reco_index)
    ws_rec.write(18, 1, 'Total (1+2)', f_yellow_label)
    for i in range(num_items):
        col_str = xlsxwriter.utility.xl_col_to_name(2 + i)
        ws_rec.write_formula(18, 2 + i, f"={col_str}15+{col_str}17", f_reco_cell_bold)
        
    # Section B
    ws_rec.write(20, 0, 'B.', f_blue_index)
    ws_rec.write(20, 1, 'Material Transferred to other contractors / returned to Warehouse', f_blue_head)
    for i in range(num_items): ws_rec.write(20, 2 + i, '', f_empty)
    
    ws_rec.write(21, 0, '1', f_reco_index)
    ws_rec.write(21, 1, 'Material Transferred to other contractors', f_reco_label)
    for i in range(num_items): ws_rec.write(21, 2 + i, 0, f_reco_cell)
        
    ws_rec.write(22, 0, '2', f_reco_index)
    ws_rec.write(22, 1, 'Material Returned to Warehouse (in line with MRN Guidelines)', f_reco_label)
    for i in range(num_items): ws_rec.write(22, 2 + i, 0, f_reco_cell)
        
    ws_rec.write(23, 0, '', f_reco_index)
    ws_rec.write(23, 1, 'Total', f_reco_label_bold)
    for i in range(num_items):
        col_str = xlsxwriter.utility.xl_col_to_name(2 + i)
        ws_rec.write_formula(23, 2 + i, f"={col_str}22+{col_str}23", f_reco_cell_bold)
        
    # Section C
    ws_rec.write(25, 0, 'C', f_yellow_index)
    ws_rec.write(25, 1, 'Balance ( A  -  B )', f_yellow_label)
    for i in range(num_items):
        col_str = xlsxwriter.utility.xl_col_to_name(2 + i)
        ws_rec.write_formula(25, 2 + i, f"={col_str}19-{col_str}24", f_reco_cell_bold)
        
    # Section D
    ws_rec.write(27, 0, 'D', f_reco_index)
    ws_rec.write(27, 1, 'CONSUMPTION', f_reco_label_bold)
    for i in range(num_items): ws_rec.write(27, 2 + i, '', f_empty)
    
    ws_rec.write(28, 0, '1', f_reco_index)
    ws_rec.write(28, 1, 'Material Consumed', f_yellow_label)
    for i in range(num_items):
        col_str = xlsxwriter.utility.xl_col_to_name(2 + i)
        annex_col_str = xlsxwriter.utility.xl_col_to_name(tot_col)
        ws_rec.write_formula(28, 2 + i, f"=Annexture!{annex_col_str}{reco_items[i]['annexure_row']}", f_reco_cell)
        
    ws_rec.write(29, 0, '2', f_reco_index)
    ws_rec.write(29, 1, 'Wastage (max as per WO norms)', f_reco_label)
    for i in range(num_items): ws_rec.write(29, 2 + i, 0, f_reco_cell)
        
    ws_rec.write(31, 0, '', f_reco_index)
    ws_rec.write(31, 1, 'Total -Actual consumption (1+2)', f_yellow_label)
    for i in range(num_items):
        col_str = xlsxwriter.utility.xl_col_to_name(2 + i)
        ws_rec.write_formula(31, 2 + i, f"={col_str}29+{col_str}30", f_reco_cell_bold)
        
    # Section E
    ws_rec.set_row(33, 40)
    ws_rec.write(33, 0, 'E.', f_yellow_index)
    ws_rec.write(33, 1, 'Excess consumption for which cost to be recovered from the Contractor (C-D)', f_yellow_label)
    for i in range(num_items):
        col_str = xlsxwriter.utility.xl_col_to_name(2 + i)
        ws_rec.write_formula(33, 2 + i, f"={col_str}26-{col_str}32", f_reco_cell_bold)

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
                'title': wb.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#DCE6F1', 'border': 1}),
                'cert_text': wb.add_format({'bold': True, 'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1}),
                'header': wb.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#DCE6F1', 'border': 1, 'text_wrap': True}),
                'header_blue': wb.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#DCE6F1', 'border': 1, 'text_wrap': True}),
                'header_yellow': wb.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00', 'border': 1, 'text_wrap': True}),
                'header_vertical': wb.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'rotation': 90, 'bg_color': '#DCE6F1', 'border': 1}),
                'cell': wb.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1}),
                'cell_left': wb.add_format({'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'border': 1}),
                'number': wb.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0'}),
                'number_bold': wb.add_format({'font_size': 9, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0', 'bg_color': '#F2F2F2'}),
                'bold_right': wb.add_format({'font_size': 9, 'bold': True, 'align': 'right', 'valign': 'vcenter'}),
                'bold_left': wb.add_format({'font_size': 9, 'bold': True, 'align': 'left', 'valign': 'vcenter'}),
                'date': wb.add_format({'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'dd-mmm-yy'})
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
