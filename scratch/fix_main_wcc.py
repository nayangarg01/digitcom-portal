def write_main_wcc(wb, df_sites, dc_number, formats):
    ws = wb.add_worksheet('Main WCC')
    ws.set_column('A:A', 2)
    ws.set_column('B:B', 20)
    ws.set_column('C:C', 10)
    ws.set_column('D:D', 10)
    ws.set_column('E:E', 10)
    ws.set_column('F:F', 10)
    ws.set_column('G:G', 10)
    ws.set_column('H:H', 30)

    # Calculate dates and sites
    num_sites = len(df_sites)
    date_col = None
    for c in ['Completion Date ', 'Completion Date', 'RFS DATE']:
        if c in df_sites.columns:
            date_col = c
            break
            
    date_range = "N/A"
    if date_col:
        dates = pd.to_datetime(df_sites[date_col], errors='coerce')
        min_date = dates.min().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"
        max_date = dates.max().strftime('%d-%b-%y').upper() if not dates.isna().all() else "N/A"
        date_range = f"{min_date} TO {max_date}"

    # Extract WO Number
    wo_number = "N/A"
    if 'W.O.Number' in df_sites.columns:
        wo_number = str(df_sites['W.O.Number'].iloc[0])
    elif 'W.O. Number' in df_sites.columns:
        wo_number = str(df_sites['W.O. Number'].iloc[0])

    # Extract Rep Names
    vendor_rep = "ANKUSH SRIVASTAVA"
    rjil_rep = "MR. Manish Nahar"
    if 'Engineer Name' in df_sites.columns:
        vendor_rep = str(df_sites['Engineer Name'].iloc[0]).upper()
    if 'JC Name' in df_sites.columns:
        rjil_rep = str(df_sites['JC Name'].iloc[0])

    # Formats
    f_title = wb.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'underline': True})
    f_label = wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter'})
    f_val = wb.add_format({'align': 'center', 'valign': 'vcenter'})
    f_dash = wb.add_format({'align': 'center', 'valign': 'vcenter', 'bottom': 1})
    f_box_empty = wb.add_format({'border': 1, 'bg_color': '#FFFFFF'})
    f_box_yellow = wb.add_format({'border': 1, 'bg_color': '#FFFF00'})
    f_border_box = wb.add_format({'border': 2})
    f_center_bold = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    f_norm = wb.add_format({'align': 'left', 'valign': 'vcenter'})
    f_bold = wb.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter'})
    
    # Logo
    logo_path = "/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/Screenshot 2026-05-04 at 6.45.10 PM.png"
    if os.path.exists(logo_path):
        ws.insert_image('B2', logo_path, {'x_scale': 0.15, 'y_scale': 0.15, 'x_offset': 5, 'y_offset': 5})

    # Main Border (Approx B2:H38)
    ws.merge_range('C2:H3', 'Work Completion Certificate', f_title)
    
    # State / MP
    ws.write('B8', 'State', f_label)
    ws.merge_range('C8:D8', 'RAJASTHAN', f_center_bold)
    ws.write('C9', '_ _ _ _ _ _ _ _', f_val)
    ws.write('D9', '_ _ _ _ _ _ _ _', f_val)
    
    ws.merge_range('E8:G8', 'Maintenance Point', f_label)
    ws.write('H8', 'Jaipur', f_center_bold)
    ws.write('H9', '_ _ _ _ _ _ _ _ _ _ _ _', f_val)

    # Project Type
    ws.write('B10', 'Project Type', f_label)
    ws.write('C10', '93 K', f_val)
    ws.write('D10', 'Infill', f_val)
    ws.write('E10', 'Growth', f_val)
    ws.merge_range('F10:G10', 'Other (Specify) _________________', f_val)
    ws.write('H10', 'Air Fiber Installation', f_val)
    
    ws.write('C11', '', f_box_empty)
    ws.write('D11', '', f_box_empty)
    ws.write('E11', '', f_box_empty)
    ws.write('F11', '', f_box_empty)

    # Site Type
    ws.write('B13', 'Site Type', f_label)
    ws.write('C13', 'Own Built', f_val)
    ws.write('D13', 'IP Colo', f_val)
    ws.write('E13', 'RP1', f_val)
    ws.write('F13', 'BSNL', f_val)
    ws.write('G13', 'MAG1 NLD AG1', f_val)
    ws.write('H13', 'ZXZ', f_val)
    
    ws.write('C14', '', f_box_empty)
    ws.write('D14', '', f_box_empty)
    ws.write('E14', '', f_box_empty)
    ws.write('F14', '', f_box_empty)
    ws.write('G14', '', f_box_empty)
    ws.write('H14', '', f_box_empty)

    # Tower type
    ws.write('B16', 'Tower type', f_label)
    ws.write('C16', 'GBT', f_val)
    ws.write('D16', 'RTT', f_val)
    ws.write('E16', 'RTP', f_val)
    ws.write('F16', 'GBM', f_val)
    ws.write('G16', 'NBT', f_val)
    ws.write('H16', 'Other (Specify) _________________', f_val)
    
    ws.write('C17', '', f_box_yellow)
    ws.write('D17', '', f_box_yellow)
    ws.write('E17', '', f_box_yellow)
    ws.write('F17', '', f_box_yellow)
    ws.write('G17', '', f_box_empty)
    ws.write('H17', '', f_box_empty)

    # Certification Text
    ws.merge_range('B19:H19', 'This is to certify that work has been completed as per specification given in workorder on the sites mentioned', f_center_bold)
    ws.merge_range('B21:H21', 'The required ITP / Checklists are available and verified in system', f_center_bold)

    # Details Section
    ws.write('B23', 'Site Name', f_label)
    ws.merge_range('C23:D23', 'As per Annexture', f_center_bold)
    ws.write('C24', '_ _ _ _ _ _ _ _', f_val)
    ws.write('D24', '_ _ _ _ _ _ _ _', f_val)

    ws.write('E23', 'SAP ID', f_label)
    ws.merge_range('F23:H23', 'As per Annexture', f_center_bold)
    ws.write('F24', '_ _ _ _ _ _ _ _', f_val)
    ws.write('G24', '_ _ _ _ _ _ _ _', f_val)
    ws.write('H24', '_ _ _ _ _ _ _ _ _ _ _ _', f_val)

    ws.write('B25', 'W.O.Number', f_label)
    ws.merge_range('C25:D25', wo_number, f_center_bold)
    ws.write('C26', '_ _ _ _ _ _ _ _', f_val)
    ws.write('D26', '_ _ _ _ _ _ _ _', f_val)

    ws.write('E25', 'Vendor Name', f_label)
    ws.write('F25', 'M/S.', f_label)
    ws.merge_range('G25:H25', 'DIGITCOM INDIA TECHNOLOGIES', f_center_bold)
    ws.write('G26', '_ _ _ _ _ _ _ _', f_val)
    ws.write('H26', '_ _ _ _ _ _ _ _ _ _ _ _', f_val)

    ws.write('B27', 'No of Sites', f_label)
    ws.merge_range('C27:D27', f'{num_sites} SITES', f_center_bold)
    ws.write('C28', '_ _ _ _ _ _ _ _', f_val)
    ws.write('D28', '_ _ _ _ _ _ _ _', f_val)

    ws.write('E27', 'Completion Date', f_label)
    ws.merge_range('F27:H27', date_range, f_center_bold)
    ws.write('F28', '_ _ _ _ _ _ _ _', f_val)
    ws.write('G28', '_ _ _ _ _ _ _ _', f_val)
    ws.write('H28', '_ _ _ _ _ _ _ _ _ _ _ _', f_val)

    # Signature Section
    ws.merge_range('B30:D30', 'Vendor Representative', f_title)
    ws.merge_range('E30:H30', 'RJIL Representative', f_title)

    ws.write('B32', 'Name', f_label)
    ws.merge_range('C32:D32', vendor_rep, f_center_bold)
    ws.write('C33', '_ _ _ _ _ _ _ _', f_val)
    ws.write('D33', '_ _ _ _ _ _ _ _', f_val)

    ws.write('E32', 'Name', f_label)
    ws.merge_range('F32:H32', rjil_rep, f_center_bold)
    ws.write('F33', '_ _ _ _ _ _ _ _', f_val)
    ws.write('G33', '_ _ _ _ _ _ _ _', f_val)
    ws.write('H33', '_ _ _ _ _ _ _ _ _ _ _ _', f_val)

    ws.write('B34', 'Sign', f_label)
    ws.write('C35', '_ _ _ _ _ _ _ _', f_val)
    ws.write('D35', '_ _ _ _ _ _ _ _', f_val)

    ws.write('E34', 'Sign', f_label)
    ws.write('F35', '_ _ _ _ _ _ _ _', f_val)
    ws.write('G35', '_ _ _ _ _ _ _ _', f_val)
    ws.write('H35', '_ _ _ _ _ _ _ _ _ _ _ _', f_val)

    ws.write('B36', 'Date', f_label)
    ws.write('C37', '_ _ _ _ _ _ _ _', f_val)
    ws.write('D37', '_ _ _ _ _ _ _ _', f_val)

    ws.write('E36', 'Date', f_label)
    ws.write('F37', '_ _ _ _ _ _ _ _', f_val)
    ws.write('G37', '_ _ _ _ _ _ _ _', f_val)
    ws.write('H37', '_ _ _ _ _ _ _ _ _ _ _ _', f_val)

    # Note
    ws.write('B39', 'Note :', f_bold)
    ws.merge_range('C39:H39', 'In case of Multiple sites, please attach applicable site details with this certificate', f_norm)

    # Thick Outside Border simulation
    ws.conditional_format('B2:H39', {'type': 'no_errors', 'format': f_border_box})
