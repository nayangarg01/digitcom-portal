import pandas as pd
import xlsxwriter
import argparse
import os
import openpyxl
import re

def safe_float(val):
    try:
        if val is None or val == '': return 0.0
        return float(val)
    except:
        return 0.0

def generate_performa_invoice(dc_files, mindump_path, iv_number, output_path):
    # 1. Load MINDUMP for WBS lookup (A6 Dump)
    wbs_mapping = {}
    try:
        df_dump = pd.read_excel(mindump_path, sheet_name='A6 DUMP')
        # We need a mapping from Site ID or PMP ID to WBS ID
        # Looking at previous inspection: Site ID or WBS ID columns exist
        for _, row in df_dump.iterrows():
            site_id = str(row.get('Site ID', '')).strip()
            wbs_id = str(row.get('WBS ID', '')).strip()
            if site_id and wbs_id:
                wbs_mapping[site_id] = wbs_id
    except Exception as e:
        print(f"Warning: Could not load MINDUMP for WBS mapping: {e}")

    # 2. Process each DC file
    all_rows = []
    
    for dc_file in dc_files:
        if not os.path.exists(dc_file):
            print(f"Warning: File not found: {dc_file}")
            continue
            
        try:
            # Using openpyxl to read precisely
            wb = openpyxl.load_workbook(dc_file, data_only=True)
            
            # Get WO Number from JMS sheet
            wo_number = "N/A"
            if 'JMS' in wb.sheetnames:
                ws_jms = wb['JMS']
                # Try to find WO number in header area (e.g., H4)
                # Looking at inspection: ('JMS', ..., 'Work Order No :P14/630330726\n', ...)
                # It was in row 4, column H (index 7)
                header_val = str(ws_jms.cell(row=4, column=8).value or "")
                if "Work Order No" in header_val:
                    wo_number = header_val.split(':')[-1].strip()
            
            # Extract Sites and their data from JMS
            # Sites are in row 12, starting from column D (index 4)
            sites = []
            for col in range(4, ws_jms.max_column + 1):
                site_id = str(ws_jms.cell(row=12, column=col).value or "").strip()
                if site_id and site_id.startswith('I-RJ'):
                    sites.append({'id': site_id, 'col': col})
                elif site_id == 'Total Quantity':
                    break # End of sites area
            
            # Get Line Items
            # Descriptions start from row 16, column B (index 2)
            # Rates are in column 'RATE AS PER SOW'
            # Find column indices for Rate and Amount
            header_row_12 = [str(ws_jms.cell(row=12, column=c).value).strip() for c in range(1, ws_jms.max_column + 1)]
            try:
                rate_col = header_row_12.index('RATE AS PER SOW') + 1
                amount_col = header_row_12.index('AMOUNT') + 1
            except:
                # Fallback if names changed
                rate_col = ws_jms.max_column - 2
                amount_col = ws_jms.max_column - 1

            for site in sites:
                site_id = site['id']
                site_col = site['col']
                wbs_id = wbs_mapping.get(site_id, "N/A")
                
                # Iterate rows for materials
                for r in range(16, ws_jms.max_row + 1):
                    desc = str(ws_jms.cell(row=r, column=2).value or "").strip()
                    if not desc or desc.upper() == 'TOTAL': break
                    
                    qty = safe_float(ws_jms.cell(row=r, column=site_col).value)
                    if qty > 0:
                        rate = safe_float(ws_jms.cell(row=r, column=rate_col).value)
                        amount = safe_float(ws_jms.cell(row=r, column=amount_col).value) # This might be total amount, not site amount
                        # Actually, site amount is qty * rate
                        site_amount = qty * rate
                        
                        all_rows.append({
                            'vendor': 'DIGITCOM INDIA TECHNOLOGIES',
                            'scope': 'ISP',
                            'iv_no': f"PERFORMA INVOICE NO. {iv_number}",
                            'wo_no': wo_number,
                            'site': site_id,
                            'wbs': wbs_id,
                            'description': desc,
                            'nature': 'AIR FIBER INSTALLATION',
                            'qty': qty,
                            'rate': rate,
                            'amount': site_amount
                        })
                        
        except Exception as e:
            print(f"Error processing {dc_file}: {e}")

    # 3. Create Final Workbook
    with xlsxwriter.Workbook(output_path) as wb:
        # Formats
        f_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9', 'font_size': 10})
        f_cell = wb.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'font_size': 9})
        f_num = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9})
        f_amount = wb.add_format({'border': 1, 'align': 'right', 'valign': 'vcenter', 'font_size': 9, 'num_format': '#,##0.00'})
        f_title = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})

        # Sheet "1"
        ws1 = wb.add_worksheet('1')
        headers = ['VENDOR', 'SCOPE', 'IV NO', 'WO NO', 'SE NO', 'JMS', 'CI', 'SITE', 'WBS', 'DESC-SRIPTION', 'NATURE OF WORK', 'QTY', 'RATE', 'AMOUNT']
        for i, h in enumerate(headers):
            ws1.write(0, i, h, f_header)
        
        ws1.set_column(0, 0, 30) # Vendor
        ws1.set_column(1, 1, 10) # Scope
        ws1.set_column(2, 2, 25) # IV No
        ws1.set_column(3, 3, 20) # WO No
        ws1.set_column(7, 7, 20) # Site
        ws1.set_column(8, 8, 20) # WBS
        ws1.set_column(9, 9, 40) # Desc
        ws1.set_column(10, 10, 25) # Nature
        
        row_idx = 1
        for r in all_rows:
            ws1.write(row_idx, 0, r['vendor'], f_cell)
            ws1.write(row_idx, 1, r['scope'], f_cell)
            ws1.write(row_idx, 2, r['iv_no'], f_cell)
            ws1.write(row_idx, 3, r['wo_no'], f_cell)
            ws1.write(row_idx, 4, '', f_cell) # SE NO
            ws1.write(row_idx, 5, '', f_cell) # JMS
            ws1.write(row_idx, 6, '', f_cell) # CI
            ws1.write(row_idx, 7, r['site'], f_cell)
            ws1.write(row_idx, 8, r['wbs'], f_cell)
            ws1.write(row_idx, 9, r['description'], f_cell)
            ws1.write(row_idx, 10, r['nature'], f_cell)
            ws1.write(row_idx, 11, r['qty'], f_num)
            ws1.write(row_idx, 12, r['rate'], f_amount)
            ws1.write(row_idx, 13, r['amount'], f_amount)
            row_idx += 1

        # Summary Sheet
        ws_sum = wb.add_worksheet('Summary sheet1')
        sum_headers = ['Sl No', 'WO NO', 'SITE ID', 'SITE COUNT', 'TOTAL AMOUNT']
        for i, h in enumerate(sum_headers):
            ws_sum.write(0, i, h, f_header)
        
        # Group by WO and Site to get summary
        df_all = pd.DataFrame(all_rows)
        if not df_all.empty:
            summary = df_all.groupby(['wo_no', 'site']).agg({'amount': 'sum'}).reset_index()
            for i, row in summary.iterrows():
                ws_sum.write(i + 1, 0, i + 1, f_num)
                ws_sum.write(i + 1, 1, row['wo_no'], f_cell)
                ws_sum.write(i + 1, 2, row['site'], f_cell)
                ws_sum.write(i + 1, 3, 1, f_num)
                ws_sum.write(i + 1, 4, row['amount'], f_amount)
            
            # Total Row
            last_row = len(summary) + 1
            ws_sum.write(last_row, 2, 'GRAND TOTAL', f_header)
            ws_sum.write(last_row, 3, len(summary), f_num)
            ws_sum.write(last_row, 4, df_all['amount'].sum(), f_amount)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--files", nargs='+', required=True)
    parser.add_argument("--mindump", required=True)
    parser.add_argument("--iv_number", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    
    generate_performa_invoice(args.files, args.mindump, args.iv_number, args.output)
