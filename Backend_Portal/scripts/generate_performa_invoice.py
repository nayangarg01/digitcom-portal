import argparse
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from collections import defaultdict

def safe_float(val):
    try:
        if val is None or val == '': return 0.0
        if isinstance(val, str):
            val = val.replace(',', '').strip()
        return float(val)
    except:
        return 0.0

def apply_border(cell):
    thin = Side(border_style="thin", color="000000")
    cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

def generate_performa_invoice(dc_files, mindump_path, iv_number, activity, output_path):
    # 1. Load MINDUMP for WBS lookup (A6 Dump)
    wbs_mapping = {}
    try:
        wb_dump = openpyxl.load_workbook(mindump_path, data_only=True)
        if 'A6 DUMP' in wb_dump.sheetnames:
            ws_dump = wb_dump['A6 DUMP']
            headers = [str(cell.value).strip() for cell in ws_dump[1]]
            site_id_idx = -1
            wbs_id_idx = -1
            for i, h in enumerate(headers):
                if h == 'Site ID': site_id_idx = i
                if h == 'WBS ID': wbs_id_idx = i
            if site_id_idx != -1 and wbs_id_idx != -1:
                for row in ws_dump.iter_rows(min_row=2, values_only=True):
                    site_id = str(row[site_id_idx]).strip() if row[site_id_idx] else ""
                    wbs_id = str(row[wbs_id_idx]).strip() if row[wbs_id_idx] else ""
                    if site_id and wbs_id and site_id.lower() != 'nan':
                        wbs_mapping[site_id] = wbs_id
        wb_dump.close()
    except Exception as e:
        print(f"Warning: Could not load MINDUMP for WBS mapping: {e}")

    # 2. Process each DC file
    all_rows = []
    for dc_file in dc_files:
        if not os.path.exists(dc_file):
            print(f"Warning: File not found: {dc_file}")
            continue
        try:
            wb = openpyxl.load_workbook(dc_file, data_only=True)
            nature_of_work = "AIR FIBER INSTALLATION"
            if activity == 'A6_B6':
                nature_of_work = "AIR FIBER INSTALLATION A6+B6"
            if 'JMS' not in wb.sheetnames:
                print(f"Warning: JMS sheet missing in {dc_file}")
                continue
            ws_jms = wb['JMS']
            wo_number = "N/A"
            vendor = "DIGITCOM INDIA TECHNOLOGIES"
            site_row = -1
            item_header_row = -1
            for r in range(1, 16):
                for c in range(1, 15):
                    cell_val = str(ws_jms.cell(row=r, column=c).value or "").strip()
                    if "Work Order No" in cell_val:
                        wo_number = cell_val.split(':')[-1].strip() if ':' in cell_val else cell_val.split()[-1].strip()
                    if "Contractor Name" in cell_val:
                        vendor = cell_val.split(':')[-1].strip() if ':' in cell_val else cell_val.replace("Contractor Name", "").strip()
            for r in range(1, 40):
                for c in range(1, 5):
                    val = str(ws_jms.cell(row=r, column=c).value or "").strip()
                    if "Site ID --" in val: site_row = r
                    if "Description of Item" in val: item_header_row = r
                if site_row != -1 and item_header_row != -1: break
            if site_row == -1 or item_header_row == -1:
                print(f"Warning: Missing headers in {dc_file}")
                continue
            sites = []
            for col in range(4, ws_jms.max_column + 1):
                val = str(ws_jms.cell(row=site_row, column=col).value or "").strip()
                if val and (val.startswith('I-RJ') or val.startswith('RJ')):
                    sites.append({'id': val, 'col': col})
                elif 'Total' in val or 'Quantity' in val: break
                elif not val and col > 4: break
            rate_col = -1
            for r_check in [site_row, item_header_row]:
                for col in range(1, ws_jms.max_column + 1):
                    val = str(ws_jms.cell(row=r_check, column=col).value or "").strip()
                    if 'RATE' in val.upper():
                        rate_col = col
                        break
                if rate_col != -1: break
            if rate_col == -1: rate_col = ws_jms.max_column - 1
            for site in sites:
                site_id = site['id']
                site_col = site['col']
                wbs_id = wbs_mapping.get(site_id, "N/A")
                empty_row_count = 0
                for r in range(item_header_row + 1, ws_jms.max_row + 1):
                    sap_code = str(ws_jms.cell(row=r, column=1).value or "").strip()
                    desc = str(ws_jms.cell(row=r, column=2).value or "").strip()
                    
                    if not desc or desc.upper() == 'TOTAL':
                        empty_row_count += 1
                        if empty_row_count > 5 or (desc and desc.upper() == 'TOTAL'):
                            break
                        continue
                    
                    empty_row_count = 0 # Reset if we find data
                    qty = safe_float(ws_jms.cell(row=r, column=site_col).value)
                    
                    if qty > 0:
                        rate = safe_float(ws_jms.cell(row=r, column=rate_col).value)
                        all_rows.append({
                            'vendor': vendor, 'scope': 'ISP', 'iv_no': f"PERFORMA INVOICE NO. {iv_number}",
                            'wo_no': wo_number, 'site': site_id, 'wbs': wbs_id, 'sap_code': sap_code,
                            'description': desc, 'nature': nature_of_work, 'qty': qty, 'rate': rate, 'amount': qty * rate
                        })
            wb.close()
        except Exception as e:
            print(f"Error processing {dc_file}: {e}")

    # 3. Create Final Workbook
    out_wb = openpyxl.Workbook()
    
    # Detail Sheet "1"
    ws1 = out_wb.active
    ws1.title = "1"
    headers = ['VENDOR', 'SCOPE', 'IV NO', 'WO NO', 'SE NO', 'JMS', 'CI', 'SITE', 'WBS', 'DESC-SRIPTION', 'NATURE OF WORK', 'QTY', 'RATE', 'AMOUNT']
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    header_font = Font(bold=True)
    for i, h in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=i)
        cell.value = h
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        apply_border(cell)
    
    for r_idx, r_data in enumerate(all_rows, 2):
        row = [r_data['vendor'], r_data['scope'], r_data['iv_no'], r_data['wo_no'], '', '', '', r_data['site'], r_data['wbs'], r_data['description'], r_data['nature'], r_data['qty'], r_data['rate'], r_data['amount']]
        for c_idx, val in enumerate(row, 1):
            cell = ws1.cell(row=r_idx, column=c_idx)
            cell.value = val
            apply_border(cell)
            if c_idx >= 12: cell.number_format = '#,##0'

    # Summary Sheet "Sheet1"
    ws_sum = out_wb.create_sheet("Sheet1")
    ws_sum.merge_cells('B2:F2')
    title_cell = ws_sum.cell(row=2, column=2)
    title_cell.value = f'JMS DATA FOR PERFORMA INVOICE NO. {iv_number}'
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal="center")
    apply_border(title_cell) # Note: apply_border on merged cell only styles the first cell
    
    sum_heads = ['DESCRIPTION', 'CODE', 'RATE', 'Sum of QTY', 'Sum of AMOUNT']
    sum_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for i, h in enumerate(sum_heads, 2):
        cell = ws_sum.cell(row=3, column=i)
        cell.value = h
        cell.fill = sum_fill
        cell.font = Font(bold=True)
        apply_border(cell)

    summary_data = defaultdict(lambda: {'qty': 0.0, 'amount': 0.0})
    for r in all_rows:
        key = (r['description'].strip(), r['sap_code'], r['rate'])
        summary_data[key]['qty'] += r['qty']
        summary_data[key]['amount'] += r['amount']
    
    sorted_keys = sorted(summary_data.keys(), key=lambda x: x[0])
    curr_row = 4
    total_qty = 0.0
    total_amt = 0.0
    for key in sorted_keys:
        desc, code, rate = key
        q, a = summary_data[key]['qty'], summary_data[key]['amount']
        ws_sum.cell(row=curr_row, column=2).value = desc
        ws_sum.cell(row=curr_row, column=3).value = code
        ws_sum.cell(row=curr_row, column=4).value = rate
        ws_sum.cell(row=curr_row, column=5).value = q
        ws_sum.cell(row=curr_row, column=6).value = a
        for c in range(2, 7): apply_border(ws_sum.cell(row=curr_row, column=c))
        total_qty += q
        total_amt += a
        curr_row += 1
    
    total_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    ws_sum.cell(row=curr_row, column=2).value = 'Grand Total'
    ws_sum.cell(row=curr_row, column=5).value = total_qty
    ws_sum.cell(row=curr_row, column=6).value = total_amt
    for c in range(2, 7):
        cell = ws_sum.cell(row=curr_row, column=c)
        cell.fill = total_fill
        cell.font = Font(bold=True)
        apply_border(cell)
        if c >= 5: cell.number_format = '#,##0.00'

    out_wb.save(output_path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--files", nargs='+', required=True)
    parser.add_argument("--mindump", required=True)
    parser.add_argument("--iv_number", required=True)
    parser.add_argument("--activity", default='A6')
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    generate_performa_invoice(args.files, args.mindump, args.iv_number, args.activity, args.output)
