import sys
import os
import argparse
import pandas as pd
import xlsxwriter

# Add BillingEngine_OOP to python search path
script_dir = os.path.dirname(os.path.abspath(__file__))
oop_dir = os.path.abspath(os.path.join(script_dir, "..", "..", "BillingEngine_OOP"))
sys.path.append(oop_dir)

try:
    from data_loader import DataFactory, TEMPLATE_ITEMS, TEMPLATE_ITEMS_A6_B6
except ImportError as e:
    print(f"CRITICAL: Failed to import BillingEngine_OOP. Ensure the BillingEngine_OOP directory exists. Error: {e}")
    sys.exit(1)

def main():
    print("=== OOP PERFORMA INVOICE GENERATION STARTED ===")
    parser = argparse.ArgumentParser()
    parser.add_argument("iv_number", help="Performa Invoice Number")
    parser.add_argument("dc_numbers", nargs="+", help="List of DC numbers to club")
    parser.add_argument("--output", default=None, help="Output Excel file path")
    parser.add_argument("--activity", default="A6", choices=["A6", "A6_B6"], help="Overall activity layout")
    args = parser.parse_args()

    iv_number = args.iv_number.strip()
    target_dcs = [dc.strip().upper() for dc in args.dc_numbers]
    output_path = args.output if args.output else f"Billing/Performa_{iv_number}.xlsx"
    activity = args.activity

    print(f"LOG: Performa Invoice Number: {iv_number}")
    print(f"LOG: Target DC Numbers: {', '.join(target_dcs)}")
    print(f"LOG: Output Path: {output_path}")

    # Load DataFactory
    print("LOG: Initializing DataFactory and loading OOP database...")
    try:
        factory = DataFactory(None)
    except Exception as e:
        print(f"ERROR: Failed to load OOP database: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    print(f"LOG: Database loaded successfully. Total registered sites: {len(factory.sites)}")

    # Filter sites that belong to target DCs
    dc_sites = [s for s in factory.sites.values() if str(s.dc_no).strip().upper() in target_dcs]

    if not dc_sites:
        print(f"ERROR: No site data found in database for DC Numbers: {', '.join(target_dcs)}")
        sys.exit(1)

    print(f"LOG: Found {len(dc_sites)} sites in database matching target DCs.")

    # Determine nature of work
    nature_of_work = "AIR FIBER INSTALLATION"
    if activity == 'A6_B6':
        nature_of_work = "AIR FIBER INSTALLATION(A6+B6)"

    # Process each site and compile items
    all_rows = []
    for site in dc_sites:
        print(f"LOG: Processing site {site.site_id} (DC: {site.dc_no}, WO: {site.wo}, WBS: {site.wbs_id})")
        
        # Decide which templates are applicable
        # If site is A6+B6 activity type, check both A6 and A6+B6 items
        applicable_templates = TEMPLATE_ITEMS.copy()
        if site.activity_type == 'A6+B6':
            applicable_templates += TEMPLATE_ITEMS_A6_B6

        for item in applicable_templates:
            sap_code = str(item['sap']).strip()
            desc = str(item['desc']).strip()
            rate = float(item['rate'])
            
            # Check quantity
            qty = float(site.items.get(sap_code, 0.0))
            if qty > 0:
                site_amount = qty * rate
                all_rows.append({
                    'vendor': 'DIGITCOM INDIA TECHNOLOGIES',
                    'scope': 'A6' if site.activity_type == 'A6' else 'A6_B6',
                    'iv_no': f"PERFORMA INVOICE NO. {iv_number}",
                    'wo_no': str(site.wo).strip(),
                    'site': site.site_id,
                    'wbs': site.wbs_id if site.wbs_id else 'N/A',
                    'sap_code': sap_code,
                    'description': desc,
                    'nature': nature_of_work,
                    'qty': qty,
                    'rate': rate,
                    'amount': site_amount
                })

    if not all_rows:
        print("ERROR: No data rows were extracted from database. Sheet will be empty.")
        sys.exit(1)

    print(f"LOG: Extracted {len(all_rows)} item-site entries. Writing Excel workbook...")
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    try:
        with xlsxwriter.Workbook(output_path) as wb:
            # Formats
            f_header = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9', 'font_color': 'black', 'font_size': 10})
            f_cell = wb.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter', 'font_size': 9})
            f_num = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9})
            f_amount = wb.add_format({'border': 1, 'align': 'right', 'valign': 'vcenter', 'font_size': 9, 'num_format': '#,##0'})
            f_title = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
            
            f_sum_title = wb.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            f_sum_head = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00', 'font_color': 'black'})
            f_sum_total = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#00B050', 'font_color': 'black'})
            f_sum_total_num = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#00B050', 'font_color': 'black', 'num_format': '#,##0'})

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
            ws_sum.set_column('A:A', 5)
            ws_sum.set_column('B:B', 50) # Description
            ws_sum.set_column('C:C', 15) # Code
            ws_sum.set_column('D:D', 15) # Rate
            ws_sum.set_column('E:E', 15) # Qty
            ws_sum.set_column('F:F', 20) # Amount
            
            ws_sum.merge_range('B2:F2', f'JMS DATA FOR PERFORMA INVOICE NO. {iv_number}', f_sum_title)
            
            sum_heads = ['DESCRIPTION', 'CODE', 'RATE', 'Sum of QTY', 'Sum of AMOUNT']
            for i, h in enumerate(sum_heads):
                ws_sum.write(2, i + 1, h, f_sum_head)
                
            df_all = pd.DataFrame(all_rows)
            if not df_all.empty:
                # Group by description, code, and rate
                pivot = df_all.groupby(['description', 'sap_code', 'rate']).agg({'qty': 'sum', 'amount': 'sum'}).reset_index()
                pivot = pivot.sort_values('description')
                
                curr_row = 3
                for _, row in pivot.iterrows():
                    ws_sum.write(curr_row, 1, row['description'], f_cell)
                    ws_sum.write(curr_row, 2, row['sap_code'], f_num)
                    ws_sum.write(curr_row, 3, row['rate'], f_num)
                    ws_sum.write(curr_row, 4, row['qty'], f_num)
                    ws_sum.write(curr_row, 5, row['amount'], f_amount)
                    curr_row += 1
                
                # Grand Total Row
                ws_sum.write(curr_row, 1, 'Grand Total', f_sum_total)
                ws_sum.write(curr_row, 2, '', f_sum_total)
                ws_sum.write(curr_row, 3, '', f_sum_total)
                ws_sum.write_formula(curr_row, 4, f"=SUM(E4:E{curr_row})", f_sum_total_num)
                ws_sum.write_formula(curr_row, 5, f"=SUM(F4:F{curr_row})", f_sum_total_num)

        print("=== OOP PERFORMA INVOICE GENERATION COMPLETED SUCCESSFULLY ===")
        print(f"SUCCESS: Generated sheet available at: {output_path}")
    except Exception as e:
        print(f"ERROR: Failed to write performa sheets: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
