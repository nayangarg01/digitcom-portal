import pdfplumber
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import argparse
import sys
import os

def parse_work_order(pdf_path, output_excel):
    print(f"Parsing {pdf_path}...")
    
    records = []
    
    current_site_index = None
    current_site_id = None
    current_site_value = None
    current_line_item = None
    
    # Global WO Summary
    wo_summary = {
        "Vendor Name": None,
        "Work Order No.": None,
        "Work Order Date": None,
        "WO Period From": None,
        "WO Period To": None,
        "Value of Work (INR)": None,
        "CGST (INR)": None,
        "SGST (INR)": None,
        "Total Order Value (INR)": None
    }
    
    # Global Regex Patterns
    vendor_name_pattern = re.compile(r'^(.*?)\s+Date\s*:')
    wo_no_pattern = re.compile(r'Work OrderNo\.\s*:\s*(\S+)')
    wo_date_pattern = re.compile(r'Date\s*:\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4})')
    wo_from_pattern = re.compile(r'WOPeriod From DT\s*:\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4})')
    wo_to_pattern = re.compile(r'To DT\s*:\s*([0-9]{2}\.[0-9]{2}\.[0-9]{4})')
    val_of_work_pattern = re.compile(r'^ValueofWork\s+INR\s+([0-9,.]+)')
    cgst_pattern = re.compile(r'^CGST\s+INR\s+([0-9,.]+)')
    sgst_pattern = re.compile(r'^SGST\s+INR\s+([0-9,.]+)')
    total_val_pattern = re.compile(r'^TOTALORDERVALUE\s+INR\s+([0-9,.]+)')

    # Line Item Regex Patterns
    site_pattern = re.compile(r'^([0-9]+)\s+(5GSiteReadiness_[A-Z0-9\-]+)\s+([0-9]+)\s+AU')
    site_val_pattern = re.compile(r'ValueofWork\s+INR/AU\s+([0-9,.]+)')
    item_pattern = re.compile(r'^([0-9]{2})\s+([0-9]{7})\s+(.*?)\s+([0-9]+)\s*([A-Za-z]+)\s*-\s*(.*)')
    item_val_pattern = re.compile(r'Netvalueofitem\s+([0-9,.]+)\s+([0-9,.]+)')
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text:
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    line = line.strip()
                    
                    # Extract Summary Data
                    if not wo_summary["Vendor Name"]:
                        m = vendor_name_pattern.search(line)
                        if m: wo_summary["Vendor Name"] = m.group(1).strip()
                    if not wo_summary["Work Order No."]:
                        m = wo_no_pattern.search(line)
                        if m: wo_summary["Work Order No."] = m.group(1)
                    if not wo_summary["Work Order Date"]:
                        m = wo_date_pattern.search(line)
                        if m: wo_summary["Work Order Date"] = m.group(1)
                    if not wo_summary["WO Period From"]:
                        m = wo_from_pattern.search(line)
                        if m: wo_summary["WO Period From"] = m.group(1)
                    if not wo_summary["WO Period To"]:
                        m = wo_to_pattern.search(line)
                        if m: wo_summary["WO Period To"] = m.group(1)
                    if not wo_summary["Value of Work (INR)"]:
                        m = val_of_work_pattern.search(line)
                        if m: wo_summary["Value of Work (INR)"] = float(m.group(1).replace(',', ''))
                    if not wo_summary["CGST (INR)"]:
                        m = cgst_pattern.search(line)
                        if m: wo_summary["CGST (INR)"] = float(m.group(1).replace(',', ''))
                    if not wo_summary["SGST (INR)"]:
                        m = sgst_pattern.search(line)
                        if m: wo_summary["SGST (INR)"] = float(m.group(1).replace(',', ''))
                    if not wo_summary["Total Order Value (INR)"]:
                        m = total_val_pattern.search(line)
                        if m: wo_summary["Total Order Value (INR)"] = float(m.group(1).replace(',', ''))

                    # 1. Match Site
                    site_match = site_pattern.match(line)
                    if site_match:
                        current_site_index = site_match.group(1)
                        current_site_id = site_match.group(2)
                        current_site_value = None
                        continue
                        
                    # 2. Match Site Value
                    if current_site_id and not current_site_value:
                        val_match = site_val_pattern.search(line)
                        if val_match:
                            current_site_value = val_match.group(1).replace(',', '')
                            continue
                            
                    # 3. Match Line Item
                    item_match = item_pattern.match(line)
                    if item_match:
                        current_line_item = {
                            "Site Index": current_site_index,
                            "Site ID": current_site_id,
                            "Site Base Value (INR)": float(current_site_value) if current_site_value else 0.0,
                            "Line Item No": item_match.group(1),
                            "Item Code": item_match.group(2),
                            "Item Description": item_match.group(3).strip(),
                            "Quantity": int(item_match.group(4)),
                            "Unit": item_match.group(5),
                            "Unit Rate": None,
                            "Total Amount": None
                        }
                        continue
                        
                    # 4. Match Line Item Value
                    if current_line_item and current_line_item["Unit Rate"] is None:
                        item_val_match = item_val_pattern.search(line)
                        if item_val_match:
                            current_line_item["Unit Rate"] = float(item_val_match.group(1).replace(',', ''))
                            current_line_item["Total Amount"] = float(item_val_match.group(2).replace(',', ''))
                            records.append(current_line_item)
                            current_line_item = None
                            
    except Exception as e:
        print(f"Error reading PDF: {e}")
        sys.exit(1)
        
    print(f"Extracted {len(records)} line items.")
    
    if not records:
        print("No records found. Regex might need adjustment.")
        return
        
    # Generate Excel using openpyxl
    print(f"Saving to {output_excel}...")
    wb = Workbook()
    
    # Create the Work Order Details sheet first
    ws_details = wb.active
    ws_details.title = "Work Order Details"
    
    ws_details.append(["Work Order Detail", "Value"])
    for key, val in wo_summary.items():
        ws_details.append([key, val])
        
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for cell in ws_details[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    for row in range(2, len(wo_summary) + 2):
        ws_details[f"A{row}"].font = Font(bold=True)
        if isinstance(ws_details[f"B{row}"].value, float):
            ws_details[f"B{row}"].number_format = '#,##0.00'
            
    ws_details.column_dimensions['A'].width = 30
    ws_details.column_dimensions['B'].width = 30

    # Now create the line items sheet
    ws = wb.create_sheet(title="Parsed Work Order")
    
    headers = list(records[0].keys())
    ws.append(headers)
    
    for record in records:
        ws.append([record[k] for k in headers])
    
    # Formatting
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        
    ws.freeze_panes = 'A2'
    
    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = get_column_letter(col[0].column)
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        if adjusted_width > 50:
            adjusted_width = 50
        ws.column_dimensions[column].width = adjusted_width
        
    # Highlight Total Amount Column
    total_amt_col_idx = headers.index("Total Amount") + 1
    total_amt_col_letter = get_column_letter(total_amt_col_idx)
    highlight_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    
    for row in range(2, len(records) + 2):
        cell = ws[f"{total_amt_col_letter}{row}"]
        cell.fill = highlight_fill
        cell.number_format = '#,##0.00'
        
    # Format other numbers
    site_val_col_idx = headers.index("Site Base Value (INR)") + 1
    site_val_col_letter = get_column_letter(site_val_col_idx)
    rate_col_idx = headers.index("Unit Rate") + 1
    rate_col_letter = get_column_letter(rate_col_idx)
    
    for row in range(2, len(records) + 2):
        ws[f"{site_val_col_letter}{row}"].number_format = '#,##0.00'
        ws[f"{rate_col_letter}{row}"].number_format = '#,##0.00'
        
    # Merge Site columns for identical sites (Site Index, Site ID, Site Base Value)
    start_row = 2
    site_id_col_idx = headers.index("Site ID") + 1
    current_site = None
    
    for row in range(2, len(records) + 3):
        if row <= len(records) + 1:
            site_id = ws.cell(row=row, column=site_id_col_idx).value
        else:
            site_id = None
            
        if site_id != current_site:
            if current_site is not None:
                end_row = row - 1
                if end_row > start_row:
                    # Merge columns 1 (Site Index), 2 (Site ID), and 3 (Site Base Value)
                    for col_idx in [1, 2, 3]:
                        ws.merge_cells(start_row=start_row, start_column=col_idx, end_row=end_row, end_column=col_idx)
                        # Center align the merged cell
                        ws.cell(row=start_row, column=col_idx).alignment = Alignment(horizontal='center', vertical='top')
            current_site = site_id
            start_row = row

    wb.save(output_excel)
    print(f"Success! Saved formatted Excel to {output_excel}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Parse Work Order PDF to Excel")
    parser.add_argument('pdf_path', nargs='?', default='DIGITCOM_630344375.pdf', help='Path to input PDF')
    parser.add_argument('--output', '-o', default='WorkOrder_Summary.xlsx', help='Output Excel file path')
    
    args = parser.parse_args()
    parse_work_order(args.pdf_path, args.output)
