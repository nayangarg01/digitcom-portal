import sys
import os
import argparse
import xlsxwriter

# Add BillingEngine_OOP to python search path
script_dir = os.path.dirname(os.path.abspath(__file__))
oop_dir = os.path.abspath(os.path.join(script_dir, "..", "..", "BillingEngine_OOP"))
sys.path.append(oop_dir)

try:
    from data_loader import DataFactory
    import oop_billing_generator
except ImportError as e:
    print(f"CRITICAL: Failed to import BillingEngine_OOP. Ensure the BillingEngine_OOP directory exists. Error: {e}")
    sys.exit(1)

def main():
    print("=== OOP BILLING GENERATION STARTED ===")
    parser = argparse.ArgumentParser()
    parser.add_argument("dc_number", help="Target DC Number (e.g. DC0122)")
    parser.add_argument("--output", default=None, help="Output Excel file path")
    parser.add_argument("--activity", default=None, choices=['A6', 'A6_B6'], help="Billing activity override")
    args = parser.parse_args()

    dc_number = args.dc_number.strip().upper()
    output_path = args.output if args.output else f"Billing/{dc_number}_Unified_Billing.xlsx"
    activity = args.activity

    print(f"LOG: Target DC Number: {dc_number}")
    print(f"LOG: Output Path: {output_path}")

    # Load DataFactory. Pass None for master_path because we are loading from database only.
    print("LOG: Initializing DataFactory and loading OOP database...")
    try:
        factory = DataFactory(None)
    except Exception as e:
        print(f"ERROR: Failed to load OOP database: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    print(f"LOG: Database loaded successfully. Total registered sites: {len(factory.sites)}")

    # Filter sites belonging to target dc_number
    print(f"LOG: Filtering database sites for DC Number '{dc_number}'...")
    dc_sites = [s for s in factory.sites.values() if str(s.dc_no).strip().upper() == dc_number]

    if not dc_sites:
        print(f"ERROR: No site data found in database for DC Number: {dc_number}")
        print("TIP: Please make sure you have synchronized the Master Tracker containing this DC Number first!")
        sys.exit(1)

    print(f"LOG: Found {len(dc_sites)} sites in database for DC {dc_number}.")
    for s in dc_sites:
        print(f"  - Site: {s.site_id}, PMP: {s.pmp_id}, Activity: {s.activity_type}, MINs: {s.min_no}")

    # Determine activity type
    if not activity:
        if any(s.activity_type == 'A6+B6' for s in dc_sites):
            activity = 'A6_B6'
        else:
            activity = 'A6'
    print(f"LOG: Selected billing layout: {activity}")

    # Get WO number from site models
    wo_numbers = list(set(str(s.wo) for s in dc_sites if s.wo and s.wo != 'N/A'))
    wo_number = wo_numbers[0] if wo_numbers else 'N/A'
    print(f"LOG: Found Work Order Number(s): {', '.join(wo_numbers) if wo_numbers else 'N/A'}")

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    print("LOG: Initializing XlsxWriter workbook...")
    try:
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

            print("LOG: Writing Main WCC placeholder sheet...")
            oop_billing_generator.write_main_wcc_placeholder(wb, formats)

            print("LOG: Writing WCC sheet...")
            oop_billing_generator.write_wcc(wb, dc_sites, dc_number, formats, activity=activity, wo_number=wo_number)

            print("LOG: Writing JMS sheet...")
            oop_billing_generator.write_matrix_sheet(wb, 'JMS', dc_sites, dc_number, formats, include_amounts=True, activity=activity, wo_number=wo_number)

            print("LOG: Writing Abstract sheet...")
            oop_billing_generator.write_matrix_sheet(wb, 'Abstract', dc_sites, dc_number, formats, include_amounts=True, activity=activity, wo_number=wo_number)

            print("LOG: Writing BOQ sheet...")
            oop_billing_generator.write_matrix_sheet(wb, 'BOQ', dc_sites, dc_number, formats, include_amounts=True, activity=activity, wo_number=wo_number)

            print("LOG: Writing Declaration sheet...")
            oop_billing_generator.write_declaration(wb, dc_sites, dc_number, formats, activity=activity, wo_number=wo_number)

            print("LOG: Writing Annexure & Reco sheets...")
            oop_billing_generator.write_annexure_and_reco(wb, dc_sites, dc_number, formats, activity=activity, wo_number=wo_number)

        print("LOG: XlsxWriter workbook closed.")
    except Exception as e:
        print(f"ERROR: Failed to write programmatic sheets: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # Resolve static template path
    ref_template = os.path.join(script_dir, '..', 'templates', 'billing_template.xlsx')
    if not os.path.exists(ref_template):
        print(f"ERROR: Static billing template not found at: {ref_template}")
        sys.exit(1)

    print("LOG: Injecting Main WCC sheet template and checkboxes using Hybrid Writer...")
    try:
        oop_billing_generator.inject_main_wcc_template(output_path, ref_template, dc_sites, dc_number, wo_number)
    except Exception as e:
        print(f"ERROR: Failed to inject Main WCC template: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    print("=== OOP BILLING GENERATION COMPLETED SUCCESSFULLY ===")
    print(f"SUCCESS: Generated sheet available at: {output_path}")

if __name__ == "__main__":
    main()
