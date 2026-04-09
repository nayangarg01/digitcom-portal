import pandas as pd
import sys

OUTPUT_FILE = "KM_A6_MM_AutoValidated.xlsx"
print(f"Loading {OUTPUT_FILE} for visual styling injection...")

try:
    df_manual = pd.read_excel(OUTPUT_FILE).fillna("")
except Exception as e:
    print("Failed to load: ", e)
    sys.exit(1)

cols = list(df_manual.columns)

print("Applying thick-bordered headers, backgrounds, and intelligent column widths...")
with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
    wb = writer.book
    ws = wb.add_worksheet("Automation Audit Matrix")
    
    # Define formats natively
    header_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#2F5597', 'font_color': 'white', 'border': 1})
    std_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    auto_fmt = wb.add_format({'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    auto_num_fmt = wb.add_format({'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0.00'})
    
    # Write Headers
    for c_idx, col_name in enumerate(cols):
        ws.write(0, c_idx, str(col_name), header_fmt)
        
    # Write Rows
    for r_idx, row in df_manual.iterrows():
        row_dict = row.to_dict()
        for c_idx, col_name in enumerate(cols):
            val = row_dict.get(col_name, "")
            
            # Formatting overrides
            if col_name == 'AUTO ROUTE CLUBBING':
                ws.write(r_idx + 1, c_idx, val, auto_fmt)
            elif col_name in ['AUTO DISTANCE WH', 'AUTO AKTBC']:
                try: val = float(val)
                except: pass
                ws.write(r_idx + 1, c_idx, val, auto_num_fmt)
            else:
                ws.write(r_idx + 1, c_idx, val, std_fmt)
                
    # Column Widths Dynamic Expansion
    for c_idx, col_name in enumerate(cols):
        max_len = max([len(str(r.get(col_name, ""))) for _, r in df_manual.iterrows()], default=0)
        max_len = max(max_len, len(str(col_name))) + 4
        max_len = min(max_len, 40) # cap
        ws.set_column(c_idx, c_idx, max_len)

print("✅ Visual styling fully bound.")
