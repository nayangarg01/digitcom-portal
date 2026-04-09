import pandas as pd
import os
import sys
import argparse

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
INPUT_FILE  = "km_calculated_all.xlsx"
OUTPUT_FILE = "grouped_by_date.xlsx"
DATE_COLUMN = "MIN DATE" # The primary date we want to group by
FALLBACK_DATE_COLUMN = "ATP11C"

def group_by_date(input_file, output_file):
    print(f"\nLoading data from: {input_file}")
    
    xl = pd.ExcelFile(input_file)
    master_df = []
    
    # Read every band's sheet
    for sheet in xl.sheet_names:
        df_sheet = pd.read_excel(xl, sheet_name=sheet)
        if df_sheet.empty: continue
        # Track the band internally so we can segregate
        df_sheet['_SourceBand'] = sheet.replace(' - Distances', '').strip()
        master_df.append(df_sheet)

    if not master_df:
        print("Error: No data found.")
        sys.exit(1)
        
    df = pd.concat(master_df, ignore_index=True)
    total_sites = len(df)
    print(f"Total rows loaded across all bands: {total_sites}")

    # ── 1. Group Data by Date ──
    # Ensure the date column exists
    active_date_col = DATE_COLUMN
    if active_date_col not in df.columns:
        if FALLBACK_DATE_COLUMN in df.columns:
            print(f"Warning: '{DATE_COLUMN}' column not found. Falling back to '{FALLBACK_DATE_COLUMN}'.")
            active_date_col = FALLBACK_DATE_COLUMN
        else:
            print(f"Error: Neither '{DATE_COLUMN}' nor '{FALLBACK_DATE_COLUMN}' column found in data.")
            sys.exit(1)

    # Handle missing/empty dates
    df[active_date_col] = df[active_date_col].fillna('Missing Date').astype(str)
    df[active_date_col] = df[active_date_col].str.strip()

    # We group by DATE and BAND to enforce strict segregation
    groups = df.groupby([active_date_col, '_SourceBand'])
    print(f"\nFound {len(groups)} unique Date+Band combinations.")

    # ── 2. Write to New Excel File ──
    print(f"\nSaving output to: {output_file}")
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        wb = writer.book
        
        # Formats for styling
        header_fmt = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})
        cell_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        dist_col_fmt = wb.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'num_format': '0.00', 'fg_color': '#FFF2CC'
        })
        
        for (date_val, band_val), group_df in groups:
            # Generate a clean sheet name (Max 31 chars)
            band_short = band_val.replace(" I&C", "").replace(" WAVE", "").strip()
            sheet_name = f"{date_val}_{band_short}"[:31].replace(":", "-").replace("/", "-")
            
            print(f"  Writing sheet '{sheet_name}' with {len(group_df)} records...")
            
            # Drop the internal tracking column before writing
            out_df = group_df.drop(columns=['_SourceBand'])
            out_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            ws = writer.sheets[sheet_name]
            
            # Apply formatting
            for i, col in enumerate(out_df.columns):
                # Calculate max length for column width
                max_len = max(out_df[col].astype(str).map(len).max(), len(str(col))) + 4
                max_len = max(min(max_len, 50), 15)

                if col == 'Distance from WH (km)':
                    ws.set_column(i, i, max_len, dist_col_fmt)
                else:
                    ws.set_column(i, i, max_len, cell_fmt)

            # Rewrite headers with header format
            for col_num, value in enumerate(out_df.columns.values):
                ws.write(0, col_num, value, header_fmt)

    print(f"\n✅ Done! File saved as '{output_file}'")

# ──────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Group the calculated distances sheet into separate sheets by date."
    )
    parser.add_argument("--input",  default=INPUT_FILE,  help="Input Excel file (e.g., km_calculated_v2.xlsx)")
    parser.add_argument("--output", default=OUTPUT_FILE, help="Output Excel file (e.g., grouped_by_date.xlsx)")
    
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' not found.")
        sys.exit(1)

    group_by_date(args.input, args.output)
