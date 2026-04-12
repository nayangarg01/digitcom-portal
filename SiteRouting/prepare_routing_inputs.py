import pandas as pd
import os

def prepare_routing_inputs(input_file, output_dir):
    print(f"Reading {input_file}...")
    df = pd.read_excel(input_file)
    
    # Ensure directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created directory: {output_dir}")

    # Convert MIN DATE to datetime
    df['MIN DATE'] = pd.to_datetime(df['MIN DATE'], errors='coerce')
    
    # Filter for year 2026
    df_2026 = df[df['MIN DATE'].dt.year == 2026].copy()
    
    if df_2026.empty:
        print("No data found for the year 2026.")
        return

    # Group by unique dates
    unique_dates = df_2026['MIN DATE'].dt.strftime('%Y-%m-%d').unique()
    print(f"Found {len(unique_dates)} unique dates in 2026.")

    for date_str in unique_dates:
        # Filter for the specific date
        daily_df = df_2026[df_2026['MIN DATE'].dt.strftime('%Y-%m-%d') == date_str].copy()
        
        # Prepare file name
        output_file = os.path.join(output_dir, f"Routing_Input_{date_str}.xlsx")
        
        print(f"  -> Processing {date_str} ({len(daily_df)} rows)...")

        # Create Excel writer with xlsxwriter engine
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            daily_df.to_excel(writer, index=False, sheet_name='Routing_Input')
            
            workbook  = writer.book
            worksheet = writer.sheets['Routing_Input']

            # Define Formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': False,
                'valign': 'vcenter',
                'align': 'center',
                'fg_color': '#2F5597',
                'font_color': 'white',
                'border': 1
            })
            
            cell_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })

            # Format Headers
            for col_num, value in enumerate(daily_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Apply cell format to all data
            for row_num in range(1, len(daily_df) + 1):
                for col_num in range(len(daily_df.columns)):
                    val = str(daily_df.iloc[row_num-1, col_num]) if not pd.isna(daily_df.iloc[row_num-1, col_num]) else ""
                    worksheet.write(row_num, col_num, val, cell_format)

            # Auto-adjust column widths
            for i, col in enumerate(daily_df.columns):
                # find maximum length of a cell in this column
                max_len = daily_df[col].astype(str).map(len).max()
                # setting the length to max_len + 5 for padding
                worksheet.set_column(i, i, max_len + 5)

    print(f"\nDone! All files are in {output_dir}")

if __name__ == "__main__":
    prepare_routing_inputs('KM VERIFICATION 3.xlsx', 'Processed_Inputs')
