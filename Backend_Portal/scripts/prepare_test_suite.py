import pandas as pd
import os
import json

def prepare_test_suite():
    input_file = '../RoutingSampleFiles/DATA-KM-NAYAN_withWareHouse.xlsx'
    output_dir = '../RoutingSampleFiles/test_data'
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"Reading {input_file}...")
    df = pd.read_excel(input_file)
    
    # Fill forward the BILLING FIL (DC group) and the last column (Warehouse)
    # The last column is unnamed in some cases, so we use iloc
    df.iloc[:, 3] = df.iloc[:, 3].ffill() # BILLING FIL
    df.iloc[:, -1] = df.iloc[:, -1].ffill() # Warehouse
    
    # Rename columns for clarity if needed, but we mostly care about the content
    groups = df.groupby(df.iloc[:, 3])
    
    warehouse_map = {
        'Jaipur': ('Jaipur - Bagru', 26.8139, 75.5450),
        'Jodhpur': ('Jodhpur - Mogra Khurd', 26.1245, 73.0543),
        'Lucknow': ('Lucknow - Safedabad', 26.8906, 81.0558)
    }
    
    task_metadata = []
    
    for group_id, group_df in groups:
        if pd.isna(group_id): continue
        group_id_str = str(group_id).strip()
        
        # Filter out rows with NaN in LAT or LONG
        clean_df = group_df.dropna(subset=[group_df.columns[1], group_df.columns[2]])
        
        if clean_df.empty:
            print(f"Skipping {group_id_str} (no valid coordinates)")
            continue
            
        # Get warehouse name from the last column
        raw_warehouse = str(clean_df.iloc[0, -1]).strip()
        warehouse_info = warehouse_map.get(raw_warehouse)
        
        if not warehouse_info:
            # Try partial match
            found = False
            for k, v in warehouse_map.items():
                if k.lower() in raw_warehouse.lower():
                    warehouse_info = v
                    found = True
                    break
            if not found:
                print(f"Unknown warehouse '{raw_warehouse}' for {group_id_str}. Defaulting to Jaipur.")
                warehouse_info = warehouse_map['Jaipur']
        
        # Store individual Excel file
        group_filename = f"{group_id_str}.xlsx"
        group_path = os.path.join(output_dir, group_filename)
        clean_df.to_excel(group_path, index=False)
        
        # Calculate manual sum (ignore NaN in AKTBC)
        manual_sum = clean_df.iloc[:, 5].sum()
        
        task_metadata.append({
            "group_id": group_id_str,
            "filename": group_filename,
            "warehouse": warehouse_info[0],
            "lat": warehouse_info[1],
            "lng": warehouse_info[2],
            "manual_aktbc_sum": round(float(manual_sum), 2),
            "num_sites": len(clean_df)
        })
    
    # Save metadata for the benchmark runner
    with open(os.path.join(output_dir, 'suite_metadata.json'), 'w') as f:
        json.dump(task_metadata, f, indent=4)
        
    print(f"Successfully split into {len(task_metadata)} test files in {output_dir}")

if __name__ == "__main__":
    prepare_test_suite()
