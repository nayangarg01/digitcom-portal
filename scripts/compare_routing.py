import pandas as pd
import os

def compare_routing_data():
    old_file = 'RoutingSampleFiles/DATA-KM-NAYAN_WHDetails.xlsx'
    new_file = 'RoutingSampleFiles/Optimized_Dispatch_Final_Master.xlsx'
    output_file = 'RoutingSampleFiles/AKTBC_Comparison.xlsx'

    if not os.path.exists(old_file) or not os.path.exists(new_file):
        print("Error: Missing source files.")
        return

    # Read old data
    df_old = pd.read_excel(old_file)
    # Read new data
    df_new = pd.read_excel(new_file)

    # Prepare for merge
    # Old cols: ['eNBsiteID', 'LAT ', 'LONG', 'BILLING FILE', 'CLUBBING', 'AKTBC', 'WH ']
    # New cols: ['SITE ID', ..., 'CLUBBING', 'AKTBC']
    
    # Standardize column naming for merging
    df_old_subset = df_old.copy()
    df_old_subset = df_old_subset.rename(columns={'AKTBC': 'AKTBC_OLD', 'CLUBBING': 'CLUBBING_OLD'})

    df_new_subset = df_new[['SITE ID', 'AKTBC', 'CLUBBING']].copy()
    df_new_subset.columns = ['SITE ID', 'AKTBC_NEW', 'CLUBBING_NEW']

    # Merge on SITE ID
    comparison_df = df_old_subset.merge(df_new_subset, left_on='eNBsiteID', right_on='SITE ID', how='left')

    # Reorder columns to show AKTBC and CLUBBING side-by-side
    final_cols = [
        'eNBsiteID', 
        'AKTBC_OLD', 'AKTBC_NEW', 
        'CLUBBING_OLD', 'CLUBBING_NEW',
        'WH ', 'LAT ', 'LONG'
    ]
    
    # Check which of these actually exist in the merged df
    existing_cols = [c for c in final_cols if c in comparison_df.columns]
    
    # Save to Excel
    comparison_df[existing_cols].to_excel(output_file, index=False)
    print(f"Comparison file created at {output_file}")

if __name__ == "__main__":
    compare_routing_data()
