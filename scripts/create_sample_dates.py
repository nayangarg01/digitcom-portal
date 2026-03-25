import pandas as pd
import os

def create_sample_dates():
    old_file = 'RoutingSampleFiles/DATA-KM-NAYAN_WHDetails.xlsx'
    new_file = 'RoutingSampleFiles/Optimized_Dispatch_Final_Master.xlsx'
    target_dir = 'routing/test_data_dates'
    
    # Read files
    df_old = pd.read_excel(old_file)
    df_new = pd.read_excel(new_file)
    
    # Standardize column naming for merge
    # Old cols: ['eNBsiteID', 'LAT ', 'LONG', 'BILLING FILE', 'CLUBBING', 'AKTBC', 'WH ']
    # New cols: ['SITE ID', ..., 'min_date', 'CLUBBING', 'AKTBC']
    
    # Rename columns in old sheet to avoid collision and clarify source
    df_old = df_old.rename(columns={'AKTBC': 'AKTBC_OLD', 'CLUBBING': 'CLUBBING_OLD'})
    
    # We'll merge the entire df_new onto df_old, but we need min_date from df_new
    # SITE ID in df_new corresponds to eNBsiteID in df_old
    
    # Let's subset df_new to what we need for the comparison but keep 'everything' as requested
    # Actually, let's merge them properly.
    merged = df_old.merge(df_new, left_on='eNBsiteID', right_on='SITE ID', how='inner')
    
    # 'everything' means we have columns from both.
    # In df_new, AKTBC and CLUBBING are the new ones.
    # We should rename them to AKTBC_NEW and CLUBBING_NEW to be clear.
    merged = merged.rename(columns={'AKTBC': 'AKTBC_NEW', 'CLUBBING': 'CLUBBING_NEW'})
    
    # Get top 5 dates by site count
    if 'min_date' in merged.columns:
        date_counts = merged['min_date'].value_counts()
        top_dates = date_counts.head(5).index.tolist()
        
        for d in top_dates:
            date_str = str(d).split(' ')[0]
            subset = merged[merged['min_date'] == d]
            
            # Priority columns first, then everything else
            cols = list(merged.columns)
            priority_cols = ['eNBsiteID', 'min_date', 'AKTBC_OLD', 'AKTBC_NEW', 'CLUBBING_OLD', 'CLUBBING_NEW']
            remaining = [c for c in cols if c not in priority_cols]
            
            final_subset = subset[priority_cols + remaining]
            
            filename = os.path.join(target_dir, f"{date_str}.xlsx")
            final_subset.to_excel(filename, index=False)
            print(f"Created {filename} with {len(subset)} sites.")
    else:
        print("Error: min_date column not found in new data.")

if __name__ == "__main__":
    create_sample_dates()
