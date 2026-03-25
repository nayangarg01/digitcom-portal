import pandas as pd
import os
import math

# Warehouse Coords (Jaipur - Bagru, Jodhpur - Mogra Khurd)
WAREHOUSE_COORDS = {
    'JAIPUR': (26.8139, 75.5450),
    'JODHPUR': (26.1245, 73.0543),
    'DEFAULT': (26.8139, 75.5450)
}

def get_distance(p1, p2):
    return math.sqrt((p1[0] - p2[0])**2 + (p1[1] - p2[1])**2) * 111.0

def partition_sizes(n):
    if n < 2: return [n]
    if n == 2: return [2]
    if n == 3: return [3]
    if n == 4: return [2, 2]
    if n % 3 == 0: return [3] * (n // 3)
    if n % 3 == 1: return [3] * ((n - 4) // 3) + [2, 2]
    return [3] * ((n - 2) // 3) + [2]

def run_routing_on_df(df_batch, wh_coords):
    sites = []
    for idx, row in df_batch.iterrows():
        sites.append({
            'id': str(row['SITE ID']),
            'coords': (float(row['latitude']), float(row['longitude'])),
            'cmp': str(row['cmp']).strip(),
            'orig_row': row.to_dict()
        })
    
    # Group by CMP
    cmp_groups = {}
    for s in sites:
        c = s['cmp']
        if c not in cmp_groups: cmp_groups[c] = []
        cmp_groups[c].append(s)
    
    final_routes = []
    for cmp_name, group_sites in cmp_groups.items():
        group_sites.sort(key=lambda x: get_distance(wh_coords, x['coords']))
        sizes = partition_sizes(len(group_sites))
        curr = 0
        for size in sizes:
            chunk = group_sites[curr : curr + size]
            p_sites = chunk.copy()
            while p_sites:
                if len(p_sites) == 1:
                    final_routes.append([p_sites.pop(0)])
                elif len(p_sites) == 2:
                    A, B = p_sites[0], p_sites[1]
                    if get_distance(wh_coords, B['coords']) < get_distance(A['coords'], B['coords']):
                        final_routes.append([p_sites.pop(0)])
                    else:
                        final_routes.append([p_sites.pop(0), p_sites.pop(0)])
                else: # 3
                    A, B, C = p_sites[0], p_sites[1], p_sites[2]
                    if get_distance(wh_coords, B['coords']) < get_distance(A['coords'], B['coords']):
                        final_routes.append([p_sites.pop(0)])
                    elif get_distance(wh_coords, C['coords']) < get_distance(B['coords'], C['coords']):
                        final_routes.append([p_sites.pop(0), p_sites.pop(0)])
                    else:
                        final_routes.append([p_sites.pop(0), p_sites.pop(0), p_sites.pop(0)])
            curr += size
    
    rows = []
    for r_idx, route in enumerate(final_routes):
        label = f"R{r_idx+1}"
        prev_coords = wh_coords
        for s_idx, site in enumerate(route):
            dist = get_distance(prev_coords, site['coords'])
            if s_idx == 0:
                dist = max(0, dist - 50)
            
            row = site['orig_row'].copy()
            row['AKTBC_NEW'] = round(dist, 2)
            row['CLUBBING_NEW'] = f"{label}-S{s_idx+1}"
            rows.append(row)
            prev_coords = site['coords']
            
    return pd.DataFrame(rows)

def process():
    old_file = 'RoutingSampleFiles/DATA-KM-NAYAN_WHDetails.xlsx'
    new_file = 'RoutingSampleFiles/Optimized_Dispatch_Final_Master.xlsx'
    target_dir = 'RoutingSampleFiles/test_data_dates'
    
    df_old = pd.read_excel(old_file)
    df_new = pd.read_excel(new_file)
    
    # Filter only those present in both (to get AKTBC_OLD)
    merged = df_old.rename(columns={'AKTBC': 'AKTBC_OLD', 'CLUBBING': 'CLUBBING_OLD'}).merge(
        df_new, left_on='eNBsiteID', right_on='SITE ID', how='inner'
    )
    
    # Identify 5 dates with 10-20 sites
    date_counts = merged['min_date'].value_counts()
    eligible_dates = date_counts[(date_counts >= 10) & (date_counts <= 20)].head(5).index.tolist()
    
    # If not enough, loosen constraints
    if len(eligible_dates) < 5:
        eligible_dates = date_counts.sort_values(ascending=False).head(5).index.tolist()
        print("Warning: Could not find exactly 10-20 site dates, taking most frequent.")

    for d in eligible_dates:
        date_str = str(d).split(' ')[0]
        subset = merged[merged['min_date'] == d].copy()
        
        # Determine warehouse for this batch (take first)
        wh_name = str(subset.iloc[0]['Warehouse_Name']).upper().strip()
        wh_coords = WAREHOUSE_COORDS.get(wh_name, WAREHOUSE_COORDS['DEFAULT'])
        
        # Re-run routing logic for this subset
        result_df = run_routing_on_df(subset, wh_coords)
        
        # Prioritize and Sort
        priority = ['eNBsiteID', 'min_date', 'AKTBC_OLD', 'AKTBC_NEW', 'CLUBBING_OLD', 'CLUBBING_NEW']
        others = [c for c in result_df.columns if c not in priority]
        
        # Natural sorting for CLUBBING_OLD (e.g., A1, A2, A10, B1)
        import re
        def natural_sort_key(s):
            return [int(text) if text.isdigit() else text.lower()
                    for text in re.split('([0-9]+)', str(s))]

        result_df['sort_key'] = result_df['CLUBBING_OLD'].apply(natural_sort_key)
        final_result = result_df.sort_values(by='sort_key').drop(columns=['sort_key'])
        
        filename = os.path.join(target_dir, f"{date_str}_Small.xlsx")
        final_result[priority + others].to_excel(filename, index=False)
        print(f"Created {filename} with {len(final_result)} sites sorted by CLUBBING_OLD.")

if __name__ == "__main__":
    process()
