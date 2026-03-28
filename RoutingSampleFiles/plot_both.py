import pandas as pd
import folium
import math
import os
import re
from pathlib import Path

# --- CONFIGURATION ---
WAREHOUSE_MAP = {
    'JAIPUR': (26.8139, 75.5450),
    'JODHPUR': (26.1245, 73.0543),
    'DEFAULT': (26.8139, 75.5450)
}

# Hex codes for Excel (Light versions of map colors)
COLOR_MAP = [
    '#FFCCCC', # Light Red
    '#CCE5FF', # Light Blue
    '#CCFFCC', # Light Green
    '#E5CCFF', # Light Purple
    '#FFE5CC', # Light Orange
    '#FF9999', # Slightly darker Red
    '#B2E0E0', # Light Cyan
    '#99C2FF', # Slightly darker Blue
    '#FFCCE5', # Light Pink
    '#E5FFCC'  # Lighter Green
]

def get_distance(p1, p2):
    """Approximate KM distance using Euclidean (1 unit ~ 111km)"""
    return math.sqrt((p1[0] - p2[0])**2 + (p1[1] - p2[1])**2) * 111.0

def get_bearing(p1, p2):
    """Calculate bearing from p1 to p2"""
    return math.atan2(p2[0] - p1[0], p2[1] - p1[1])

def partition_sizes(n):
    if n < 2: return [n]
    if n == 2: return [2]
    if n == 3: return [3]
    if n == 4: return [2, 2]
    if n % 3 == 0: return [3] * (n // 3)
    if n % 3 == 1: return [3] * ((n - 4) // 3) + [2, 2]
    return [3] * ((n - 2) // 3) + [2]

def natural_sort_key(s):
    """Sort strings with numbers correctly (A1, A2, B1...)"""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def run_automated_logic(warehouse_coords, sites_all):
    if not sites_all: return []
    
    # NEW: Group sites by BILLING FILE to ensure they never mix in a route
    dc_groups = {}
    for s in sites_all:
        dc = s['orig_row'].get('BILLING FILE', 'DEFAULT_DC')
        if dc not in dc_groups: dc_groups[dc] = []
        dc_groups[dc].append(s)
        
    all_final_routes = []
    
    # Process each DC independently
    for dc_name, sites in dc_groups.items():
        for s in sites:
            s['angle'] = get_bearing(warehouse_coords, s['coords'])
            s['dist_to_wh'] = get_distance(warehouse_coords, s['coords'])
            
        sites.sort(key=lambda x: x['angle'])
        
        sizes = partition_sizes(len(sites))
        curr = 0
        for size in sizes:
            chunk = sites[curr : curr + size]
            chunk.sort(key=lambda x: x['dist_to_wh'])
            
            p_sites = chunk.copy()
            while p_sites:
                if len(p_sites) == 1:
                    A = p_sites.pop(0)
                    d = get_distance(warehouse_coords, A['coords'])
                    all_final_routes.append([{"site": A, "dist": max(0, d - 50)}])
                elif len(p_sites) == 2:
                    A, B = p_sites[0], p_sites[1]
                    if get_distance(warehouse_coords, B['coords']) < get_distance(A['coords'], B['coords']):
                        d = get_distance(warehouse_coords, A['coords'])
                        all_final_routes.append([{"site": p_sites.pop(0), "dist": max(0, d - 50)}])
                    else:
                        d_a, d_ab = get_distance(warehouse_coords, A['coords']), get_distance(A['coords'], B['coords'])
                        all_final_routes.append([{"site": p_sites.pop(0), "dist": max(0, d_a - 50)}, {"site": p_sites.pop(0), "dist": d_ab}])
                else: # 3 sites
                    A, B, C = p_sites[0], p_sites[1], p_sites[2]
                    if get_distance(warehouse_coords, B['coords']) < get_distance(A['coords'], B['coords']):
                        d = get_distance(warehouse_coords, A['coords'])
                        all_final_routes.append([{"site": p_sites.pop(0), "dist": max(0, d - 50)}])
                    elif get_distance(warehouse_coords, C['coords']) < get_distance(B['coords'], C['coords']):
                        d_a, d_ab = get_distance(warehouse_coords, A['coords']), get_distance(A['coords'], B['coords'])
                        all_final_routes.append([{"site": p_sites.pop(0), "dist": max(0, d_a - 50)}, {"site": p_sites.pop(0), "dist": d_ab}])
                    else:
                        d_a, d_ab, d_bc = get_distance(warehouse_coords, A['coords']), get_distance(A['coords'], B['coords']), get_distance(B['coords'], C['coords'])
                        all_final_routes.append([{"site": p_sites.pop(0), "dist": max(0, d_a - 50)}, {"site": p_sites.pop(0), "dist": d_ab}, {"site": p_sites.pop(0), "dist": d_bc}])
            curr += size
    return all_final_routes

def plot_map(warehouse_coords, routes, output_file, wh_name):
    m = folium.Map(location=[warehouse_coords[0], warehouse_coords[1]], zoom_start=9)
    folium.Marker(warehouse_coords, popup=f"WAREHOUSE: {wh_name}", icon=folium.Icon(color='black', icon='home', prefix='fa')).add_to(m)
    # Bright colors for maps
    colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue', 'darkblue', 'pink', 'lightgreen', 'black', 'darkgreen', 'darkpurple']
    for r_idx, route in enumerate(routes):
        color = colors[r_idx % len(colors)]
        coords = [warehouse_coords]
        for leg in route:
            site, lat, lon = leg['site'], leg['site']['coords'][0], leg['site']['coords'][1]
            coords.append((lat, lon))
            popup_text = f"Site: {site['id']}<br>DC: {site['orig_row'].get('BILLING FILE','N/A')}<br>Club: {leg.get('clubbing','')}"
            folium.Marker((lat, lon), popup=popup_text, icon=folium.Icon(color=color)).add_to(m)
        if len(coords) > 1:
            folium.PolyLine(coords, color=color, weight=4, opacity=0.8).add_to(m)
    m.save(output_file)

def main():
    # Updated paths for the new project structure
    root_dir = Path(__file__).parent.parent / 'Routing'
    input_dir = root_dir / 'test_cases' / 'Perfect_Format'
    output_dir = root_dir / 'results'
    output_dir.mkdir(parents=True, exist_ok=True)
    
    excel_file = output_dir / "Final_Route_Comparison_Report.xlsx"
    global_summary = []
    
    if not input_dir.exists():
        print(f"Error: Input directory {input_dir} not found.")
        return

    with pd.ExcelWriter(excel_file, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        workbook = writer.book
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'bg_color': '#D3D3D3', 'border': 1})
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'border': 1})
        total_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FDE9D9', 'border': 1})

        for file_path in sorted(input_dir.glob('*.xlsx')):
            print(f"Processing {file_path.name}...")
            df = pd.read_excel(file_path).fillna("")
            df.columns = df.columns.str.strip()
            
            # --- DEDUPLICATE FOR ROUTING ---
            # To handle multiple items at one site (visit level)
            df = df.drop_duplicates(subset=['SITE ID', 'CLUBBING', 'BILLING FILE'])
            print(f"  {file_path.name}: Sites after deduplication: {len(df)}")
            
            # --- COLUMN MAPPING FOR TEST CASES ---
            lat_col, lon_col = 'LATITUDE', 'LONGITUDE'
            wh_col, date_col, id_col = 'Warehouse', 'MIN DATE', 'SITE ID'
            aktbc_old_col, clubbing_old_col = 'AKTBC', 'CLUBBING'
            dc_col, jc_col = 'BILLING FILE', 'JC NAME'
            
            wh_val = str(df.iloc[0][wh_col]).split(' ')[0].upper() if wh_col in df.columns else 'JAIPUR'
            warehouse_coords = WAREHOUSE_MAP.get(wh_val, WAREHOUSE_MAP['DEFAULT'])

            # Automated Logic
            # Automated Logic (Filters Invalid Coords)
            site_list = []
            for _, row in df.iterrows():
                lat, lon = row[lat_col], row[lon_col]
                if pd.isna(lat) or pd.isna(lon) or (lat == 0 and lon == 0):
                    continue
                site_list.append({"id": row[id_col], "coords": (lat, lon), "orig_row": row})
            auto_routes = run_automated_logic(warehouse_coords, site_list)
            
            # Historical Mapping
            df['route_key'] = df[dc_col].astype(str) + "_" + df[clubbing_old_col].astype(str).str.extract(r'([A-Za-z]+)', expand=False).fillna('OTH')
            unique_manual_keys = df['route_key'].unique()
            manual_key_to_idx = {k: i for i, k in enumerate(unique_manual_keys)}
            
            manual_routes = []
            for key in unique_manual_keys:
                route_data = df[df['route_key'] == key].copy()
                route_data['route_num'] = route_data[clubbing_old_col].str.extract(r'(\d+)').astype(float).fillna(1)
                route_data = route_data.sort_values(by='route_num')
                manual_routes.append([{"site": {"id": row[id_col], "coords": (row[lat_col], row[lon_col]), "orig_row": row}, "dist": row[aktbc_old_col], "clubbing": row[clubbing_old_col]} for _, row in route_data.iterrows()])

            # Summaries
            site_to_auto = {}
            for r_idx, route in enumerate(auto_routes):
                label = f"AutoR{r_idx+1}"
                for s_idx, leg in enumerate(route):
                    leg['clubbing'] = f"{label}-S{s_idx+1}"
                    site_to_auto[leg['site']['id']] = {"clubbing": leg['clubbing'], "aktbc": round(leg['dist'], 2), "route_idx": r_idx}

            manual_table_rows = []
            for _, row in df.iterrows():
                auto_data = site_to_auto.get(row[id_col], {"clubbing": "N/A", "aktbc": 0, "route_idx": -1})
                manual_table_rows.append({
                    "site id": row[id_col], "dc no": row[dc_col], "jc name": row.get(jc_col,""), "min date": row.get(date_col,""),
                    "clubbing old": row[clubbing_old_col], "clubbing new": auto_data['clubbing'],
                    "AKTBC old": row[aktbc_old_col], "AKTBC new": auto_data['aktbc'], "wh": wh_val, "lat": row[lat_col], "long": row[lon_col],
                    "_m_key_idx": manual_key_to_idx.get(row['route_key'], -1), "_a_route_idx": auto_data['route_idx']
                })
            manual_table_rows.sort(key=lambda x: (str(x['dc no']), natural_sort_key(x['clubbing old'])))
            auto_table_rows = sorted(manual_table_rows, key=lambda x: natural_sort_key(x['clubbing new']))

            sheet_name = file_path.stem.replace('-', '_')[:31]
            worksheet = workbook.add_worksheet(sheet_name)
            order = ["site id", "dc no", "jc name", "min date", "clubbing old", "clubbing new", "AKTBC old", "AKTBC new", "wh", "lat", "long"]
            
            # --- WRITE SHEET TABLES ---
            worksheet.merge_range(0, 0, 0, len(order)-1, f"HISTORICAL KMS ROUTING - {file_path.stem}", title_format)
            for c, col in enumerate(order): worksheet.write(1, c, col, header_format)
            for r, rd in enumerate(manual_table_rows):
                idx = rd['_m_key_idx']; bg = COLOR_MAP[idx % len(COLOR_MAP)] if idx != -1 else '#FFFFFF'
                fmt = workbook.add_format({'bg_color': bg, 'align': 'center', 'valign': 'vcenter', 'border': 1})
                for c, col in enumerate(order): worksheet.write(r + 2, c, rd[col], fmt)

            start_r_auto = len(manual_table_rows) + 4
            worksheet.merge_range(start_r_auto, 0, start_r_auto, len(order)-1, f"AUTOMATED (DC-ISOLATED) ROUTING - {file_path.stem}", title_format)
            for c, col in enumerate(order): worksheet.write(start_r_auto + 1, c, col, header_format)
            for r, rd in enumerate(auto_table_rows):
                idx = rd['_a_route_idx']; bg = COLOR_MAP[idx % len(COLOR_MAP)] if idx != -1 else '#FFFFFF'
                fmt = workbook.add_format({'bg_color': bg, 'align': 'center', 'valign': 'vcenter', 'border': 1})
                for c, col in enumerate(order): worksheet.write(start_r_auto + r + 2, c, rd[col], fmt)

            total_old = sum(r['AKTBC old'] for r in manual_table_rows)
            total_new = sum(r['AKTBC new'] for r in manual_table_rows)
            diff = round(total_old - total_new, 2)
            
            # --- LOCAL SUMMARY TABLE ---
            sum_r = start_r_auto + len(auto_table_rows) + 3
            summary_headers = ["Metric", "Historical (Old)", "Automated (New)", "Savings"]
            worksheet.merge_range(sum_r, 0, sum_r, 3, f"PERFORMANCE SUMMARY - {file_path.stem}", title_format)
            for c, h in enumerate(summary_headers): worksheet.write(sum_r + 1, c, h, header_format)
            
            metrics = [
                ("Total Routes", len(manual_routes), len(auto_routes), len(manual_routes) - len(auto_routes)),
                ("Route Size: 3", sum(1 for r in manual_routes if len(r) == 3), sum(1 for r in auto_routes if len(r) == 3), ""),
                ("Route Size: 2", sum(1 for r in manual_routes if len(r) == 2), sum(1 for r in auto_routes if len(r) == 2), ""),
                ("Route Size: 1", sum(1 for r in manual_routes if len(r) == 1), sum(1 for r in auto_routes if len(r) == 1), ""),
                ("Avg Route Size", round(len(df)/len(manual_routes), 2) if manual_routes else 0, round(len(df)/len(auto_routes), 2) if auto_routes else 0, ""),
                ("Total Distance (KMS)", round(total_old, 2), round(total_new, 2), f"{diff} KMS")
            ]
            
            for i, (m, old, new, s) in enumerate(metrics):
                row_idx = sum_r + 2 + i
                worksheet.write(row_idx, 0, m, header_format)
                worksheet.write(row_idx, 1, old, total_format)
                worksheet.write(row_idx, 2, new, total_format)
                worksheet.write(row_idx, 3, s, total_format)

            for i in range(len(order)): worksheet.set_column(i, i, 16)
            plot_map(warehouse_coords, manual_routes, output_dir / f"{file_path.stem}_historical.html", wh_val)
            plot_map(warehouse_coords, auto_routes, output_dir / f"{file_path.stem}_automated.html", wh_val)

            global_summary.append({
                "Date": file_path.stem, "Total Sites": len(df),
                "Routes (Old)": len(manual_routes), "Routes (New)": len(auto_routes),
                "Old (S3)": sum(1 for r in manual_routes if len(r) == 3),
                "Old (S2)": sum(1 for r in manual_routes if len(r) == 2),
                "Old (S1)": sum(1 for r in manual_routes if len(r) == 1),
                "New (S3)": sum(1 for r in auto_routes if len(r) == 3),
                "New (S2)": sum(1 for r in auto_routes if len(r) == 2),
                "New (S1)": sum(1 for r in auto_routes if len(r) == 1),
                "Avg Size (Old)": round(len(df)/len(manual_routes), 2) if manual_routes else 0,
                "Avg Size (New)": round(len(df)/len(auto_routes), 2) if auto_routes else 0,
                "KMS (Old)": round(total_old, 2), "KMS (New)": round(total_new, 2), "Savings (KMS)": diff
            })

        # --- EXECUTIVE SUMMARY SHEET ---
        summary_ws = workbook.add_worksheet("EXECUTIVE SUMMARY")
        cols = ["Date", "Total Sites", "Routes (Old)", "Routes (New)", 
                "Old (S3)", "Old (S2)", "Old (S1)", 
                "New (S3)", "New (S2)", "New (S1)", 
                "Avg Size (Old)", "Avg Size (New)", 
                "KMS (Old)", "KMS (New)", "Savings (KMS)"]
        
        summary_ws.merge_range(0, 0, 0, len(cols)-1, "ROUTING PERFORMANCE COMPARISON OVERVIEW", title_format)
        for c, col in enumerate(cols): summary_ws.write(1, c, col, header_format)
        for r, row in enumerate(global_summary):
            for c, col in enumerate(cols): 
                summary_ws.write(r + 2, c, row[col], workbook.add_format({'align': 'center', 'border': 1}))
        
        # Grand Totals
        g_row = len(global_summary) + 2
        summary_ws.write(g_row, 0, "GRAND TOTAL", total_format)
        for c in range(1, len(cols)):
            col_key = cols[c]
            val = sum(r[col_key] for r in global_summary)
            if "Avg" in col_key: val = round(val/len(global_summary), 2)
            summary_ws.write(g_row, c, round(val, 2), total_format)
        for i in range(len(cols)): summary_ws.set_column(i, i, 16)

    print(f"\nSuccess! DC-Isolated Comparison results with Executive Summary saved in: {excel_file}")

if __name__ == "__main__":
    main()

if __name__ == "__main__":
    main()
