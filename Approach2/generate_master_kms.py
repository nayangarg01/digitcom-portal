import pandas as pd
import numpy as np
from pathlib import Path
import math
import folium
import os
import re
import itertools

# --- CONFIG ---
KMS_FILE = Path("Routing/KMS.xlsx")
DPR_FILE = Path("Routing/DPR_common.xlsx")
OUTPUT_FILE = Path("Approach2/Master_KMS_Analysis.xlsx")
MAPS_DIR = Path("Approach2/Maps")
DEEP_DIVES_ROOT = Path("Approach2/DeepDives")

WAREHOUSE_MAP = {
    'JAIPUR': (26.9124, 75.7873),
    'JODHPUR': (26.2389, 73.0243),
    'DEFAULT': (26.9124, 75.7873)
}

HEX_COLORS = ['#FFCCCC', '#CCCCFF', '#CCFFCC', '#E5CCFF', '#FFE5CC', '#FFB2B2', '#CCFFFF']

def haversine(p1, p2):
    R = 6371
    lat1, lon1 = math.radians(p1[0]), math.radians(p1[1])
    lat2, lon2 = math.radians(p2[0]), math.radians(p2[1])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

def optimize_segment(warehouse_coords, cluster):
    best_p = None; min_d = float('inf')
    for p in itertools.permutations(cluster):
        d = haversine(warehouse_coords, p[0]['coords'])
        if len(p) > 1: d += haversine(p[0]['coords'], p[1]['coords'])
        if len(p) > 2: d += haversine(p[1]['coords'], p[2]['coords'])
        if d < min_d: min_d = d; best_p = p
    return best_p

def segment_to_legs(warehouse_coords, segment):
    route_legs = []; cp = warehouse_coords
    for s in segment:
        d = haversine(cp, s['coords'])
        route_legs.append({"site": s, "dist": round(d, 2)})
        cp = s['coords']
    return route_legs

def run_routing(warehouse_coords, cluster):
    # 1. Group by JC
    jc_groups = {}
    for s in cluster:
        row = s.get('row', s.get('orig', {}))
        jc = str(row.get('JC', '')).strip().upper()
        if jc not in jc_groups: jc_groups[jc] = []
        jc_groups[jc].append(s)
    
    # 2. Transition-Aware Metric: Find distance to nearest site in a DIFFERENT JC
    for s in cluster:
        my_jc = str(s.get('row', s.get('orig', {})).get('JC', '')).strip().upper()
        others = [o for o in cluster if str(o.get('row', o.get('orig', {})).get('JC', '')).strip().upper() != my_jc]
        if others:
            s['trans_dist'] = min(haversine(s['coords'], o['coords']) for o in others)
        else:
            s['trans_dist'] = 999999.0 # Effectively isolated
            
    final_routes = []; mixer_pool = []
    
    # Tier 1: Intra-Town Triplets (Prioritize isolated core sites)
    for jc, sites in jc_groups.items():
        # Sort DESC: isolated core sites first, bridges (small trans_dist) last
        unvisited = sorted(sites, key=lambda x: x['trans_dist'], reverse=True)
        while len(unvisited) >= 3:
            seed = unvisited.pop(0) # Pull most isolated core site
            clump = [seed]
            for _ in range(2):
                nearest = min(unvisited, key=lambda s: haversine(seed['coords'], s['coords']))
                clump.append(nearest)
                unvisited.remove(nearest)
            best_p = optimize_segment(warehouse_coords, clump)
            final_routes.append(segment_to_legs(warehouse_coords, best_p))
        mixer_pool.extend(unvisited) # Residues (the 'Bridges') go to Mixer
    
    # Tier 2: Global Mixer (Consolidate the bridges)
    unvisited_mixer = list(mixer_pool)
    while unvisited_mixer:
        # Mixer still uses Outlier-First logic (furthest outlier first)
        seed = max(unvisited_mixer, key=lambda s: haversine(warehouse_coords, s['coords']))
        unvisited_mixer.remove(seed)
        clump = [seed]
        for _ in range(2):
            if not unvisited_mixer: break
            nearest = min(unvisited_mixer, key=lambda s: haversine(seed['coords'], s['coords']))
            clump.append(nearest)
            unvisited_mixer.remove(nearest)
        best_p = optimize_segment(warehouse_coords, clump)
        final_routes.append(segment_to_legs(warehouse_coords, best_p))
        
    return final_routes

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def plot_map(warehouse_coords, routes, output_file, wh_name, title):
    m = folium.Map(location=[warehouse_coords[0], warehouse_coords[1]], zoom_start=9)
    folium.Marker(warehouse_coords, popup=f"WAREHOUSE: {wh_name}", icon=folium.Icon(color='black', icon='home')).add_to(m)
    colors = ['red', 'blue', 'green', 'purple', 'orange', 'darkred', 'cadetblue']
    for r_idx, route in enumerate(routes):
        color = colors[r_idx % len(colors)]
        coords = [warehouse_coords]
        for leg in route:
            s = leg['site']; lat, lon = s['coords']
            coords.append((lat, lon))
            pop = f"Site: {s['id']}<br>Route: {leg.get('club_new','Manual')}"
            folium.Marker((lat, lon), popup=pop, icon=folium.Icon(color=color)).add_to(m)
        if len(coords) > 1:
            folium.PolyLine(coords, color=color, weight=4, opacity=0.8).add_to(m)
    os.makedirs(output_file.parent, exist_ok=True)
    m.save(output_file)

def export_deep_dive(dc_name, date_str, date_df, wh_coords, wh_name, manual_routes, auto_routes):
    folder_name = f"{str(dc_name).replace('/','_')}_{date_str}"
    target_dir = DEEP_DIVES_ROOT / folder_name
    os.makedirs(target_dir, exist_ok=True)
    
    # 1. Styled Data Excel
    id_col, lat_col, lon_col, wh_col = 'eNBsiteID', 'LAT', 'LONG', 'WAREHOUSE'
    data_cols = [id_col, 'BILLING FILE', 'MIN DATE', 'JC', 'CMP', lat_col, lon_col, wh_col]
    
    with pd.ExcelWriter(target_dir / f"{folder_name}_data.xlsx", engine='xlsxwriter') as writer:
        date_df[data_cols].to_excel(writer, index=False, sheet_name='Site Data')
        wb = writer.book; ws = writer.sheets['Site Data']
        head_fmt = wb.add_format({'bold':True,'bg_color':'#2F5597','font_color':'white','border':1,'align':'center'})
        data_fmt = wb.add_format({'border':1,'align':'center'})
        for c, col in enumerate(data_cols):
            ws.write(0, c, col, head_fmt)
            # Apply format to all rows
            for r in range(1, len(date_df) + 1):
                val = date_df.iloc[r-1][col]
                ws.write(r, c, val, data_fmt)
        ws.set_column(0, 0, 25); ws.set_column(1, 7, 15)
    
    # 2. Routing Excel with Colors
    with pd.ExcelWriter(target_dir / f"{folder_name}_routing.xlsx", engine='xlsxwriter') as writer:
        wb = writer.book
        title_fmt = wb.add_format({'bold':True,'align':'center','bg_color':'#2F5597','font_color':'white','border':1})
        for name, routes in [("Manual", manual_routes), ("Automated", auto_routes)]:
            ws = wb.add_worksheet(name); curr = 0
            for r_idx, r in enumerate(routes):
                color = HEX_COLORS[r_idx % len(HEX_COLORS)]
                route_fmt = wb.add_format({'bg_color': color, 'border': 1, 'align': 'center'})
                header_fmt = wb.add_format({'bold': True, 'bg_color': color, 'border': 1, 'align': 'center'})
                ws.merge_range(curr, 0, curr, 5, f"ROUTE {r_idx+1}", title_fmt); curr += 1
                cols = ["SITE ID", "JC", "CMP", "SEQ", "LEG DIST", "CLUB LABEL"]
                for c, h in enumerate(cols): ws.write(curr, c, h, header_fmt)
                curr += 1
                for s_idx, leg in enumerate(r):
                    s = leg['site']
                    # Manual stores in 'orig', Automated in 'row'
                    row_data = s.get('orig', s.get('row', {}))
                    ws.write(curr, 0, s['id'], route_fmt)
                    ws.write(curr, 1, row_data.get('JC', ''), route_fmt)
                    ws.write(curr, 2, row_data.get('CMP', ''), route_fmt)
                    ws.write(curr, 3, s_idx + 1, route_fmt)
                    ws.write(curr, 4, leg['dist'], route_fmt)
                    ws.write(curr, 5, leg.get('club_new', leg.get('club','')), route_fmt)
                    curr += 1
                curr += 2
            ws.set_column(0,0,25); ws.set_column(1,2,15); ws.set_column(3,5,12)
            
    # 3. Maps
    plot_map(wh_coords, manual_routes, target_dir / f"{date_str}_manual.html", wh_name, "Manual")
    plot_map(wh_coords, auto_routes, target_dir / f"{date_str}_automated.html", wh_name, "Automated")

def main():
    print("Loading data...")
    kms_df = pd.read_excel(KMS_FILE).fillna("")
    dpr_df = pd.read_excel(DPR_FILE).fillna("")
    kms_df.columns = kms_df.columns.str.strip()
    dpr_df.columns = dpr_df.columns.str.strip()
    
    id_col, dc_col, date_col, aktbc_col, club_col, wh_col = 'eNBsiteID', 'BILLING FILE', 'MIN DATE', 'AKTBC', 'CLUBBING', 'WAREHOUSE'
    lat_col = 'LAT' if 'LAT' in kms_df.columns else 'LAT '
    lon_col = 'LONG' if 'LONG' in kms_df.columns else 'LONG '
    
    cmp_map = {str(k).strip().upper(): str(v).strip().upper() for k, v in zip(dpr_df['SITE ID'], dpr_df['CMP'])}
    kms_df['CMP'] = kms_df[id_col].apply(lambda x: cmp_map.get(str(x).strip().upper(), "UNKNOWN"))
    
    is_valid = lambda x: str(x).strip().upper() not in ["A", "NR", "B", "OTH", "", "NAN"]
    kms_df = kms_df[kms_df[club_col].apply(is_valid)]
    
    dc_groups = kms_df.groupby(dc_col)
    global_stats = []
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        wb = writer.book
        title_fmt = wb.add_format({'bold':True,'font_size':14,'align':'center','bg_color':'#2F5597','font_color':'white','border':1})
        head_fmt = wb.add_format({'bold':True,'align':'center','bg_color':'#D9E1F2','border':1})
        total_fmt = wb.add_format({'bold':True,'bg_color':'#FCE4D6','border':1,'align':'center'})
        data_fmt = wb.add_format({'align':'center','border':1})
        
        for dc_name, dc_df in sorted(dc_groups, key=lambda x: str(x[0])):
            if not dc_name: continue
            print(f"Processing DC: {dc_name}")
            dc_folder = str(dc_name).replace('/', '_')
            ws = wb.add_worksheet(dc_folder[:31])
            curr_row, dc_old, dc_new = 0, 0, 0
            
            for min_date, date_df in sorted(dc_df.groupby(date_col), key=lambda x: str(x[0])):
                date_str = str(min_date)[:10]
                # Manual routes reconstruction
                hist_groups = {}
                for _, row in date_df.iterrows():
                    c = str(row[club_col]).strip().upper(); key = c[0] if c else "X"
                    if key not in hist_groups: hist_groups[key] = []
                    hist_groups[key].append({"id":row[id_col],"coords":(float(row[lat_col]),float(row[lon_col])),"orig":row})
                
                manual_routes = []
                for k in sorted(hist_groups.keys()):
                    s_list = hist_groups[k]
                    s_list.sort(key=lambda x: natural_sort_key(x['orig'][club_col]))
                    wh_name = str(s_list[0]['orig'][wh_col]).split(' ')[0].upper()
                    wh_coords = WAREHOUSE_MAP.get(wh_name, WAREHOUSE_MAP['DEFAULT'])
                    m_route = []; cp = wh_coords
                    for s in s_list:
                        m_route.append({"site":s,"dist":round(haversine(cp, s['coords']),2),"club":s['orig'][club_col]})
                        cp = s['coords']
                    manual_routes.append(m_route)

                # Automated routes (Refined Outlier-First logic)
                all_auto_legs, auto_for_map = [], []
                for cmp_name, cmp_df in date_df.groupby('CMP'):
                    wh_name = str(cmp_df.iloc[0][wh_col]).split(' ')[0].upper()
                    wh_coords = WAREHOUSE_MAP.get(wh_name, WAREHOUSE_MAP['DEFAULT'])
                    sites = []
                    for _, row in cmp_df.iterrows():
                        try:
                            lat, lon = float(row[lat_col]), float(row[lon_col])
                            if lat != 0: sites.append({"id":row[id_col],"coords":(lat,lon),"row":row})
                        except: continue
                    if not sites: continue
                    routes = run_routing(wh_coords, sites)
                    auto_for_map.extend(routes)
                    for r_idx, r in enumerate(routes):
                        for s_idx, leg in enumerate(r):
                            leg['club_new'] = f"{cmp_name}-R{r_idx+1}-S{s_idx+1}"; leg['r_id'] = f"{cmp_name}-R{r_idx+1}"
                            all_auto_legs.append(leg)
                
                if not all_auto_legs: continue
                
                # Check for Multiple of 3 -> Deep Dive
                if len(date_df) % 3 == 0:
                    export_deep_dive(dc_name, date_str, date_df, wh_coords, wh_name, manual_routes, auto_for_map)

                id_to_auto = {leg['site']['id']: leg for leg in all_auto_legs}
                rows = []
                for _, row in date_df.iterrows():
                    a = id_to_auto.get(row[id_col], {"club_new":"N/A","dist":0,"r_id":"N/A"})
                    rows.append({"SITE ID":row[id_col],"BILLING NO":row[dc_col],"MIN DATE":date_str,"CMP":row['CMP'],
                                 "AKTBC OLD":row[aktbc_col],"AKTBC NEW":a['dist'],"CLUBBING OLD":row[club_col],
                                 "CLUBBING NEW":a['club_new'],"LAT LONG":f"{row[lat_col]},{row[lon_col]}","WAREHOUSE":row[wh_col],"_rid":a['r_id']})
                rows.sort(key=lambda x: natural_sort_key(x['CLUBBING NEW']))
                
                rid_map = {rid: i for i, rid in enumerate(list(dict.fromkeys([r['_rid'] for r in rows])))}
                cols = ["SITE ID","BILLING NO","MIN DATE","CMP","AKTBC OLD","AKTBC NEW","CLUBBING OLD","CLUBBING NEW","LAT LONG","WAREHOUSE"]
                ws.merge_range(curr_row, 0, curr_row, len(cols)-1, f"MIN DATE: {date_str} (Sites: {len(date_df)})", title_fmt)
                curr_row += 1
                for c, h in enumerate(cols): ws.write(curr_row, c, h, head_fmt)
                curr_row += 1
                d_old, d_new = 0, 0
                for r in rows:
                    fmt = wb.add_format({'bg_color': HEX_COLORS[rid_map[r['_rid']] % len(HEX_COLORS)], 'border':1, 'align':'center'}) if r['_rid'] != "N/A" else data_fmt
                    for c, col in enumerate(cols): ws.write(curr_row, c, r[col], fmt)
                    d_old += r['AKTBC OLD']; d_new += r['AKTBC NEW']; curr_row += 1
                
                ws.merge_range(curr_row, 0, curr_row, 3, "TOTAL", total_fmt)
                ws.write(curr_row, 4, round(d_old,2), total_fmt); ws.write(curr_row, 5, round(d_new,2), total_fmt)
                ws.merge_range(curr_row, 6, curr_row, 9, f"SAVINGS: {round(d_old-d_new,2)} KMS", total_fmt)
                curr_row += 3; dc_old += d_old; dc_new += d_new
                
                # Global Map (Always generated)
                wh_name = str(date_df.iloc[0][wh_col]).split(' ')[0].upper()
                wh_coords = WAREHOUSE_MAP.get(wh_name, WAREHOUSE_MAP['DEFAULT'])
                plot_map(wh_coords, manual_routes, MAPS_DIR/dc_folder/f"{date_str}_historical.html", wh_name, "Historical")
                plot_map(wh_coords, auto_for_map, MAPS_DIR/dc_folder/f"{date_str}_automated.html", wh_name, "Automated")
            
            global_stats.append({"DC":dc_name,"Old":dc_old,"New":dc_new})
            for i in range(len(cols)): ws.set_column(i, i, 15)

        sum_ws = wb.add_worksheet("EXECUTIVE SUMMARY")
        sum_cols = ["DC Number", "Old KMS", "New KMS", "Savings", "%"]
        for c, h in enumerate(sum_cols): sum_ws.write(0, c, h, head_fmt)
        for i, s in enumerate(global_stats):
            r = i+1; sav = s['Old']-s['New']; p = (sav/s['Old']*100) if s['Old']>0 else 0
            sum_ws.write(r,0,s['DC'],data_fmt); sum_ws.write(r,1,round(s['Old'],2),data_fmt); sum_ws.write(r,2,round(s['New'],2),data_fmt)
            sum_ws.write(r,3,round(sav,2),data_fmt); sum_ws.write(r,4,f"{round(p,1)}%",data_fmt)
        for i in range(len(sum_cols)): sum_ws.set_column(i, i, 18)

    print(f"Report: {OUTPUT_FILE}\nMaps: {MAPS_DIR}\nDeep Dives: {DEEP_DIVES_ROOT}")

if __name__ == "__main__":
    main()
