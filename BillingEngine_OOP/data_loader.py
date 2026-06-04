import pandas as pd
import os
from site_models import A6Site, A6B6Site

TEMPLATE_ITEMS = [
  {"sap": "3367489", "desc": "CHRG BASE RADIO INST&COMS AZIMUTH", "uom": "EA", "rate": 200},
  {"sap": "3367548", "desc": "CHRG LAYING-MULTIMODE FIBER CABLE", "uom": "M", "rate": 5},
  {"sap": "3137158", "desc": "LAYING-2CX2.5SQMM POWER CABLE", "uom": "M", "rate": 12},
  {"sap": "3397253", "desc": "CHRG EXTRA TRANSPORT. > 50 KM (PICKUP", "uom": "KM", "rate": 20},
  {"sap": "3397271", "desc": "CHRG TRANSPORT-UPTO 50KM-SCV-2-3SITE", "uom": "EA", "rate": 1500},
  {"sap": "3367713", "desc": "CHRG ATP-11C - WIFI A6 DEPLOYMENT", "uom": "EA", "rate": 700},
  {"sap": "3367739", "desc": "CHRG ATP-11A - WIFI A6 DEPLOYMENT", "uom": "EA", "rate": 700},
  {"sap": "3317347", "desc": "SITC 1CX6 SQMM CU CBL YY FRLSH", "uom": "M", "rate": 68},
  {"sap": "3383067", "desc": "CHRG APPLY PUFF SEALENT AT CABLE ENTRY", "uom": "EA", "rate": 330},
  {"sap": "3269867", "desc": "TERMINATION OF CABLE 1CX6 SQMM Y", "uom": "M", "rate": 40},
  {"sap": "3397248", "desc": "EXTRA VISIT", "uom": "EA", "rate": 1000},
  {"sap": "3268025", "desc": "INSTALLATION OF POLE MOUNT ON TOWER", "uom": "EA", "rate": 500}
]

TEMPLATE_ITEMS_A6_B6 = [
  {"sap": "3398758", "desc": "ITC A6 & B6 - ONE-SECTOR SITE", "uom": "EA", "rate": 13015},
  {"sap": "3398834", "desc": "ITC A6 & B6 - TWO-SECTOR SITE", "uom": "EA", "rate": 14915},
  {"sap": "3398764", "desc": "ITC A6 & B6 - THREE-SECTOR SITE", "uom": "EA", "rate": 16815},
  {"sap": "3339581", "desc": "CHRG TRANSPORTATION-BEYOND 100 KM-SCV", "uom": "KM", "rate": 20}
]

def safe_float(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().upper() in ["NR", "NA", "N.A", "-"]:
        return 0.0
    try:
        return float(val)
    except:
        return 0.0

class DataFactory:
    """
    Factory class to maintain a Master Dictionary of all sites based EXCLUSIVELY on the Master DPR.
    """
    def __init__(self, master_tracker_path, db_path=None):
        self.master_tracker_path = master_tracker_path
        if db_path is None:
            db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sites_db.pkl")
        self.db_path = db_path
        self.sites = {} # Master Dictionary: {unique_key: SiteObject}
        self.load_database()

    def load_database(self):
        if os.path.exists(self.db_path):
            try:
                import pickle
                with open(self.db_path, "rb") as f:
                    self.sites = pickle.load(f)
                print(f"DEBUG: Loaded {len(self.sites)} sites from local database {self.db_path}")
            except Exception as e:
                print(f"WARNING: Could not load local database: {e}")
                self.sites = {}

    def save_database(self):
        try:
            import pickle
            with open(self.db_path, "wb") as f:
                pickle.dump(self.sites, f)
            print(f"DEBUG: Saved {len(self.sites)} sites to local database {self.db_path}")
        except Exception as e:
            print(f"WARNING: Could not save local database: {e}")

    def sync_from_master(self):
        """
        Ingests every single site from the Master Tracker.
        This is our main database.
        Detects sheet layouts dynamically and parses A6 and A6+B6 sheets.
        """
        print(f"DEBUG: Syncing all sites from Master Tracker: {os.path.basename(self.master_tracker_path)}")
        try:
            xl = pd.ExcelFile(self.master_tracker_path)
            
            for sheet_name in xl.sheet_names:
                df_first2 = xl.parse(sheet_name, header=None, nrows=2)
                if df_first2.empty or len(df_first2) < 1:
                    continue
                    
                # Check for eNBsiteID in row 0 or row 1
                row0 = [str(val).strip().upper() for val in df_first2.iloc[0].tolist() if not pd.isna(val)]
                
                is_a6b6 = False
                is_a6 = False
                
                if any('ENBSITEID' in val for val in row0):
                    is_a6b6 = True
                elif len(df_first2) > 1:
                    row1 = [str(val).strip().upper() for val in df_first2.iloc[1].tolist() if not pd.isna(val)]
                    if any('ENBSITEID' in val for val in row1):
                        is_a6 = True
                        
                if not is_a6b6 and not is_a6:
                    # Not a master billing sheet
                    continue
                    
                print(f"DEBUG: Detected sheet '{sheet_name}' layout: {'A6+B6' if is_a6b6 else 'A6'}")
                df_full = xl.parse(sheet_name, header=None)
                
                if is_a6:
                    # --- PROCESS A6 SHEET LAYOUT ---
                    # Build header and code mappings
                    code_to_col_idx = {}
                    row0_full = df_full.iloc[0].tolist()
                    row1_full = df_full.iloc[1].tolist()
                    for i in range(len(row0_full)):
                        sap = str(row0_full[i]).split('.')[0].strip()
                        if sap and sap != 'nan':
                            code_to_col_idx[sap] = i
                        desc = str(row1_full[i]).strip()
                        if desc and desc != 'nan':
                            code_to_col_idx[desc] = i
                    
                    headers = [str(h).strip() for h in row1_full]
                    header_to_idx = {h: i for i, h in enumerate(headers)}
                    
                    master_map_a6 = {
                        'eNBsiteID': 'site_id',
                        'PMP ID': 'pmp_id',
                        'GIS SECTOR_ID': 'sector_id',
                        'FB-FT HOP ID': 'hop_id',
                        'Tower type ': 'tower_type',
                        'JC': 'jc',
                        'WH': 'wh',
                        'VEHICLE NO': 'vehicle_no',
                        'KM FROM WH TO SITE': 'km_actual',
                        'KM IN WO': 'km_wo',
                        'KM-50(for a6+b6-100)': 'km_threshold',
                        'WO': 'wo',
                        'BILLING FILE': 'dc_no',
                        'PERFORMA INVOICE NO': 'performa_no',
                        'MIN NO': 'min_no',
                        'MIN DATE': 'min_date',
                        'Completion Date ': 'completion_date',
                        'LAT ': 'latitude',
                        'LONG': 'longitude',
                        'REMARKS': 'remarks',
                        'NO OF SECTOR': 'no_of_sectors',
                        'CLUBBING': 'clubbing'
                    }
                    
                    for r_idx in range(2, len(df_full)):
                        row = df_full.iloc[r_idx].tolist()
                        site_id_idx = header_to_idx.get('eNBsiteID')
                        if site_id_idx is None or site_id_idx >= len(row): continue
                        site_id = str(row[site_id_idx]).strip()
                        if not site_id or site_id.lower() in ["nan", "none", "site id"]: continue
                        
                        pmp_idx = header_to_idx.get('PMP ID')
                        pmp_id = str(row[pmp_idx]).strip() if (pmp_idx is not None and pmp_idx < len(row)) else "N/A"
                        
                        min_no_idx = header_to_idx.get('MIN NO')
                        min_no = str(row[min_no_idx]).strip().split('.')[0] if (min_no_idx is not None and min_no_idx < len(row)) else "N/A"
                        
                        unique_key = f"{site_id}_{pmp_id}_{min_no}_A6"
                        
                        if unique_key in self.sites:
                            site_obj = self.sites[unique_key]
                        else:
                            site_obj = A6Site(site_id)
                            self.sites[unique_key] = site_obj
                        
                        for excel_col, attr_name in master_map_a6.items():
                            col_i = header_to_idx.get(excel_col.strip())
                            if col_i is not None and col_i < len(row):
                                val = row[col_i]
                                if attr_name in ['km_actual', 'km_wo', 'km_threshold', 'no_of_sectors']:
                                    val = safe_float(val)
                                elif attr_name in ['latitude', 'longitude']:
                                    try: val = float(val) if not pd.isna(val) else None
                                    except: val = None
                                elif attr_name in ['min_date', 'completion_date']:
                                    if pd.isna(val) or str(val).strip() == "":
                                        val = None
                                setattr(site_obj, attr_name, val)
                        
                        for item in TEMPLATE_ITEMS:
                            sap_code = item['sap']
                            if sap_code == "3397248":
                                col_idx = header_to_idx.get('EXTRA VISIT IN WO') if 'EXTRA VISIT IN WO' in header_to_idx else code_to_col_idx.get(sap_code)
                            elif sap_code == "3268025":
                                col_idx = header_to_idx.get('Polemount in wo') if 'Polemount in wo' in header_to_idx else code_to_col_idx.get(sap_code)
                            else:
                                col_idx = code_to_col_idx.get(sap_code)
                            
                            val = safe_float(row[col_idx]) if col_idx is not None and col_idx < len(row) else 0.0
                            site_obj.add_item(sap_code, val)
                            
                elif is_a6b6:
                    # --- PROCESS A6+B6 SHEET LAYOUT ---
                    headers = [str(h).strip() for h in df_full.iloc[0].tolist()]
                    header_to_idx = {h: i for i, h in enumerate(headers)}
                    
                    master_map_a6b6 = {
                        'eNBsiteID': 'site_id',
                        'PMP ID': 'pmp_id',
                        'SEC ID': 'sector_id',
                        'FB-FT HOP ID': 'hop_id',
                        'TOWER': 'tower_type',
                        'JC': 'jc',
                        'WAREHOUSE': 'wh',
                        'VEHICLE NO': 'vehicle_no',
                        'AKTBC(FT)': 'km_actual',
                        'KM IN WO': 'km_wo',
                        'KM-100': 'km_threshold',
                        'WO': 'wo',
                        'BILLING FILE': 'dc_no',
                        'PERFORMA INVOICE NO': 'performa_no',
                        'MIN NO': 'min_no',
                        'MIN DATE': 'min_date',
                        'RFS DATE': 'completion_date',
                        'LAT': 'latitude',
                        'LONG': 'longitude',
                        'REMARKS': 'remarks',
                        'NO OF SECTOR': 'no_of_sectors',
                        'CLUBBING': 'clubbing'
                    }
                    
                    for r_idx in range(1, len(df_full)):
                        row = df_full.iloc[r_idx].tolist()
                        site_id_idx = header_to_idx.get('eNBsiteID')
                        if site_id_idx is None or site_id_idx >= len(row): continue
                        site_id = str(row[site_id_idx]).strip()
                        if not site_id or site_id.lower() in ["nan", "none", "site id"]: continue
                        
                        pmp_idx = header_to_idx.get('PMP ID')
                        pmp_id = str(row[pmp_idx]).strip() if (pmp_idx is not None and pmp_idx < len(row)) else "N/A"
                        
                        min_no_idx = header_to_idx.get('MIN NO')
                        min_no = str(row[min_no_idx]).strip().split('.')[0] if (min_no_idx is not None and min_no_idx < len(row)) else "N/A"
                        
                        unique_key = f"{site_id}_{pmp_id}_{min_no}_A6B6"
                        
                        if unique_key in self.sites:
                            site_obj = self.sites[unique_key]
                        else:
                            site_obj = A6B6Site(site_id)
                            self.sites[unique_key] = site_obj
                        
                        for excel_col, attr_name in master_map_a6b6.items():
                            col_i = header_to_idx.get(excel_col.strip())
                            if col_i is not None and col_i < len(row):
                                val = row[col_i]
                                if attr_name in ['km_actual', 'km_wo', 'km_threshold', 'no_of_sectors']:
                                    val = safe_float(val)
                                elif attr_name in ['latitude', 'longitude']:
                                    try: val = float(val) if not pd.isna(val) else None
                                    except: val = None
                                elif attr_name in ['min_date', 'completion_date']:
                                    if pd.isna(val) or str(val).strip() == "":
                                        val = None
                                setattr(site_obj, attr_name, val)
                        
                        for item in TEMPLATE_ITEMS_A6_B6:
                            sap_code = item['sap']
                            if sap_code == "3339581":
                                val = safe_float(row[header_to_idx['AKTBC(FT)']]) if 'AKTBC(FT)' in header_to_idx and header_to_idx['AKTBC(FT)'] < len(row) else 0.0
                            else:
                                found_col_idx = None
                                for key, idx in header_to_idx.items():
                                    if sap_code in key:
                                        found_col_idx = idx
                                        break
                                val = safe_float(row[found_col_idx]) if found_col_idx is not None and found_col_idx < len(row) else 0.0
                            site_obj.add_item(sap_code, val)
                            
            self.save_database()
            print(f"DEBUG: Master Dictionary successfully synchronized with {len(self.sites)} sites.")
        except Exception as e:
            print(f"ERROR during Master Sync: {e}")
            import traceback
            traceback.print_exc()
            import traceback
            traceback.print_exc()

    def get_site(self, site_id):
        """Retrieve a specific site object by ID."""
        return self.sites.get(str(site_id).strip())

    def sync_from_mindump(self, mindump_path):
        """
        Loads dispatches from MIN Dump Excel and associates them with Site objects.
        """
        print(f"DEBUG: Syncing MIN Dump: {os.path.basename(mindump_path)}")
        try:
            xl = pd.ExcelFile(mindump_path)
            
            # Reset dispatches to prevent duplicate appends on re-sync
            for site_obj in self.sites.values():
                site_obj.dispatches = []
                
            # Robust sheet selection logic:
            sheets_to_process = []
            a6_options = ["A6 DUMP", "A6", "a6"]
            a6_sheet = next((s for s in a6_options if s in xl.sheet_names), xl.sheet_names[0])
            sheets_to_process.append((a6_sheet, 'A6'))
            
            b6_options = ["B6 DUMP", "B6", "b6"]
            b6_sheet = next((s for s in b6_options if s in xl.sheet_names), xl.sheet_names[0])
            sheets_to_process.append((b6_sheet, 'B6'))
                
            # Precalculate A6 PMP mappings
            pmp_to_site = []
            for site_id, site_obj in self.sites.items():
                if site_obj.pmp_id and site_obj.pmp_id != 'N/A':
                    pmp_to_site.append((site_obj, site_obj.pmp_id.upper().strip()))
            
            # Precalculate B6 Hop ID mappings
            site_hop_ends = []
            for site_id, site_obj in self.sites.items():
                hop_id = getattr(site_obj, 'hop_id', '')
                if hop_id and hop_id != 'N/A' and str(hop_id).upper() != 'NONE':
                    clean_hop = str(hop_id).replace('_A6', '')
                    if '-I-RJ-' in clean_hop:
                        parts = clean_hop.split('-I-RJ-')
                        ends = [parts[0], 'I-RJ-' + parts[1]]
                    else:
                        ends = [clean_hop]
                    site_hop_ends.append((site_obj, [e.upper().strip() for e in ends if e]))

            total_dispatches = 0
            for sheet_name, sub_activity in sheets_to_process:
                print(f"DEBUG: Processing MIN Dump sheet: {sheet_name} as sub_activity: {sub_activity}")
                df_dump = xl.parse(sheet_name)
                
                # Check for standard columns
                required_cols = ['SAP Code', 'No. Of Qty']
                missing = [c for c in required_cols if c not in df_dump.columns]
                if missing:
                    print(f"WARNING: Sheet {sheet_name} is missing columns {missing}, skipping.")
                    continue
                
                for _, row in df_dump.iterrows():
                    sap_code = str(row.get('SAP Code', '')).strip().split('.')[0]
                    if not sap_code or sap_code.lower() in ['nan', 'none']: continue
                    
                    qty = row.get('No. Of Qty', 0.0)
                    try: qty = float(qty) if not pd.isna(qty) else 0.0
                    except: qty = 0.0
                    
                    desc = str(row.get('Material Description', '')).strip()
                    min_no = str(row.get('MIN Number', '')).strip()
                    min_date = row.get('Date')
                    if pd.isna(min_date): min_date = None
                    remarks = str(row.get('Remarks-MIN', '')).strip()
                    
                    matched_sites = []
                    pmp_id_val = None
                    
                    if sub_activity == 'B6':
                        enb_id = str(row.get('ENB ID', '')).upper().strip()
                        site_id_val = str(row.get('Site ID', '')).upper().strip()
                        dwg = str(row.get('DWG', '')).upper().strip()
                        
                        pmp_id_val = str(row.get('COMMON ID', '')).strip()
                        if not pmp_id_val or pmp_id_val.lower() in ['nan', 'none']:
                            pmp_id_val = str(row.get('Site ID', '')).strip()
                            
                        for site_obj, ends in site_hop_ends:
                            for end in ends:
                                if (end in enb_id) or (end in site_id_val) or (end in dwg):
                                    if site_obj not in matched_sites:
                                        matched_sites.append(site_obj)
                                    break
                    else:
                        wbs_id = str(row.get('WBS ID', '')).upper().strip()
                        site_id_val = str(row.get('Site ID', '')).upper().strip()
                        
                        for site_obj, p_id in pmp_to_site:
                            if (p_id in wbs_id) or (p_id in site_id_val):
                                if site_obj not in matched_sites:
                                    matched_sites.append(site_obj)
                    
                    for site_obj in matched_sites:
                        p_val = site_obj.pmp_id if sub_activity == 'A6' else pmp_id_val
                        site_obj.add_dispatch(sap_code, desc, qty, min_no, min_date, remarks, pmp_id=p_val, activity=sub_activity)
                        if sub_activity == 'A6' and wbs_id and wbs_id.upper() != 'NAN':
                            site_obj.wbs_id = wbs_id
                        total_dispatches += 1
                        
            self.save_database()
            print(f"DEBUG: Successfully synchronized {total_dispatches} material dispatches across sites.")
            
        except Exception as e:
            print(f"ERROR during MIN Dump Sync: {e}")

if __name__ == "__main__":
    MASTER_PATH = "MASTER_TRACKER_DATA.xlsx"
    MINDUMP_PATH = "MIN_DUMP_DATA.xlsx"
    
    factory = DataFactory(MASTER_PATH)
    factory.sync_from_master()
    factory.sync_from_mindump(MINDUMP_PATH)
    
    print(f"\nFinal Master Dictionary Size: {len(factory.sites)} sites.")
    
    # Quick Check: Look at one site with dispatches
    sites_with_dispatches = [s for s in factory.sites.values() if s.dispatches]
    print(f"Sites with dispatches: {len(sites_with_dispatches)}")
    if sites_with_dispatches:
        print(f"Sample Site with dispatch: {sites_with_dispatches[0]}")
        print(f"First dispatch: {sites_with_dispatches[0].dispatches[0]}")
