import sys
import os
import pickle
import json
import argparse
from datetime import datetime

# Add BillingEngine_OOP to python search path
script_dir = os.path.dirname(os.path.abspath(__file__))
oop_dir = os.path.abspath(os.path.join(script_dir, "..", "..", "BillingEngine_OOP"))
sys.path.append(oop_dir)

try:
    from site_models import A6Site, A6B6Site
except ImportError as e:
    print(json.dumps({"error": f"Failed to import site models: {e}"}))
    sys.exit(1)

def load_db():
    db_path = os.path.join(oop_dir, "sites_db.pkl")
    if not os.path.exists(db_path):
        return {}
    try:
        with open(db_path, "rb") as f:
            return pickle.load(f)
    except Exception as e:
        sys.stderr.write(f"Error loading pickle: {e}\n")
        return {}

def format_date(val):
    if not val:
        return None
    try:
        return pd.to_datetime(val).strftime('%Y-%m-%d')
    except:
        return str(val)

def clean_val(val):
    if val is None:
        return None
    import math
    import pandas as pd
    # Check float NaN
    if isinstance(val, float) and math.isnan(val):
        return None
    # Check pandas NaT/NaN
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    if val_str.upper() in ["NAN", "NAT", "NONE", "N/A"]:
        return None
    return val

def site_to_dict(key, site, full=False):
    # Base attributes
    d = {
        "unique_key": key,
        "site_id": clean_val(getattr(site, "site_id", "")),
        "pmp_id": clean_val(getattr(site, "pmp_id", "")),
        "sector_id": clean_val(getattr(site, "sector_id", "")),
        "hop_id": clean_val(getattr(site, "hop_id", "")),
        "latitude": clean_val(getattr(site, "latitude", None)),
        "longitude": clean_val(getattr(site, "longitude", None)),
        "tower_type": clean_val(getattr(site, "tower_type", "")),
        "jc": clean_val(getattr(site, "jc", "")),
        "wh": clean_val(getattr(site, "wh", "")),
        "vehicle_no": clean_val(getattr(site, "vehicle_no", "")),
        "km_actual": clean_val(getattr(site, "km_actual", 0.0)),
        "km_wo": clean_val(getattr(site, "km_wo", 0.0)),
        "km_threshold": clean_val(getattr(site, "km_threshold", 0.0)),
        "wo": clean_val(getattr(site, "wo", "")),
        "dc_no": clean_val(getattr(site, "dc_no", "")),
        "performa_no": clean_val(getattr(site, "performa_no", "")),
        "wbs_id": clean_val(getattr(site, "wbs_id", "")),
        "po_no": clean_val(getattr(site, "po_no", "")),
        "activity_type": clean_val(getattr(site, "activity_type", "")),
        "min_no": clean_val(getattr(site, "min_no", "")),
        "min_date": clean_val(str(site.min_date)) if getattr(site, "min_date", None) else None,
        "completion_date": clean_val(str(site.completion_date)) if getattr(site, "completion_date", None) else None,
        "remarks": clean_val(getattr(site, "remarks", "")),
        "no_of_sectors": clean_val(getattr(site, "no_of_sectors", 0.0)),
        "clubbing": clean_val(getattr(site, "clubbing", ""))
    }
    
    if full:
        # Include items and dispatches
        d["items"] = {k: clean_val(v) for k, v in getattr(site, "items", {}).items()}
        d["dispatches"] = []
        if hasattr(site, "dispatches"):
            for disp in site.dispatches:
                d["dispatches"].append({
                    "sap_code": clean_val(disp.get("sap_code", "")),
                    "description": clean_val(disp.get("description", "")),
                    "quantity": clean_val(disp.get("quantity", 0.0)),
                    "min_number": clean_val(disp.get("min_number", "")),
                    "date": clean_val(str(disp.get("date"))) if disp.get("date") else None,
                    "remarks": clean_val(disp.get("remarks", "")),
                    "pmp_id": clean_val(disp.get("pmp_id", "")),
                    "activity": clean_val(disp.get("activity", ""))
                })
    return d

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--summary", action="store_true", help="Get summary of database status")
    parser.add_argument("--list", action="store_true", help="List all site summaries")
    parser.add_argument("--detail", type=str, default=None, help="Get detailed view of a site by unique key")
    args = parser.parse_args()

    sites = load_db()

    if args.summary:
        db_path = os.path.join(oop_dir, "sites_db.pkl")
        last_modified = None
        if os.path.exists(db_path):
            mtime = os.path.getmtime(db_path)
            last_modified = datetime.fromtimestamp(mtime).isoformat()
            
        a6_count = sum(1 for s in sites.values() if getattr(s, "activity_type", "") == "A6")
        a6b6_count = sum(1 for s in sites.values() if getattr(s, "activity_type", "") == "A6+B6")
        
        summary = {
            "success": True,
            "total_sites": len(sites),
            "a6_sites": a6_count,
            "a6b6_sites": a6b6_count,
            "last_updated": last_modified
        }
        print(json.dumps(summary))
        return

    if args.list:
        summary_list = []
        for key, s in sites.items():
            summary_list.append(site_to_dict(key, s, full=False))
        print(json.dumps(summary_list))
        return

    if args.detail:
        site_key = args.detail
        if site_key in sites:
            print(json.dumps(site_to_dict(site_key, sites[site_key], full=True)))
        else:
            # Fallback search by site_id if key matches site_id
            matched_key = None
            for key, s in sites.items():
                if getattr(s, "site_id", "").strip().upper() == site_key.strip().upper():
                    matched_key = key
                    break
            if matched_key:
                print(json.dumps(site_to_dict(matched_key, sites[matched_key], full=True)))
            else:
                print(json.dumps({"error": f"Site with key/ID '{site_key}' not found"}))
        return

    # Default: print instructions
    print(json.dumps({"error": "No action specified. Use --summary, --list, or --detail"}))

if __name__ == "__main__":
    main()
