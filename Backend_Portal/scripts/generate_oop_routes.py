import sys
import os
import argparse
import pandas as pd
import json

# Ensure we can import from unified_routing_engine
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(script_dir)

# Add BillingEngine_OOP to python search path
oop_dir = os.path.abspath(os.path.join(script_dir, "..", "..", "BillingEngine_OOP"))
sys.path.append(oop_dir)

try:
    from data_loader import DataFactory
    import unified_routing_engine
except ImportError as e:
    print(json.dumps({"error": f"Failed to import required scripts. Error: {e}"}))
    sys.exit(1)

def format_date_str(val):
    if pd.isna(val) or not val:
        return ""
    try:
        dt = pd.to_datetime(val)
        return dt.strftime('%-d-%b-%y')
    except:
        return str(val)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dc_numbers", default=None, help="Comma-separated list of DC Numbers")
    parser.add_argument("--dates", default=None, help="Comma-separated list of MIN/completion dates")
    parser.add_argument("--api_key", required=True, help="Google Maps API Key")
    parser.add_argument("--output", required=True, help="Output optimized Excel path")
    args = parser.parse_args()

    api_key = args.api_key
    output_path = args.output

    # Load DataFactory
    try:
        factory = DataFactory(None)
    except Exception as e:
        print(json.dumps({"error": f"Failed to load database: {str(e)}"}))
        sys.exit(1)

    # Filter sites
    filtered_sites = []
    
    if args.dc_numbers:
        target_dcs = [dc.strip().upper() for dc in args.dc_numbers.split(',') if dc.strip()]
        filtered_sites = [s for s in factory.sites.values() if str(s.dc_no).strip().upper() in target_dcs]
        sys.stderr.write(f"LOG: Filtering routes by DC Numbers: {target_dcs} (Found {len(filtered_sites)} sites)\n")
    elif args.dates:
        target_dates = [format_date_str(d.strip()) for d in args.dates.split(',') if d.strip()]
        # Filter by min_date or completion_date
        filtered_sites = []
        for s in factory.sites.values():
            s_min_date = format_date_str(s.min_date) if s.min_date else ""
            s_comp_date = format_date_str(s.completion_date) if s.completion_date else ""
            if s_min_date in target_dates or s_comp_date in target_dates:
                filtered_sites.append(s)
        sys.stderr.write(f"LOG: Filtering routes by dates: {target_dates} (Found {len(filtered_sites)} sites)\n")
    else:
        # If neither is specified, default to routing all sites in the database that have coordinates
        filtered_sites = [s for s in factory.sites.values()]
        sys.stderr.write(f"LOG: No filters provided, routing all sites in database (Total {len(filtered_sites)} sites)\n")

    # Clean out sites without coordinates
    valid_sites = []
    for s in filtered_sites:
        try:
            lat = float(s.latitude) if s.latitude is not None else None
            lng = float(s.longitude) if s.longitude is not None else None
            if lat is not None and lng is not None and not pd.isna(lat) and not pd.isna(lng):
                valid_sites.append(s)
        except:
            pass

    if not valid_sites:
        print(json.dumps({"error": "No sites with valid Lat/Long coordinates found for the given criteria."}))
        sys.exit(1)

    sys.stderr.write(f"LOG: Found {len(valid_sites)} valid sites for routing.\n")

    # Build routing DataFrame
    rows = []
    for s in valid_sites:
        # Handle clubbing/mrn mapping
        club_val = str(s.clubbing).strip() if s.clubbing else ""
        if club_val.upper() in ["NAN", "NONE", "N/A"]:
            club_val = ""
            
        mrn_val = "NO"
        if club_val.upper() == "MRN":
            mrn_val = "YES"
            club_val = ""

        # Ensure warehouse name is clean
        wh_name = str(s.wh).strip()
        if not wh_name or wh_name.upper() in ["NAN", "NONE", "N/A"]:
            wh_name = "Jaipur - JLKD" # Default fallback
            
        rows.append({
            'ENB SITE ID': s.site_id,
            'JC': s.jc if s.jc else 'N/A',
            'CMP': 'DIGITCOM',
            'WH': wh_name,
            'MIN DATE': format_date_str(s.min_date) if s.min_date else format_date_str(s.completion_date),
            'LAT ': float(s.latitude),
            'LONG': float(s.longitude),
            'MRN REQD OR NOT': mrn_val,
            'KM CAP': float(s.km_threshold) if s.km_threshold else (100.0 if s.activity_type == 'A6+B6' else 50.0),
            'CLUBBING': club_val,
            'KM FROM WH TO SITE': 0.0,
            'KM FROM SITE TO WH': 0.0,
            'CHARGEBLE KMS': 0.0
        })

    df = pd.DataFrame(rows)

    # Save to a temporary input file
    temp_dir = os.path.join(script_dir, "..", "uploads")
    os.makedirs(temp_dir, exist_ok=True)
    temp_input_path = os.path.join(temp_dir, f"temp_routing_input_{os.getpid()}.xlsx")
    df.to_excel(temp_input_path, index=False)

    sys.stderr.write(f"LOG: Written temporary routing input sheet containing {len(df)} rows.\n")

    try:
        # Execute the unified routing engine's logic
        res = unified_routing_engine.process_billing(temp_input_path, api_key, output_path)
        
        # Clean up temp file
        if os.path.exists(temp_input_path):
            os.unlink(temp_input_path)
            
        if "output" in res:
            res["filename"] = os.path.basename(res["output"])
            
        print(json.dumps(res))
    except Exception as e:
        if os.path.exists(temp_input_path):
            os.unlink(temp_input_path)
        print(json.dumps({"error": f"Routing optimizer failed: {str(e)}"}))
        sys.exit(1)

if __name__ == "__main__":
    main()
