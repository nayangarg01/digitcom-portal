import json
import os
from data_loader import DataFactory

def export_sites_to_json(output_path):
    MASTER_PATH = "MASTER_TRACKER_DATA.xlsx"
    
    # Initialize and sync
    factory = DataFactory(MASTER_PATH)
    factory.sync_from_master()
    
    # Convert all site objects to serializable dictionaries
    export_data = []
    for site_id, obj in factory.sites.items():
        site_dict = {
            "site_id": obj.site_id,
            "pmp_id": obj.pmp_id,
            "activity_type": obj.activity_type,
            "wo": obj.wo,
            "dc_no": str(obj.dc_no),
            "jc": obj.jc,
            "wh": obj.wh,
            "tower_type": obj.tower_type,
            "km_actual": obj.km_actual,
            "completion_date": str(obj.completion_date) if obj.completion_date else "N/A"
        }
        export_data.append(site_dict)
    
    # Save to JSON
    with open(output_path, 'w') as f:
        json.dump(export_data, f, indent=4)
        
    print(f"DEBUG: Successfully exported {len(export_data)} sites to {output_path}")

if __name__ == "__main__":
    export_sites_to_json("SiteViewer_UI/sites_data.json")
