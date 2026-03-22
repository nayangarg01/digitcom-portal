import sys
import json
import os
import subprocess
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

def process_group(entry, test_data_dir, api_key, python_path):
    group_id = entry['group_id']
    file_path = os.path.join(test_data_dir, entry['filename'])
    lat, lng = entry['lat'], entry['lng']
    temp_output = os.path.join(test_data_dir, f"temp_{group_id}.xlsx")
    
    try:
        # Use the provided python_path to ensure environment consistency
        cmd = [
            python_path, "scripts/route_optimizer.py",
            file_path, str(lat), str(lng), api_key, temp_output
        ]
        
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        stdout, stderr = process.communicate()
        
        # Check if the Excel file was actually created
        if not os.path.exists(temp_output):
            return {**entry, "status": "Error", "error": f"Excel not created. Stderr: {stderr[:100]}"}

        # Calculate total AKTBC from the optimized Excel
        final_df = pd.read_excel(temp_output)
        auto_sum = final_df['AKTBC'].sum()
        
        diff = entry['manual_aktbc_sum'] - auto_sum
        improvement = (diff / entry['manual_aktbc_sum'] * 100) if entry['manual_aktbc_sum'] > 0 else 0
        
        # Cleanup temp file
        try: os.remove(temp_output)
        except: pass
        
        print(f"  Finished {group_id}: Manual {entry['manual_aktbc_sum']} vs Auto {round(auto_sum, 2)}")
        
        return {
            **entry,
            "status": "Success",
            "auto_aktbc_sum": round(float(auto_sum), 2),
            "diff_km": round(float(diff), 2),
            "improvement_pct": round(float(improvement), 2)
        }
    except Exception as e:
        return {**entry, "status": "Error", "error": str(e)[:100]}

def run_benchmark():
    metadata_path = '../RoutingSampleFiles/test_data/suite_metadata.json'
    test_data_dir = '../RoutingSampleFiles/test_data'
    api_key = "AIzaSyAols-dVTGpR4yWBbQppczhzgKwu9xaOKI" 
    
    # Force the use of the verified python interpreter
    python_path = sys.executable 
    
    if not os.path.exists(metadata_path):
        print("Metadata file not found. Run prepare_test_suite.py first.")
        return

    with open(metadata_path, 'r') as f:
        suite = json.load(f)

    print(f"Starting PARALLEL benchmark for {len(suite)} groups (5 workers)...")
    print(f"Using Python: {python_path}")

    with ThreadPoolExecutor(max_workers=5) as executor:
        results = list(executor.map(lambda x: process_group(x, test_data_dir, api_key, python_path), suite))

    # Generate Report
    report_path = '../RoutingSampleFiles/benchmark_report.md'
    with open(report_path, 'w') as f:
        f.write("# Routing Optimization Benchmark Report\n\n")
        f.write(f"Total Groups Tested: {len(suite)}\n")
        success_count = len([r for r in results if r['status'] == 'Success'])
        f.write(f"Successful Runs: {success_count}\n\n")
        
        f.write("| Group ID | Sites | Manual KM | Auto KM | Diff (km) | Improvement % |\n")
        f.write("|----------|-------|-----------|---------|-----------|---------------|\n")
        
        # Sort results by improvement_pct
        sorted_results = sorted(results, key=lambda x: x.get('improvement_pct', -9999), reverse=True)
        
        for r in sorted_results:
            if r['status'] == 'Success':
                indicator = "✅" if r['diff_km'] >= 0 else "❌"
                f.write(f"| {r['group_id']} | {r['num_sites']} | {r['manual_aktbc_sum']} | {r['auto_aktbc_sum']} | {r['diff_km']} | {r['improvement_pct']}% {indicator} |\n")
            else:
                f.write(f"| {r['group_id']} | {r['num_sites']} | {r['manual_aktbc_sum']} | N/A | Error | {r['error']} |\n")

    print(f"Benchmark complete. Report saved to {report_path}")

if __name__ == "__main__":
    run_benchmark()
