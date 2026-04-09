import pandas as pd
import numpy as np

f_raw = "km required 2.xlsx"
f_calculated = "km_calculated_v2.xlsx"
f_manual = "KM A6 MMWAVE.xlsx"
f_out = "km_required_2_enriched.xlsx"

# 1. Load Raw Data
df_raw = pd.read_excel(f_raw)

# 2. Extract MIN DATE from km_calculated_v2 (106 sites)
df_calc = pd.read_excel(f_calculated)
min_date_map = dict(zip(df_calc['SITE ID'].astype(str).str.strip(), df_calc['MIN DATE']))

# Apply to raw data
df_raw['SITE ID CLEAN'] = df_raw['SITE ID'].astype(str).str.strip()
df_raw['MIN DATE'] = df_raw['SITE ID CLEAN'].map(min_date_map)
df_raw = df_raw.drop(columns=['SITE ID CLEAN'])

# 3. Extract the 5 missing sites from manual MMWave sheet
df_man = pd.read_excel(f_manual)
missing_sites = ['I-RJ-DENA-ENB-6082', 'I-RJ-PNDW-ENB-6000', 'I-RJ-RAPR-ENB-9005', 'I-RJ-JLOR-ENB-V001', 'I-RJ-SNCE-ENB-A003']
df_missing = df_man[df_man['eNBsiteID'].astype(str).str.strip().isin(missing_sites)].copy()

jc_to_cmp = {
    'Jalor': 'Sirohi',
    'Makrana': 'Nagaur',
    'Pali': 'Sirohi',
    'Sirohi': 'Sirohi'
}

new_rows = []
for _, row in df_missing.iterrows():
    jc = str(row.get('JC', '')).strip()
    cmp_val = jc_to_cmp.get(jc, 'UNKNOWN')
    
    new_row = {
        'SITE ID': row.get('eNBsiteID'),
        'PMP ID': row.get('PMP ID'),
        'NO OF\\n SECTOR': row.get('NO OF SECTOR'),
        'BAND': 'MM WAVE', # Explicitly MM WAVE
        'JC NAME': jc,
        'CMP': cmp_val,
        'Current Status': row.get('Activity'), # Approx mapping
        'LATITUDE': row.get('LAT '),
        'LONGITUDE': row.get('LONG'),
        'INTEGRATION\\nDATE': np.nan,
        'ATP 11A DATE': np.nan,
        'ATP11C': np.nan,
        'WH': row.get('WAREHOUSE'),
        'MIN DATE': row.get('MIN DATE')
    }
    new_rows.append(new_row)

df_new = pd.DataFrame(new_rows)

# 4. Append and Save
df_final = pd.concat([df_raw, df_new], ignore_index=True)

with pd.ExcelWriter(f_out, engine='xlsxwriter') as writer:
    df_final.to_excel(writer, index=False)
    
print(f"Successfully injected 5 manual sites and MIN DATE column into {f_out}")
print(f"Total rows in new raw dataset: {len(df_final)}")
