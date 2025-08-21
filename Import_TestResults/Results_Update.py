import pandas as pd
import requests
import math

# ==== CONFIGURATION ====
testrail_url = "https://zola.testrail.io/"  # e.g. https://mycompany.testrail.io
username = "swapna.arepalli@zolaelectric.com"
api_key = "BMzZr8hwukFCZ/u2l4zo-8XjFkr7Pu14Nst1ohnly"  # or password if API key not enabled
#password = "Saibaba@12345"
run_id = 1617  # e.g. 1234

# Status mapping (TestRail default)
status_map = {
    'passed': 1,
    'failed': 5,
    'blocked': 2,
    'retest': 4,
    'untested': 3
}

# ==== READ EXCEL FILE ====
df = pd.read_excel("aloha_firmware.xlsx")

results = []

for _, row in df.iterrows():
    case_id_raw = str(row['Case ID']).strip()

    # Skip empty IDs
    if not case_id_raw or case_id_raw.lower() == 'nan':
        continue

    # Remove 'C' prefix if present and convert to int
    case_id = int(case_id_raw.replace('C', '').strip())

    # Status handling
    status_raw = str(row['Status']).strip().lower()

    if status_raw.isdigit():  # already numeric
        status_id = int(float(status_raw))  # handle 1.0 case
    else:
        status_id = status_map.get(status_raw)

    if not status_id:
        print(f"Skipping unknown status for case {case_id_raw}: {status_raw}")
        continue

    # Add result
    results.append({
        "case_id": case_id,
        "status_id": status_id,
        "comment": str(row.get('Comment', '')).strip()
    })

# ==== SEND TO TESTRAIL ====
if results:
    url = f"{testrail_url}/index.php?/api/v2/add_results_for_cases/{run_id}"
    response = requests.post(url, auth=(username, api_key), json={"results": results})
    #response = requests.post(url, auth=(username, password), json={"results": results})

    if response.status_code == 200:
        print(f"✅ Successfully updated {len(results)} cases in TestRail.")
    else:
        print(f"❌ Error updating TestRail: {response.status_code} - {response.text}")
else:
    print("No valid test results found to import.")
