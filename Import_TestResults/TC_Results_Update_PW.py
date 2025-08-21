import pandas as pd
import requests
import argparse

# ==== ARGUMENT PARSER ====
parser = argparse.ArgumentParser(description="Upload test results to TestRail from Excel using password")
parser.add_argument("--url", help="TestRail base URL, e.g. https://mycompany.testrail.io")
parser.add_argument("--username", help="TestRail username/email")
parser.add_argument("--password", help="TestRail password for authentication")
parser.add_argument("--run_id", type=int, help="TestRail Run ID")
parser.add_argument("--file", help="Path to Excel file containing results")

args = parser.parse_args()

# ==== INTERACTIVE PROMPTS (if missing args) ====
testrail_url = (args.url or "").strip()
if not testrail_url:
    testrail_url = input("🔗 Enter TestRail URL (e.g. https://mycompany.testrail.io): ").strip()
testrail_url = testrail_url.rstrip("/")

username = (args.username or "").strip()
if not username:
    username = input("👤 Enter TestRail Username/Email: ").strip()

password = (args.password or "").strip()
if not password:
    password = input("🔑 Enter TestRail Password (visible in console): ").strip()

run_id = args.run_id
if not run_id:
    run_id = int(input("📌 Enter TestRail Run ID: ").strip())

excel_file = (args.file or "").strip()
if not excel_file:
    excel_file = input("📄 Enter path to Excel file: ").strip()

# ==== STATUS MAPPING ====
status_map = {
    'passed': 1,
    'failed': 4,
    'blocked': 2,
    'retest': 5,
    'untested': 3
}

# ==== READ EXCEL FILE ====
df = pd.read_excel(excel_file)
results = []

for _, row in df.iterrows():
    case_id_raw = str(row['Case ID']).strip()

    if not case_id_raw or case_id_raw.lower() == 'nan':
        continue

    case_id = int(case_id_raw.replace('C', '').strip())

    status_raw = str(row['Status']).strip().lower()
    if status_raw.isdigit():
        status_id = int(float(status_raw))
    else:
        status_id = status_map.get(status_raw)

    if not status_id:
        print(f"⚠️ Skipping unknown status for case {case_id_raw}: {status_raw}")
        continue

    results.append({
        "case_id": case_id,
        "status_id": status_id,
        "comment": str(row.get('Comment', '')).strip()
    })

# ==== SEND TO TESTRAIL ====
if results:
    url = f"{testrail_url}/index.php?/api/v2/add_results_for_cases/{run_id}"
    response = requests.post(url, auth=(username, password), json={"results": results})

    if response.status_code == 200:
        print(f"✅ Successfully updated {len(results)} cases in TestRail.")
    else:
        print(f"❌ Error updating TestRail: {response.status_code} - {response.text}")
else:
    print("⚠️ No valid test results found to import.")
