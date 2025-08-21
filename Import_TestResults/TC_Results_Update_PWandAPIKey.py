import pandas as pd
import requests
import argparse

# ==== ARGUMENT PARSER ====
parser = argparse.ArgumentParser(description="Upload test results to TestRail from Excel (API key or password)")
parser.add_argument("--url", help="TestRail base URL, e.g. https://mycompany.testrail.io")
parser.add_argument("--username", help="TestRail username/email")
parser.add_argument("--password", help="TestRail password (if API key not provided)")
parser.add_argument("--api_key", help="TestRail API key (preferred)")
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

# API Key or Password
api_key = (getattr(args, 'api_key', None) or "").strip()
password = (getattr(args, 'password', None) or "").strip()

if not api_key and not password:
    choice = input("Do you want to use API Key or Password for authentication? (key/password): ").strip().lower()
    if choice == 'key':
        api_key = input("🔑 Enter TestRail API Key (visible in PyCharm/console): ").strip()
    else:
        password = input("🔑 Enter TestRail Password (visible in PyCharm/console): ").strip()

run_id = args.run_id
if not run_id:
    run_id = int(input("📌 Enter TestRail Run ID: ").strip())

excel_file = (args.file or "").strip()
if not excel_file:
    excel_file = input("📄 Enter path to Excel file: ").strip()

# ==== STATUS MAPPING ====
status_map = {
    'passed': 1,
    'failed': 5,
    'blocked': 2,
    'retest': 4,
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
    if api_key:
        auth = (username, api_key)
    elif password:
        auth = (username, password)
    else:
        print("❌ No authentication method provided. Cannot proceed.")
        exit(1)

    response = requests.post(url, auth=auth, json={"results": results})

    if response.status_code == 200:
        print(f"✅ Successfully updated {len(results)} cases in TestRail.")
    elif response.status_code == 401:
        print("❌ Authentication failed. Check your API key or password. 401 Unauthorized")
    else:
        print(f"❌ Error updating TestRail: {response.status_code} - {response.text}")
else:
    print("⚠️ No valid test results found to import.")
