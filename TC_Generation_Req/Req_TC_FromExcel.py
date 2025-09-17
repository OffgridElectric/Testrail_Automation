# Test Cases Generation from Requirements with Excel file via OpenAI and importing into Testrail
import json
import requests
import pandas as pd
from openai import OpenAI

# === Collect User Inputs ===
testrail_url = input("Enter TestRail URL (e.g., https://yourcompany.testrail.io): ").strip()
testrail_user = input("Enter TestRail username/email: ").strip()
testrail_api_key = input("Enter TestRail API key: ").strip()
project_id = input("Enter TestRail Project ID: ").strip()
section_id = input("Enter TestRail Section ID: ").strip()
openai_api_key = input("Enter OpenAI API key: ").strip()
excel_path = input("Enter path to Excel file with requirements (.xlsx): ").strip()

# === Init OpenAI ===
client = OpenAI(api_key=openai_api_key)

def read_requirements_from_excel(path):
    """Read requirements from first column of Excel"""
    df = pd.read_excel(path, engine="openpyxl")
    requirements = df.iloc[:, 0].dropna().astype(str).tolist()
    return requirements

def generate_test_cases(requirement, model="gpt-4o-mini"):
    """Generate structured test cases from requirement using OpenAI"""
    prompt = f"""
    You are a QA engineer. From the following requirement, generate 2–3 structured test cases.

    Each test case must be **pure JSON only**, with this schema:
    [
      {{
        "title": "string",
        "steps": [
          {{"content": "step action", "expected": "expected result"}}
        ]
      }}
    ]

    Requirement:
    {requirement}
    """

    response = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0
    )

    content = response.choices[0].message.content.strip()

    # Ensure no code fences
    if content.startswith("```"):
        content = content.split("```")[1]  # take middle
        content = content.replace("json", "", 1).strip()

    return json.loads(content)

def add_test_case_to_testrail(title, steps):
    """Add a test case to TestRail using API"""
    url = f"{testrail_url}/index.php?/api/v2/add_case/{section_id}"
    headers = {"Content-Type": "application/json"}

    # Map to TestRail Steps Separated format
    formatted_steps = [{"content": s["content"], "expected": s["expected"]} for s in steps]

    payload = {
        "title": title,
        "type_id": 1,   # Adjust if needed
        "priority_id": 2,
        "custom_steps_separated": formatted_steps
    }

    response = requests.post(url, headers=headers, auth=(testrail_user, testrail_api_key), json=payload)

    if response.status_code == 200:
        print(f"✅ Test case '{title}' added to TestRail.")
    else:
        print(f"❌ Failed to add '{title}'. Response: {response.text}")

# === Run Flow ===
try:
    requirements = read_requirements_from_excel(excel_path)
    print(f"\n📌 Found {len(requirements)} requirements in {excel_path}")

    for req in requirements:
        print(f"\n➡️ Processing requirement: {req}")
        try:
            test_cases = generate_test_cases(req)
            for tc in test_cases:
                add_test_case_to_testrail(tc["title"], tc["steps"])
        except Exception as e:
            print("⚠️ Error generating or importing test cases:", e)

except Exception as e:
    print("⚠️ Failed to read Excel file:", e)
