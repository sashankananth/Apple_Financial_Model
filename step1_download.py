import json
import requests

# --- Settings -------------------------------------------------------------
CIK = "0000320193"  # Apple Inc.
URL = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json"

# The SEC requires a descriptive User-Agent: your name and a real email.
HEADERS = {"User-Agent": "Your Name your.email@example.com"}

# --- Request ----------------------------------------------------------------
# requests.get sends the HTTP request. timeout stops it hanging forever.
try:
    response = requests.get(URL, headers=HEADERS, timeout=30)
    response.raise_for_status()  # raises an error for 403, 404, 500, etc.
except requests.exceptions.RequestException as error:
    print(f"Request failed: {error}")
    raise SystemExit(1)

# --- Parse and save ---------------------------------------------------------
data = response.json()  # converts the JSON text into a Python dictionary

with open("apple_companyfacts.json", "w") as file:
    json.dump(data, file)  # save a local copy so we don't re-download each run

# --- Basic validation -------------------------------------------------------
print("Company:", data.get("entityName"))
print("CIK:", data.get("cik"))

us_gaap = data.get("facts", {}).get("us-gaap", {})
print("Number of us-gaap tags:", len(us_gaap))

# Peek at one tag to see the structure
sample = us_gaap.get("NetIncomeLoss")
if sample:
    print("Units for NetIncomeLoss:", list(sample["units"].keys()))
    print("Example data point:", sample["units"]["USD"][-1])
else:
    print("NetIncomeLoss tag not found - something is wrong.")