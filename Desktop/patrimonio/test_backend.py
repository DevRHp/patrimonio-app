import requests
import os
from openpyxl import Workbook

# 1. Create Dummy Master Spreadsheet
wb = Workbook()
ws1 = wb.active
ws1.title = "Room A"
ws1.append(["Code", "Name"])
ws1.append(["A001", "Item A1"])
ws1.append(["A002", "Item A2"])
ws1.append(["A003", "Item A3"])

ws2 = wb.create_sheet("Room B")
ws2.append(["Code", "Name"])
ws2.append(["B001", "Item B1"])
ws2.append(["B002", "Item B2"])

wb.save("dummy_master.xlsx")
print("Created dummy_master.xlsx")

# 2. Test Upload
url = "http://127.0.0.1:5000"
files = {'file': open('dummy_master.xlsx', 'rb')}
try:
    response = requests.post(f"{url}/upload_master", files=files)
    print(f"Upload Status: {response.status_code}")
    print(f"Upload Response: {response.json()}")
except Exception as e:
    print(f"Upload failed (is server running?): {e}")
    exit()

# 3. Test Verify
# Scenarios:
# - A001: Correct (in Room A)
# - B001: Wrong Location (in Room B, but we are checking Room A)
# - A002: Missing (in Room A, but not scanned)
# - C001: Unknown (not in any room) - Logic handles this as "Not found in master" if we scan it.

scanned_codes = """
A001
B001
C001
"""

payload = {
    "analyst_name": "TestUser",
    "room_name": "Room A",
    "scanned_codes": scanned_codes
}

try:
    response = requests.post(f"{url}/verify", json=payload)
    print(f"Verify Status: {response.status_code}")
    
    if response.status_code == 200:
        with open("test_results.zip", "wb") as f:
            f.write(response.content)
        print("Saved test_results.zip")
    else:
        print(f"Verify Error: {response.text}")

except Exception as e:
    print(f"Verify failed: {e}")
