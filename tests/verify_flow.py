import requests
import os
from openpyxl import Workbook
import time

BASE_URL = 'http://127.0.0.1:5000'
SESSION = requests.Session()

def create_dummy_excel(filename, sheets):
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    for sheet_name, items in sheets.items():
        ws = wb.create_sheet(sheet_name)
        # Header
        ws.append(["Codigo", "Denominação", "Local"])
        for item in items:
            ws.append(item) # [Code, Name, Loc]
            
    wb.save(filename)
    print(f"Created {filename}")

def test_flow():
    # 1. Login
    print("\n--- 1. Testing Login ---")
    res = SESSION.post(f'{BASE_URL}/login', json={'email': 'admin@123', 'password': 'admin123'})
    print(f"Login Status: {res.status_code}")
    if not res.ok: return
    
    # 2. Create Dummy Files
    print("\n--- 2. Creating Dummy Files ---")
    create_dummy_excel('test_master_1.xlsx', {
        'Sala_A': [['CODE_A1', 'Item A1', 'Sala_A'], ['CODE_A2', 'Item A2', 'Sala_A']],
        'Sala_B': [['CODE_B1', 'Item B1', 'Sala_B']]
    })
    create_dummy_excel('test_master_2.xlsx', {
        'Sala_C': [['CODE_C1', 'Item C1', 'Sala_C']], # Intruder!
        'Sala_D': [['CODE_D1', 'Item D1', 'Sala_D']]
    })
    
    # 3. Upload Files
    print("\n--- 3. Uploading Files ---")
    with open('test_master_1.xlsx', 'rb') as f:
        res = SESSION.post(f'{BASE_URL}/upload_master', files={'file': f})
        print(f"Upload 1: {res.json()}")
        
    with open('test_master_2.xlsx', 'rb') as f:
        res = SESSION.post(f'{BASE_URL}/upload_master', files={'file': f})
        print(f"Upload 2: {res.json()}")

    # 4. List Masters
    print("\n--- 4. List Masters ---")
    res = SESSION.get(f'{BASE_URL}/list_masters')
    print(f"Masters: {res.json()}")
    
    # 5. Get Rooms (Selecting both files)
    print("\n--- 5. Get Rooms ---")
    res = SESSION.post(f'{BASE_URL}/get_rooms', json={'filenames': ['test_master_1.xlsx', 'test_master_2.xlsx']})
    rooms = res.json().get('rooms', [])
    print(f"Rooms found: {len(rooms)}")
    for r in rooms:
        print(f" - {r['name']} ({r['id']}) from {r['source']}")
        
    # 6. Verify (The Big Test)
    print("\n--- 6. Verify Analysis ---")
    # Scenario: Auditing 'Sala_A' (in test_master_1)
    # Scanned: 
    # - CODE_A1 (Correct)
    # - CODE_C1 (Intruder from test_master_2)
    # - CODE_UNKNOWN (Completely unknown)
    
    payload = {
        'analyst_name': 'Tester Bot',
        'room_name': 'Sala_A',
        'source_file': 'test_master_1.xlsx',
        'selected_files': ['test_master_1.xlsx', 'test_master_2.xlsx'],
        'scanned_codes': "CODE_A1\nCODE_C1\nCODE_UNKNOWN"
    }
    
    res = SESSION.post(f'{BASE_URL}/verify', json=payload)
    print(f"Verify Status: {res.status_code}")
    
    if res.ok:
        with open('test_report.zip', 'wb') as f:
            f.write(res.content)
        print("Report saved to test_report.zip")
        
        # Check inside zip?
        import zipfile
        with zipfile.ZipFile('test_report.zip', 'r') as z:
            print("Files in ZIP:", z.namelist())
            # Expecting: Tester Bot_Verificados.xlsx, Tester Bot_Nao_Encontrados.xlsx, Tester Bot_Local_Incorreto.xlsx
            
    else:
        print("Verify Failed:", res.text)

    # 7. Check server report list
    print("\n--- 7. Check Reports on Server ---")
    res = SESSION.get(f'{BASE_URL}/list_reports')
    print(f"Server Reports: {res.json()}")

    # Clean up local dummy files
    try:
        os.remove('test_master_1.xlsx')
        os.remove('test_master_2.xlsx')
        # os.remove('test_report.zip') # Keep for manual check if needed
    except: pass

if __name__ == '__main__':
    try:
        test_flow()
    except Exception as e:
        print(f"Test Failed: {e}")
