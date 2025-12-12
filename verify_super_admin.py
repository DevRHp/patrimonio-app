import urllib.request
import urllib.parse
import json
import time

BASE_URL = 'http://127.0.0.1:5000'
COOKIES = {}

def request(method, endpoint, data=None, session_key=None):
    url = f"{BASE_URL}{endpoint}"
    req = None
    
    if data:
        json_data = json.dumps(data).encode('utf-8')
        req = urllib.request.Request(url, data=json_data, method=method)
        req.add_header('Content-Type', 'application/json')
    else:
        req = urllib.request.Request(url, method=method)

    if session_key and session_key in COOKIES:
        req.add_header('Cookie', COOKIES[session_key])

    try:
        with urllib.request.urlopen(req) as response:
            res_body = response.read().decode('utf-8')
            cookie = response.getheader('Set-Cookie')
            if cookie and session_key:
                COOKIES[session_key] = cookie.split(';')[0]
            try:
                return response.status, json.loads(res_body)
            except:
                return response.status, res_body
    except Exception as e:
        return 0, str(e)

def run_test():
    ts = str(int(time.time()))
    
    # 1. Create a "Foreign" Admin (City: Paris)
    email_foreign = f"paris_admin_{ts}@test.com"
    print(f"Creating Foreign Admin: {email_foreign}")
    status, res = request('POST', '/register_admin', {
        'email': email_foreign, 'password': 'pass', 
        'city': 'Paris', 'network_name': f'ParisNet_{ts}', 'network_password': 'np'
    })
    if status != 200: print("Failed to register foreign admin:", res); return

    # Login Foreign Admin
    print("Logging in Foreign Admin...")
    status, res = request('POST', '/login', {'email': email_foreign, 'password': 'pass'}, session_key='foreign')
    if not res.get('success'): print("Foreign login failed:", res); return
    
    # Foreign Admin uploads a file (We can't easily upload via this simple script without multipart)
    # But wait, we just need to verify list_masters logic.
    # Actually, simpler test: verify list_masters returns EMPTY for foreign admin initially.
    # Then verify Admin@123 returns ALL.
    # But Admin@123 will return "valid_files" from DISK. 
    # If I don't upload a real file, list_masters won't show it because of checks:
    # `if os.path.exists...`
    # So I *must* upload a file or manually create one in uploads/ folder and insert to DB.
    
    print("Simulating File Upload...")
    import sqlite3
    
    # Manually insert into DB to bypass upload complexity
    # We need to find where the DB is. Assuming d:/patrimonio/backend/database.db
    # But the script might run from root.
    # Let's try to assume standard path.
    db_path = 'backend/database.db'
    
    # We need to create a dummy file on disk too.
    dummy_filename = f"Paris_Master_{ts}.xlsx"
    with open(f"uploads/{dummy_filename}", "w") as f: f.write("dummy")
    
    # We need the foreign admin User ID.
    # We can get it from login response? No, login response has city/netname.
    # We can query DB.
    
    try:
        con = sqlite3.connect(db_path)
        cur = con.cursor()
        cur.execute("SELECT id FROM users WHERE email = ?", (email_foreign,))
        foreign_id = cur.fetchone()[0]
        
        cur.execute("INSERT INTO files (filename, city, user_id) VALUES (?, ?, ?)", 
                    (dummy_filename, 'Paris', foreign_id))
        con.commit()
        con.close()
    except Exception as e:
        print(f"DB Error (make sure to run from root d:/patrimonio): {e}")
        return

    # 2. Check Foreign Admin sees it
    print("Checking Foreign Admin View...")
    status, res = request('GET', '/list_masters', session_key='foreign')
    masters = res.get('masters', [])
    if dummy_filename in masters:
        print("PASS: Foreign Admin sees their file.")
    else:
        print("FAIL: Foreign Admin DOES NOT see their file.")

    # 3. Login Super Admin
    print("Logging in Super Admin...")
    status, res = request('POST', '/login', {'email': 'admin@123', 'password': 'admin123'}, session_key='super')
    if not res.get('success'): print("Super Admin login failed:", res); return
    
    # 4. Check Super Admin sees it
    print("Checking Super Admin View...")
    status, res = request('GET', '/list_masters', session_key='super')
    masters = res.get('masters', [])
    if dummy_filename in masters:
        print("PASS: Super Admin sees the foreign file.")
    else:
        print("FAIL: Super Admin DOES NOT see the foreign file.")
        print("Seen:", masters)

if __name__ == '__main__':
    run_test()
