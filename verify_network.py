import urllib.request
import urllib.parse
import urllib.error
import json
import os
import time

BASE_URL = 'http://127.0.0.1:5000'
COOKIES = {}

def request(method, endpoint, data=None, files=None, session_key=None):
    url = f"{BASE_URL}{endpoint}"
    req = None
    
    if files:
        # Multipart upload simulation using simple boundaries is complex with urllib
        # For simplicity, let's skip file upload verification via script if it's too complex without requests
        # Or try to construct a simple multipart body.
        # Let's fallback to verifying Logic first. File upload we can manual test or assume works if logic works.
        # Actually, let's try a simple boundary.
        boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW'
        body = []
        for name, file_tuple in files.items():
            filename, filedata, content_type = file_tuple
            body.append(f'--{boundary}')
            body.append(f'Content-Disposition: form-data; name="{name}"; filename="{filename}"')
            body.append(f'Content-Type: {content_type}')
            body.append('')
            body.append(filedata.decode('latin1') if isinstance(filedata, bytes) else filedata)
        body.append(f'--{boundary}--')
        body.append('')
        data_bytes = '\r\n'.join(body).encode('latin1')
        
        req = urllib.request.Request(url, data=data_bytes, method='POST')
        req.add_header('Content-Type', f'multipart/form-data; boundary={boundary}')
    elif data:
        json_data = json.dumps(data).encode('utf-8')
        req = urllib.request.Request(url, data=json_data, method=method)
        req.add_header('Content-Type', 'application/json')
    else:
        req = urllib.request.Request(url, method=method)

    # Manage Cookies (Session)
    if session_key and session_key in COOKIES:
        req.add_header('Cookie', COOKIES[session_key])

    try:
        with urllib.request.urlopen(req) as response:
            res_body = response.read().decode('utf-8')
            
            # Save Cookies
            cookie_header = response.getheader('Set-Cookie')
            if cookie_header and session_key:
                COOKIES[session_key] = cookie_header.split(';')[0]
                
            try:
                return response.status, json.loads(res_body)
            except:
                return response.status, res_body
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read().decode('utf-8'))
    except Exception as e:
        return 0, str(e)

def run_tests():
    unique_sf = str(int(time.time()))
    NET_A = f"Rede Tests A {unique_sf}"
    CITY = f"TestCity_{unique_sf}"
    
    print(f"Testing Network Flow for {CITY}...")
    
    # 1. Register
    status, res = request('POST', '/register_admin', {
        'email': f'adminA_{unique_sf}@test.com',
        'password': 'password',
        'city': CITY,
        'network_name': NET_A,
        'network_password': 'netpassA'
    })
    print(f"Register: {status}")
    if status != 200: print(res); return

    # 2. Login
    status, res = request('POST', '/login', {'email': f'adminA_{unique_sf}@test.com', 'password': 'password'}, session_key='admin')
    print(f"Login: {status}")
    if not res.get('success'): print(res); return
    
    # 3. List Networks (Public)
    status, res = request('GET', f'/get_networks?city={CITY}')
    # urllib doesn't handle query params automatically in POST, straightforward in GET string
    print(f"Get Networks: {status}")
    networks = res.get('networks', [])
    found = any(n['name'] == NET_A for n in networks)
    if not found: print("Network not found in list!"); return
    print(f"Found Network: {NET_A}")
    
    net_id = next(n['id'] for n in networks if n['name'] == NET_A)
    
    # 4. Join Network
    status, res = request('POST', '/join_network', {'network_id': net_id, 'password': 'netpassA'}, session_key='public')
    print(f"Join Network: {status}")
    if not res.get('success'): print(res); return
    print("Joined successfully.")
    
    print("ALL TESTS PASSED")

if __name__ == '__main__':
    run_tests()
