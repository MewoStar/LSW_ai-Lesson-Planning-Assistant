import requests

BASE_URL = 'http://127.0.0.1:5024'

login_resp = requests.post(f'{BASE_URL}/login', json={'username': 'LSW', 'password': 'LSWYYDS'})
cookies = login_resp.cookies

resp = requests.post(f'{BASE_URL}/api/visual/generate', 
                    json={'topic': '牛顿第一定律', 'type': 'mindmap', 'layout': 'logic'}, 
                    cookies=cookies, timeout=120)
print(f"状态码: {resp.status_code}")
print(f"完整响应: {resp.text}")