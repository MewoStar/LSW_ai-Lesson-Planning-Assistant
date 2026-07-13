import requests
session = requests.Session()
login_resp = session.post('http://127.0.0.1:5024/login', json={'username': 'LSW', 'password': 'LSWYYDS'})
print('login:', login_resp.status_code)

resp = session.post('http://127.0.0.1:5024/api/visual/generate', json={
    'topic': '勾股定理',
    'type': 'mindmap',
    'layout': 'tree'
}, timeout=60)
print('status:', resp.status_code)
data = resp.json()
print('format:', data.get('format'))
print('layout:', data.get('layout'))
content = data.get('content', '')
print('starts_with_flowchart:', content.strip().startswith('flowchart'))
print('content[:300]:')
print(content[:300])
