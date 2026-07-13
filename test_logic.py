import requests

BASE_URL = 'http://127.0.0.1:5024'

print("=== 测试逻辑图生成 ===")

login_resp = requests.post(f'{BASE_URL}/login', json={'username': 'LSW', 'password': 'LSWYYDS'})
cookies = login_resp.cookies
print(f"登录: {login_resp.status_code}")

resp = requests.post(f'{BASE_URL}/api/visual/generate', 
                    json={'topic': '牛顿第一定律', 'type': 'mindmap', 'layout': 'logic'}, 
                    cookies=cookies, timeout=120)
data = resp.json()
print(f"\n生成结果:")
print(f"  成功: {data.get('success')}")
print(f"  布局: {data.get('layout')}")
print(f"  格式: {data.get('format')}")
if data.get('content'):
    print(f"\nMermaid代码:\n{data['content']}")