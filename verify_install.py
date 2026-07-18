import sys
import importlib.util

pyc_path = r'C:\Program Files\AI备课助手\__pycache__\web_app.cpython-312.pyc'
py_path = r'C:\Program Files\AI备课助手\web_app.py'

import os
print('py 文件存在:', os.path.exists(py_path))
print('pyc 文件存在:', os.path.exists(pyc_path))

spec = importlib.util.spec_from_file_location('web_app_install', py_path)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

app = mod.app

print()
print('=== /logout 路由信息 ===')
for rule in app.url_map.iter_rules():
    if rule.rule == '/logout':
        print(f'  规则: {rule.rule}')
        print(f'  方法: {sorted(rule.methods)}')
        print(f'  端点: {rule.endpoint}')

print()
print('=== 测试客户端 ===')
client = app.test_client()
r = client.get('/logout', follow_redirects=False)
print(f'  GET /logout 状态码: {r.status_code}')
print(f'  Location: {r.headers.get("Location")}')
r = client.post('/logout', follow_redirects=False)
print(f'  POST /logout 状态码: {r.status_code}')
