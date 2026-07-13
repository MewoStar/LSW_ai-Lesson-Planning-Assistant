import sys
sys.path.insert(0, '.')
# 强制重新加载
if 'web_app' in sys.modules:
    del sys.modules['web_app']

from web_app import app

with app.test_client() as client:
    # 登录
    client.post('/login', json={'username': 'LSW', 'password': 'LSWYYDS'})
    
    # 测试tree布局
    resp = client.post('/api/visual/generate', json={
        'topic': '勾股定理',
        'type': 'mindmap',
        'layout': 'tree'
    })
    data = resp.get_json()
    print('format:', data.get('format'))
    print('layout:', data.get('layout'))
    content = data.get('content', '')
    print('starts_with_flowchart:', content.strip().startswith('flowchart'))
    print('content[:200]:', content[:200])
