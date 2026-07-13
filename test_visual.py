from web_app import app, MODULES_DATA, MODEL, database

token = database.generate_session_token()
database.set_session_token(1, token)
print(f'设置用户 LSW 的 session token: {token}')

with app.test_client() as client:
    resp = client.get('/module/visualization', headers={'Cookie': f'session_token={token}'})
    print('Status:', resp.status_code)
    print('Content length:', len(resp.data))
    if resp.status_code == 500:
        print('500 Error')
        print('Content:', resp.data[:2000].decode('utf-8', errors='ignore'))
    else:
        print('First 500 chars:', resp.data[:500].decode('utf-8', errors='ignore'))
