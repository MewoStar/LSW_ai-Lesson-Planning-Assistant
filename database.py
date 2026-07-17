import sqlite3
import uuid
import re
from pathlib import Path
import supabase_client

def get_db_path():
    if getattr(__import__('sys'), 'frozen', False):
        return Path(__import__('sys').executable).parent / 'app.db'
    return Path(__file__).parent / 'app.db'

DB_PATH = get_db_path()

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT UNIQUE NOT NULL,
            username TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute("PRAGMA table_info(users)")
    columns = [row[1] for row in cursor.fetchall()]
    if 'user_id' not in columns:
        cursor.execute("ALTER TABLE users ADD COLUMN user_id TEXT")
        cursor.execute("UPDATE users SET user_id = 'user_' || id WHERE user_id IS NULL")
        cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_user_id ON users(user_id)")
    
    if 'username' not in columns:
        cursor.execute("ALTER TABLE users ADD COLUMN username TEXT")
    
    if 'email' not in columns:
        cursor.execute("ALTER TABLE users ADD COLUMN email TEXT")

    if 'avatar_url' not in columns:
        cursor.execute("ALTER TABLE users ADD COLUMN avatar_url TEXT")

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS user_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL,
            session_id TEXT UNIQUE NOT NULL,
            title TEXT DEFAULT '新对话',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS chat_messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_session_id INTEGER NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_session_id) REFERENCES user_sessions (id) ON DELETE CASCADE
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS search_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL,
            query TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            admin_username TEXT NOT NULL,
            action TEXT NOT NULL,
            target_user_id TEXT,
            details TEXT,
            ip_address TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_login_attempts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ip_address TEXT NOT NULL,
            username TEXT,
            success INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 管理员个人资料表（独立于 users 表，避免 UNIQUE 约束冲突）
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_profile (
            user_id TEXT PRIMARY KEY,
            username TEXT,
            avatar_url TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS session_tokens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL,
            token TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_session_tokens_token ON session_tokens(token)')

    conn.commit()
    conn.close()

def register_user(username, password, email=None):
    if not email:
        return False
    # M2: 注册前校验用户名是否已在 Supabase 存在，避免重复注册
    try:
        exists_result = supabase_client.username_exists(username)
        if exists_result.get("success") and exists_result.get("exists"):
            return False
    except Exception:
        # 校验异常不阻断注册流程，交由 sign_up 自行处理
        pass
    result = supabase_client.sign_up(email, password, username)
    if result['success']:
        user_id = result['user']['id']
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            # M6: 使用 ON CONFLICT 替代 INSERT OR REPLACE，仅按 user_id 冲突时更新
            cursor.execute(
                'INSERT INTO users (user_id, username, email) VALUES (?, ?, ?) '
                'ON CONFLICT(user_id) DO UPDATE SET username=excluded.username, email=excluded.email',
                (user_id, username, email)
            )
            conn.commit()
        except Exception as e:
            # H2: 本地写入失败时尝试回滚 Supabase 端已创建的用户，避免数据不一致
            try:
                supabase_client.delete_user(user_id)
            except Exception:
                pass
            return False
        finally:
            conn.close()
        return True
    return False

def get_email_by_username(username):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT email FROM users WHERE username = ?', (username,))
        row = cursor.fetchone()
        if row:
            return row[0]
        return None
    finally:
        conn.close()

def get_all_local_users():
    """从本地数据库获取所有用户"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT user_id, username, email, created_at FROM users ORDER BY created_at DESC')
        rows = cursor.fetchall()
        users = []
        for row in rows:
            users.append({
                "id": row[0],
                "username": row[1],
                "email": row[2],
                "created_at": row[3],
                "role": "user",
                "last_sign_in_at": None
            })
        return users
    finally:
        conn.close()

def delete_local_user(user_id):
    """从本地数据库删除用户"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        # M1: 使用 with conn 事务块包裹多条 DELETE，任一失败自动回滚，保证数据一致性
        with conn:
            cursor.execute('DELETE FROM users WHERE user_id = ?', (user_id,))
            cursor.execute('DELETE FROM user_sessions WHERE user_id = ?', (user_id,))
            cursor.execute('DELETE FROM search_history WHERE user_id = ?', (user_id,))
        return True
    except Exception as e:
        print(f"删除本地用户失败: {e}")
        return False
    finally:
        conn.close()

def authenticate_user(email_or_username, password):
    # H1: 支持用户名或邮箱登录，非邮箱格式时先转换为邮箱再调用 sign_in
    if not email_or_username or not password:
        return None
    email = email_or_username
    # 判断是否为邮箱格式，若不是则按用户名查询对应邮箱
    if not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$', email_or_username):
        result = supabase_client.get_email_by_username(email_or_username)
        if not result.get("success") or not result.get("email"):
            return None
        email = result["email"]
    result = supabase_client.sign_in(email, password)
    if result['success']:
        user = result['user']
        session = result.get('session') or {}
        # M3: 空值安全处理，避免 user['username'] 或 user['email'] 为空时抛异常
        email_val = user.get('email') or ''
        username_val = user.get('username') or (email_val.split('@')[0] if email_val else 'user')
        return {
            'id': user['id'],
            'username': username_val,
            'email': email_val,
            'access_token': session.get('access_token')
        }
    return None

def generate_session_token():
    return str(uuid.uuid4())

def set_session_token(user_id, token):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        # M6: 使用 ON CONFLICT(token) 替代 INSERT OR REPLACE，按 token 唯一约束更新
        cursor.execute(
            'INSERT INTO session_tokens (user_id, token, created_at) VALUES (?, ?, CURRENT_TIMESTAMP) '
            'ON CONFLICT(token) DO UPDATE SET user_id=excluded.user_id, created_at=CURRENT_TIMESTAMP',
            (user_id, token)
        )
        conn.commit()
    finally:
        conn.close()

def get_user_by_token(token):
    result = supabase_client.get_user(token)
    if result['success']:
        user = result['user']
        # M3: 空值安全处理，避免 user['username'] 或 user['email'] 为空时抛异常
        email_val = user.get('email') or ''
        username_val = user.get('username') or (email_val.split('@')[0] if email_val else 'user')
        return {
            'id': user['id'],
            'username': username_val,
            'email': email_val
        }
    return None

def clear_session_token(token):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('DELETE FROM session_tokens WHERE token = ?', (token,))
        conn.commit()
    finally:
        conn.close()

def is_valid_session_token(token):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        # L3: 注意 SQLite 的 datetime('now') 使用 UTC 时间，与写入时 CURRENT_TIMESTAMP 一致
        cursor.execute('SELECT user_id FROM session_tokens WHERE token = ? AND created_at > datetime("now", "-7 days")',
                      (token,))
        row = cursor.fetchone()
        return row is not None
    finally:
        conn.close()

def save_user_chat_session(user_id, session_id, title='新对话'):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        # M6: 使用 ON CONFLICT(session_id) 替代 INSERT OR REPLACE，按 session_id 唯一约束更新标题
        cursor.execute(
            'INSERT INTO user_sessions (user_id, session_id, title) VALUES (?, ?, ?) '
            'ON CONFLICT(session_id) DO UPDATE SET title=excluded.title',
            (user_id, session_id, title)
        )
        conn.commit()
        return cursor.lastrowid
    finally:
        conn.close()

def get_user_sessions(user_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT session_id, title, created_at FROM user_sessions WHERE user_id = ? ORDER BY created_at DESC',
            (user_id,)
        )
        sessions = []
        for row in cursor.fetchall():
            sessions.append({
                'session_id': row[0],
                'title': row[1],
                'created_at': row[2]
            })
        return sessions
    finally:
        conn.close()

def save_chat_message(user_session_id, role, content):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'INSERT INTO chat_messages (user_session_id, role, content) VALUES (?, ?, ?)',
            (user_session_id, role, content)
        )
        conn.commit()
    finally:
        conn.close()

def get_chat_messages(user_session_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT role, content, created_at FROM chat_messages WHERE user_session_id = ? ORDER BY created_at',
            (user_session_id,)
        )
        messages = []
        for row in cursor.fetchall():
            messages.append({
                'role': row[0],
                'content': row[1],
                'created_at': row[2]
            })
        return messages
    finally:
        conn.close()

def save_search_history(user_id, query):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'INSERT INTO search_history (user_id, query) VALUES (?, ?)',
            (user_id, query)
        )
        conn.commit()
    finally:
        conn.close()

def get_search_history(user_id, limit=50):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT query, created_at FROM search_history WHERE user_id = ? ORDER BY created_at DESC LIMIT ?',
            (user_id, limit)
        )
        history = []
        for row in cursor.fetchall():
            history.append({
                'query': row[0],
                'created_at': row[1]
            })
        return history
    finally:
        conn.close()

def delete_user_session(user_id, session_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'DELETE FROM user_sessions WHERE user_id = ? AND session_id = ?',
            (user_id, session_id)
        )
        conn.commit()
    finally:
        conn.close()

def update_session_title(user_id, session_id, title):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'UPDATE user_sessions SET title = ? WHERE user_id = ? AND session_id = ?',
            (title, user_id, session_id)
        )
        conn.commit()
    finally:
        conn.close()

def get_user_session_pk(user_id, session_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT id FROM user_sessions WHERE user_id = ? AND session_id = ? LIMIT 1',
            (user_id, session_id)
        )
        row = cursor.fetchone()
        return row[0] if row else None
    finally:
        conn.close()

def _guess_module_id_from_text(text):
    if not isinstance(text, str) or not text:
        return None
    t = text[:1500]
    if ('[PPT:' in t) or ('ppt' in t.lower() and ('幻灯片' in t or '课件' in t)):
        return 'ppt_generate'
    if ('教案' in t) or ('教学目标' in t and '教学重难点' in t):
        return 'lesson_plan_generate'
    if ('习题' in t) or ('选择题' in t and '填空题' in t) or ('试卷' in t):
        return 'exercise_generate'
    if ('作业' in t) and (('错题' in t) or ('平均分' in t) or ('得分率' in t) or ('知识点掌握' in t)):
        return 'homework_analysis'
    if ('讲解纲要' in t) or ('演讲纲要' in t) or ('讲解要点' in t) or ('提问设计' in t):
        return 'ppt_outline'
    if ('[PPT:' in t) or ('课件' in t and '第' in t and '页' in t):
        return 'ppt_generate'
    return None

def _extract_files_from_text(text):
    if not isinstance(text, str) or not text:
        return []
    import re
    files = []
    for m in re.finditer(r'\[(文件|PPT|教案|练习册|作业分析|报告)\s*[:：]\s*([^\]]{1,120})\]', text):
        name = m.group(2).strip()
        kind = m.group(1)
        files.append({"kind": kind, "name": name})
    return files[:8]

def get_user_sessions_with_stats(user_id):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute(
            '''SELECT s.id as pk, s.session_id, s.title, s.created_at,
                      SUM(CASE WHEN m.role = 'user' THEN 1 ELSE 0 END) as user_cnt,
                      SUM(CASE WHEN m.role = 'assistant' THEN 1 ELSE 0 END) as assistant_cnt,
                      COUNT(m.id) as total_cnt,
                      MAX(m.created_at) as last_msg_at
               FROM user_sessions s
               LEFT JOIN chat_messages m ON m.user_session_id = s.id
               WHERE s.user_id = ?
               GROUP BY s.id, s.session_id, s.title, s.created_at
               ORDER BY COALESCE(MAX(m.created_at), s.created_at) DESC''',
            (user_id,)
        )
        rows = cursor.fetchall()
        result = []
        for r in rows:
            pk = r['pk']
            sid = r['session_id']
            cursor.execute(
                '''SELECT role, content, created_at FROM chat_messages
                   WHERE user_session_id = ? AND role = 'user'
                   ORDER BY created_at ASC LIMIT 1''',
                (pk,)
            )
            first_user_row = cursor.fetchone()
            cursor.execute(
                '''SELECT role, content, created_at FROM chat_messages
                   WHERE user_session_id = ? AND role = 'assistant'
                   ORDER BY created_at DESC LIMIT 1''',
                (pk,)
            )
            last_asst_row = cursor.fetchone()

            first_user = first_user_row['content'] if first_user_row else ""
            last_asst = last_asst_row['content'] if last_asst_row else ""

            mod = None
            for txt in (r['title'] or "", first_user, last_asst):
                mod = _guess_module_id_from_text(txt)
                if mod:
                    break

            cursor.execute(
                '''SELECT content FROM chat_messages
                   WHERE user_session_id = ? AND role = 'assistant'
                   ORDER BY created_at DESC LIMIT 12''',
                (pk,)
            )
            all_files = []
            seen = set()
            for mr in cursor.fetchall():
                for f in _extract_files_from_text(mr['content'] or ""):
                    key = (f['kind'], f['name'])
                    if key in seen:
                        continue
                    seen.add(key)
                    all_files.append(f)

            total_cnt = r['total_cnt'] or 0
            user_cnt = r['user_cnt'] or 0
            assistant_cnt = r['assistant_cnt'] or 0

            user_preview = (first_user or "").replace("\n", " ").strip()
            ai_preview = (last_asst or "").replace("\n", " ").strip()
            import re as _re
            ai_preview_clean = _re.sub(r'\[(文件|PPT|教案|练习册|作业分析|报告)\s*[:：][^\]]{1,120}\]', ' ', ai_preview).strip()

            title = (r['title'] or '').strip()
            if not title or title == '新对话':
                seed = user_preview or ai_preview_clean or '空对话'
                title = seed[:42] + ('...' if len(seed) > 42 else '')

            result.append({
                "session_id": sid,
                "title": title,
                "created_at": r['created_at'] or "",
                "updated_at": r['last_msg_at'] or (r['created_at'] or ""),
                "user_count": user_cnt,
                "assistant_count": assistant_cnt,
                "total_count": total_cnt,
                "module_id": mod or "general",
                "files": all_files,
                "file_count": len(all_files),
                "user_preview": user_preview[:220],
                "ai_preview": ai_preview_clean[:260],
                "pk_id": pk,
            })
        return result
    finally:
        conn.close()

def delete_session_by_pk(user_id, pk_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'DELETE FROM user_sessions WHERE user_id = ? AND id = ?',
            (user_id, pk_id)
        )
        conn.commit()
    finally:
        conn.close()


def delete_all_user_sessions(user_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'DELETE FROM user_sessions WHERE user_id = ?',
            (user_id,)
        )
        conn.commit()
        affected = cursor.rowcount
        return affected
    finally:
        conn.close()


def clear_search_history(user_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'DELETE FROM search_history WHERE user_id = ?',
            (user_id,)
        )
        conn.commit()
        return cursor.rowcount
    finally:
        conn.close()


def get_user_stats():
    """获取用户统计信息"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT COUNT(*) FROM users')
        total = cursor.fetchone()[0]
        
        cursor.execute('''
            SELECT COUNT(DISTINCT user_id) FROM user_sessions 
            WHERE created_at >= datetime('now', '-7 days')
        ''')
        weekly_active = cursor.fetchone()[0]
        
        cursor.execute('''
            SELECT COUNT(*) FROM users 
            WHERE created_at >= datetime('now', '-7 days')
        ''')
        weekly_new = cursor.fetchone()[0]
        
        cursor.execute('''
            SELECT COUNT(DISTINCT s.user_id) FROM chat_messages m
            JOIN user_sessions s ON m.user_session_id = s.id
            WHERE m.created_at >= datetime('now', '-24 hours')
        ''')
        daily_active = cursor.fetchone()[0]
        
        return {
            "total_users": total,
            "weekly_active": weekly_active,
            "weekly_new": weekly_new,
            "daily_active": daily_active
        }
    finally:
        conn.close()


def get_system_stats():
    """获取系统统计信息"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT COUNT(*) FROM user_sessions')
        total_sessions = cursor.fetchone()[0]
        
        cursor.execute('SELECT COUNT(*) FROM chat_messages')
        total_messages = cursor.fetchone()[0]
        
        cursor.execute('''
            SELECT COUNT(*) FROM chat_messages 
            WHERE created_at >= datetime('now', '-7 days')
        ''')
        weekly_messages = cursor.fetchone()[0]
        
        return {
            "total_sessions": total_sessions,
            "total_messages": total_messages,
            "weekly_messages": weekly_messages
        }
    finally:
        conn.close()


def get_recent_activities(limit=10):
    """获取最近的用户活动"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute('''
            SELECT u.username, u.email, s.title, s.created_at, s.session_id
            FROM user_sessions s
            JOIN users u ON s.user_id = u.user_id
            ORDER BY s.created_at DESC
            LIMIT ?
        ''', (limit,))
        rows = cursor.fetchall()
        activities = []
        for row in rows:
            activities.append({
                "username": row['username'],
                "email": row['email'],
                "session_title": row['title'],
                "created_at": row['created_at'],
                "session_id": row['session_id']
            })
        return activities
    finally:
        conn.close()


def log_admin_action(admin_username, action, target_user_id=None, details=None, ip_address=None):
    """记录管理员操作日志"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'INSERT INTO admin_logs (admin_username, action, target_user_id, details, ip_address) VALUES (?, ?, ?, ?, ?)',
            (admin_username, action, target_user_id, details, ip_address)
        )
        conn.commit()
    except Exception as e:
        print(f"记录管理员日志失败: {e}")
    finally:
        conn.close()


def get_admin_logs(limit=50, offset=0, action_type='', target_user='', start_time='', end_time=''):
    """获取管理员操作日志，支持筛选"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        query = 'SELECT * FROM admin_logs WHERE 1=1'
        params = []

        if action_type:
            query += ' AND action = ?'
            params.append(action_type)
        if target_user:
            query += ' AND target_user_id = ?'
            params.append(target_user)
        if start_time:
            query += ' AND created_at >= ?'
            params.append(start_time)
        if end_time:
            query += ' AND created_at <= ?'
            params.append(end_time)

        query += ' ORDER BY created_at DESC LIMIT ? OFFSET ?'
        params.extend([limit, offset])

        cursor.execute(query, params)
        rows = cursor.fetchall()
        logs = []
        for row in rows:
            logs.append({
                "id": row['id'],
                "admin_username": row['admin_username'],
                "action": row['action'],
                "target_user_id": row['target_user_id'],
                "details": row['details'],
                "ip_address": row['ip_address'],
                "created_at": row['created_at']
            })
        return logs
    finally:
        conn.close()


def record_login_attempt(ip_address, username, success=False):
    """记录管理员登录尝试"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'INSERT INTO admin_login_attempts (ip_address, username, success) VALUES (?, ?, ?)',
            (ip_address, username, 1 if success else 0)
        )
        conn.commit()
        # 清理超过7天的旧记录
        cursor.execute(
            "DELETE FROM admin_login_attempts WHERE created_at < datetime('now', '-7 days')"
        )
        conn.commit()
    except Exception as e:
        print(f"记录登录尝试失败: {e}")
    finally:
        conn.close()


def get_recent_failed_logins(ip_address, minutes=15):
    """获取指定IP在最近几分钟内的失败登录次数"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        # H3: 参数化时间偏移，避免 f-string 拼接 SQL 造成的注入风险
        cursor.execute(
            "SELECT COUNT(*) FROM admin_login_attempts WHERE ip_address = ? AND success = 0 AND created_at >= datetime('now', ?)",
            (ip_address, f'-{int(minutes)} minutes')
        )
        return cursor.fetchone()[0]
    finally:
        conn.close()


def get_user_detail_with_sessions(user_id):
    """获取用户详情及会话统计"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT user_id, username, email, created_at FROM users WHERE user_id = ?',
            (user_id,)
        )
        row = cursor.fetchone()
        if not row:
            return None

        cursor.execute(
            'SELECT COUNT(*) FROM user_sessions WHERE user_id = ?',
            (user_id,)
        )
        session_count = cursor.fetchone()[0]

        cursor.execute(
            '''SELECT COUNT(*) FROM chat_messages m
               JOIN user_sessions s ON m.user_session_id = s.id
               WHERE s.user_id = ?''',
            (user_id,)
        )
        message_count = cursor.fetchone()[0]

        cursor.execute(
            '''SELECT session_id, title, created_at FROM user_sessions
               WHERE user_id = ? ORDER BY created_at DESC LIMIT 10''',
            (user_id,)
        )
        sessions = []
        for r in cursor.fetchall():
            sessions.append({
                "session_id": r['session_id'],
                "title": r['title'],
                "created_at": r['created_at']
            })

        return {
            "id": row['user_id'],
            "username": row['username'],
            "email": row['email'],
            "created_at": row['created_at'],
            "session_count": session_count,
            "message_count": message_count,
            "recent_sessions": sessions
        }
    finally:
        conn.close()


def get_all_local_users_paginated(page=1, per_page=20, sort_by='created_at', order='DESC', search=None):
    """分页获取本地用户，支持排序和搜索"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        allowed_sort = {'created_at': 'created_at', 'username': 'username', 'email': 'email'}
        sort_col = allowed_sort.get(sort_by, 'created_at')
        order = 'DESC' if order.upper() == 'DESC' else 'ASC'

        where_clause = ''
        params = []
        if search:
            # M4: 转义 LIKE 通配符 % 和 _，避免搜索词中包含这些字符时被当作通配符匹配
            escaped_search = search.replace('\\', '\\\\').replace('%', '\\%').replace('_', '\\_')
            where_clause = "WHERE username LIKE ? ESCAPE '\\' OR email LIKE ? ESCAPE '\\'"
            params = [f'%{escaped_search}%', f'%{escaped_search}%']

        cursor.execute(
            f'SELECT COUNT(*) FROM users {where_clause}',
            params
        )
        total = cursor.fetchone()[0]

        cursor.execute(
            f'SELECT user_id, username, email, created_at FROM users {where_clause} ORDER BY {sort_col} {order} LIMIT ? OFFSET ?',
            params + [per_page, (page - 1) * per_page]
        )
        rows = cursor.fetchall()
        users = []
        for row in rows:
            users.append({
                "id": row['user_id'],
                "username": row['username'],
                "email": row['email'],
                "created_at": row['created_at'],
                "role": "user",
                "last_sign_in_at": None
            })
        return {"total": total, "users": users, "page": page, "per_page": per_page}
    finally:
        conn.close()


def update_local_avatar(user_id, avatar_url):
    """更新本地用户头像URL（用于管理员账号等非Supabase用户）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'UPDATE admin_profile SET avatar_url = ?, updated_at = CURRENT_TIMESTAMP WHERE user_id = ?',
            (avatar_url, user_id)
        )
        if cursor.rowcount == 0:
            cursor.execute(
                'INSERT INTO admin_profile (user_id, avatar_url, username) VALUES (?, ?, ?)',
                (user_id, avatar_url, user_id)
            )
        conn.commit()
        return True
    except Exception as e:
        print(f"[database] 更新本地头像失败: {e}")
        return False
    finally:
        conn.close()


def get_local_avatar(user_id):
    """获取本地用户头像URL"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT avatar_url FROM admin_profile WHERE user_id = ?', (user_id,))
        row = cursor.fetchone()
        return row[0] if row else None
    except Exception as e:
        print(f"[database] 获取本地头像失败: {e}")
        return None
    finally:
        conn.close()


def get_local_admin_username(user_id):
    """获取管理员持久化的用户名（重启后仍可用）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT username FROM admin_profile WHERE user_id = ?', (user_id,))
        row = cursor.fetchone()
        return row[0] if row else None
    except Exception as e:
        print(f"[database] 获取管理员用户名失败: {e}")
        return None
    finally:
        conn.close()


def update_local_admin_username(user_id, new_username):
    """更新管理员持久化的用户名"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'UPDATE admin_profile SET username = ?, updated_at = CURRENT_TIMESTAMP WHERE user_id = ?',
            (new_username, user_id)
        )
        if cursor.rowcount == 0:
            cursor.execute(
                'INSERT INTO admin_profile (user_id, username) VALUES (?, ?)',
                (user_id, new_username)
            )
        conn.commit()
        return True
    except Exception as e:
        print(f"[database] 更新管理员用户名失败: {e}")
        return False
    finally:
        conn.close()


# M5: 模块导入时初始化数据库，用 try/except 包裹防止初始化失败导致整个模块导入失败
# 如果应用启动依赖该初始化，调用方应在启动时显式调用 init_db() 以获取错误信息
try:
    init_db()
except Exception as _init_err:
    print(f"[database] 模块导入时 init_db 失败: {_init_err}")