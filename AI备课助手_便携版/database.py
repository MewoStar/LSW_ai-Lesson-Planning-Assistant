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
    # 字段说明：
    #   - username: 旧字段，曾用作"持久化登录账号"，已弃用（仅保留向后兼容）
    #   - display_name: 显示名（个人中心可修改），与登录账号严格分离
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_profile (
            user_id TEXT PRIMARY KEY,
            username TEXT,
            avatar_url TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # 迁移：为旧库新增 display_name 列
    cursor.execute("PRAGMA table_info(admin_profile)")
    admin_cols = [row[1] for row in cursor.fetchall()]
    if 'display_name' not in admin_cols:
        cursor.execute("ALTER TABLE admin_profile ADD COLUMN display_name TEXT")
        # 将旧 username 字段中的数据迁移到 display_name，避免历史数据丢失
        cursor.execute("UPDATE admin_profile SET display_name = username WHERE display_name IS NULL AND username IS NOT NULL")
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS session_tokens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT NOT NULL,
            token TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_session_tokens_token ON session_tokens(token)')

    # ===== 管理员账号与会话 =====
    # admin_users: 管理员账号主表（替代旧 ADMIN_CREDENTIALS 内存字典）
    #   - username: 登录账号（不可改）
    #   - password_hash + salt: 每用户独立盐的 pbkdf2_hmac 密码哈希
    #   - display_name: 个人中心可改的显示名（与登录账号严格分离）
    #   - avatar_url: 个人中心头像
    #   - is_super: 1=超管（可管理其他管理员），0=普通管理员
    #   - is_active: 1=可登录，0=被禁用
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_users (
            id TEXT PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            display_name TEXT,
            avatar_url TEXT,
            is_super INTEGER DEFAULT 0,
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # admin_sessions: 管理员登录会话（服务端存储，替代旧 admin_token cookie 比对）
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admin_sessions (
            token TEXT PRIMARY KEY,
            admin_id TEXT NOT NULL,
            ip_address TEXT,
            user_agent TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            expires_at TIMESTAMP NOT NULL
        )
    ''')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_admin_sessions_admin_id ON admin_sessions(admin_id)')

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

def save_search_history(user_id, query, access_token=None):
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

    if access_token:
        try:
            supabase_client.sync_search_history(access_token, user_id, query)
        except Exception:
            pass

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
                'INSERT INTO admin_profile (user_id, avatar_url) VALUES (?, ?)',
                (user_id, avatar_url)
            )
        conn.commit()
        return True
    except sqlite3.Error:
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
    except sqlite3.Error:
        return None
    finally:
        conn.close()


def get_local_admin_display_name(user_id):
    """获取管理员显示名（个人中心可修改，与登录账号严格分离）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT display_name FROM admin_profile WHERE user_id = ?', (user_id,))
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error:
        return None
    finally:
        conn.close()


def update_local_admin_display_name(user_id, new_display_name):
    """更新管理员显示名（不影响登录账号）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'UPDATE admin_profile SET display_name = ?, updated_at = CURRENT_TIMESTAMP WHERE user_id = ?',
            (new_display_name, user_id)
        )
        if cursor.rowcount == 0:
            cursor.execute(
                'INSERT INTO admin_profile (user_id, display_name) VALUES (?, ?)',
                (user_id, new_display_name)
            )
        conn.commit()
        return True
    except sqlite3.Error:
        return False
    finally:
        conn.close()


# M5: 模块导入时初始化数据库，用 try/except 包裹防止初始化失败导致整个模块导入失败
# 如果应用启动依赖该初始化，调用方应在启动时显式调用 init_db() 以获取错误信息
try:
    init_db()
except Exception as _init_err:
    print(f"[database] 模块导入时 init_db 失败: {_init_err}")


# ===== 管理员账号与会话管理 =====
# 设计要点：
#   - 密码哈希：每用户独立盐，pbkdf2_hmac sha256，100000 次迭代（与现有 _hash_password 强度一致）
#   - 会话：登录成功生成随机 token，服务端持久化，cookie 仅携带 token；不再依赖进程级常量
#   - 超管：is_super=1 的管理员可创建/删除/禁用其他管理员

import hashlib as _hashlib
import secrets as _secrets
import time as _time

ADMIN_SESSION_TTL = 8 * 3600  # 管理员会话有效期 8 小时


def _hash_password_with_salt(password: str, salt: bytes) -> str:
    """使用给定盐对密码做 pbkdf2_hmac 哈希"""
    return _hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 100000).hex()


def create_admin_user(username: str, password: str, display_name=None, is_super: bool = False) -> dict:
    """创建管理员账号。
    返回 {"success": bool, "admin": {...}} 或 {"success": False, "error": str}
    """
    if not username or not password:
        return {"success": False, "error": "用户名和密码不能为空"}
    if len(password) < 6:
        return {"success": False, "error": "密码长度至少6位"}
    if len(username) < 2 or len(username) > 20:
        return {"success": False, "error": "用户名长度需在2-20位之间"}
    if not re.match(r'^[a-zA-Z0-9_\u4e00-\u9fa5]+$', username):
        return {"success": False, "error": "用户名只允许字母、数字、下划线和中文"}

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        # 检查用户名是否已存在
        cursor.execute('SELECT id FROM admin_users WHERE username = ?', (username,))
        if cursor.fetchone():
            return {"success": False, "error": "该管理员用户名已存在"}

        admin_id = f"admin_{uuid.uuid4().hex[:12]}"
        salt = _secrets.token_bytes(32)
        password_hash = _hash_password_with_salt(password, salt)
        cursor.execute(
            'INSERT INTO admin_users (id, username, password_hash, salt, display_name, is_super, is_active) '
            'VALUES (?, ?, ?, ?, ?, ?, 1)',
            (admin_id, username, password_hash, salt.hex(), display_name or username, 1 if is_super else 0)
        )
        conn.commit()
        return {
            "success": True,
            "admin": {
                "id": admin_id, "username": username,
                "display_name": display_name or username,
                "is_super": is_super, "is_active": True
            }
        }
    except sqlite3.IntegrityError:
        return {"success": False, "error": "该管理员用户名已存在"}
    except sqlite3.Error as e:
        return {"success": False, "error": f"创建失败: {e}"}
    finally:
        conn.close()


def get_admin_by_username(username: str) -> dict | None:
    """按登录账号查询管理员完整记录（含密码哈希）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT id, username, password_hash, salt, display_name, avatar_url, '
            'is_super, is_active, created_at FROM admin_users WHERE username = ?',
            (username,)
        )
        row = cursor.fetchone()
        if not row:
            return None
        return {
            "id": row[0], "username": row[1], "password_hash": row[2],
            "salt": row[3], "display_name": row[4], "avatar_url": row[5],
            "is_super": bool(row[6]), "is_active": bool(row[7]), "created_at": row[8]
        }
    except sqlite3.Error:
        return None
    finally:
        conn.close()


def get_admin_by_id(admin_id: str) -> dict | None:
    """按 ID 查询管理员（不含密码哈希）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT id, username, display_name, avatar_url, is_super, is_active, created_at '
            'FROM admin_users WHERE id = ?',
            (admin_id,)
        )
        row = cursor.fetchone()
        if not row:
            return None
        return {
            "id": row[0], "username": row[1], "display_name": row[2],
            "avatar_url": row[3], "is_super": bool(row[4]),
            "is_active": bool(row[5]), "created_at": row[6]
        }
    except sqlite3.Error:
        return None
    finally:
        conn.close()


def verify_admin_password(username: str, password: str) -> dict:
    """验证管理员密码。
    返回 {"success": bool, "admin": {...}} 或 {"success": False, "error": str, "reason": "not_found"|"inactive"|"wrong_password"}
    """
    admin = get_admin_by_username(username)
    if not admin:
        return {"success": False, "error": "用户名或密码错误", "reason": "not_found"}
    if not admin["is_active"]:
        return {"success": False, "error": "账号已被禁用", "reason": "inactive"}
    try:
        salt = bytes.fromhex(admin["salt"])
    except ValueError:
        return {"success": False, "error": "用户名或密码错误", "reason": "corrupt"}
    computed = _hash_password_with_salt(password, salt)
    import hmac as _hmac
    if _hmac.compare_digest(computed, admin["password_hash"]):
        return {"success": True, "admin": admin}
    return {"success": False, "error": "用户名或密码错误", "reason": "wrong_password"}


def update_admin_password(admin_id: str, new_password: str) -> dict:
    """更新管理员密码（每用户重新生成盐）"""
    if not new_password or len(new_password) < 6:
        return {"success": False, "error": "密码长度至少6位"}
    salt = _secrets.token_bytes(32)
    password_hash = _hash_password_with_salt(new_password, salt)
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'UPDATE admin_users SET password_hash = ?, salt = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
            (password_hash, salt.hex(), admin_id)
        )
        if cursor.rowcount == 0:
            return {"success": False, "error": "管理员不存在"}
        conn.commit()
        # 改密后吊销该管理员所有会话，强制重新登录
        cursor.execute('DELETE FROM admin_sessions WHERE admin_id = ?', (admin_id,))
        conn.commit()
        return {"success": True}
    except sqlite3.Error as e:
        return {"success": False, "error": f"更新失败: {e}"}
    finally:
        conn.close()


def list_admin_users() -> list:
    """列出所有管理员（不含密码哈希）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT id, username, display_name, avatar_url, is_super, is_active, created_at '
            'FROM admin_users ORDER BY is_super DESC, created_at ASC'
        )
        rows = cursor.fetchall()
        admins = []
        for row in rows:
            admins.append({
                "id": row[0], "username": row[1], "display_name": row[2],
                "avatar_url": row[3], "is_super": bool(row[4]),
                "is_active": bool(row[5]), "created_at": row[6]
            })
        return admins
    except sqlite3.Error:
        return []
    finally:
        conn.close()


def delete_admin_user(admin_id: str) -> dict:
    """删除管理员账号（同时清理其会话）。不能删除最后一个超管。"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        # 防止删除最后一个超管导致无超管可管理
        cursor.execute('SELECT COUNT(*) FROM admin_users WHERE is_super = 1')
        super_count = cursor.fetchone()[0]
        cursor.execute('SELECT is_super FROM admin_users WHERE id = ?', (admin_id,))
        row = cursor.fetchone()
        if not row:
            return {"success": False, "error": "管理员不存在"}
        if row[0] == 1 and super_count <= 1:
            return {"success": False, "error": "不能删除最后一个超级管理员"}

        cursor.execute('DELETE FROM admin_sessions WHERE admin_id = ?', (admin_id,))
        cursor.execute('DELETE FROM admin_users WHERE id = ?', (admin_id,))
        conn.commit()
        return {"success": True}
    except sqlite3.Error as e:
        return {"success": False, "error": f"删除失败: {e}"}
    finally:
        conn.close()


def set_admin_active(admin_id: str, is_active: bool) -> dict:
    """启用/禁用管理员账号。禁用时同步吊销其所有会话。"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('UPDATE admin_users SET is_active = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                       (1 if is_active else 0, admin_id))
        if cursor.rowcount == 0:
            return {"success": False, "error": "管理员不存在"}
        if not is_active:
            cursor.execute('DELETE FROM admin_sessions WHERE admin_id = ?', (admin_id,))
        conn.commit()
        return {"success": True}
    except sqlite3.Error as e:
        return {"success": False, "error": f"操作失败: {e}"}
    finally:
        conn.close()


# ----- 管理员会话 -----

def create_admin_session(admin_id: str, ip_address: str = None, user_agent: str = None) -> str:
    """创建管理员会话，返回 session token。同时清理该用户过期旧会话。"""
    token = _secrets.token_hex(32)
    expires_at = _time.time() + ADMIN_SESSION_TTL
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        # 同一管理员最多保留 5 个活跃会话，超出删除最早的
        cursor.execute(
            'SELECT token FROM admin_sessions WHERE admin_id = ? AND expires_at > ? '
            'ORDER BY created_at ASC',
            (admin_id, _time.time())
        )
        existing = cursor.fetchall()
        if len(existing) >= 5:
            for old_token in existing[:len(existing) - 4]:
                cursor.execute('DELETE FROM admin_sessions WHERE token = ?', (old_token[0],))
        # 插入新会话；expires_at 以 Unix 时间戳存储，便于直接和 time.time() 比较
        cursor.execute(
            'INSERT INTO admin_sessions (token, admin_id, ip_address, user_agent, expires_at) '
            'VALUES (?, ?, ?, ?, ?)',
            (token, admin_id, ip_address, user_agent, expires_at)
        )
        conn.commit()
        return token
    except sqlite3.Error:
        return None
    finally:
        conn.close()


def get_admin_session(token: str) -> dict | None:
    """按 token 查询会话。若已过期则删除并返回 None。
    返回 {"admin_id", "username", "display_name", "is_super", "is_active", "expires_at", "ip_address"} 或 None
    """
    if not token:
        return None
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'SELECT s.admin_id, s.ip_address, s.expires_at, '
            'a.username, a.display_name, a.is_super, a.is_active '
            'FROM admin_sessions s JOIN admin_users a ON s.admin_id = a.id '
            'WHERE s.token = ?',
            (token,)
        )
        row = cursor.fetchone()
        if not row:
            return None
        expires_at = row[2]
        if _time.time() > expires_at:
            cursor.execute('DELETE FROM admin_sessions WHERE token = ?', (token,))
            conn.commit()
            return None
        # 管理员被禁用后，会话立即失效
        if not row[6]:
            cursor.execute('DELETE FROM admin_sessions WHERE token = ?', (token,))
            conn.commit()
            return None
        return {
            "admin_id": row[0], "ip_address": row[1], "expires_at": expires_at,
            "username": row[3], "display_name": row[4],
            "is_super": bool(row[5]), "is_active": True
        }
    except sqlite3.Error:
        return None
    finally:
        conn.close()


def delete_admin_session(token: str) -> bool:
    """删除指定会话（登出时调用）"""
    if not token:
        return False
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('DELETE FROM admin_sessions WHERE token = ?', (token,))
        conn.commit()
        return cursor.rowcount > 0
    except sqlite3.Error:
        return False
    finally:
        conn.close()


def delete_all_admin_sessions(admin_id: str) -> int:
    """吊销指定管理员所有会话"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('DELETE FROM admin_sessions WHERE admin_id = ?', (admin_id,))
        conn.commit()
        return cursor.rowcount
    except sqlite3.Error:
        return 0
    finally:
        conn.close()


def cleanup_expired_admin_sessions() -> int:
    """清理所有过期会话，返回删除条数"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('DELETE FROM admin_sessions WHERE expires_at < ?', (_time.time(),))
        deleted = cursor.rowcount
        conn.commit()
        return deleted
    except sqlite3.Error:
        return 0
    finally:
        conn.close()


# ----- 管理员显示名与头像（替代旧 admin_profile 表的 'admin' 字面量路径） -----

def update_admin_display_name(admin_id: str, new_display_name: str) -> bool:
    """更新管理员显示名（与登录账号严格分离）"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'UPDATE admin_users SET display_name = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
            (new_display_name, admin_id)
        )
        conn.commit()
        return cursor.rowcount > 0
    except sqlite3.Error:
        return False
    finally:
        conn.close()


def get_admin_display_name(admin_id: str) -> str | None:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT display_name FROM admin_users WHERE id = ?', (admin_id,))
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error:
        return None
    finally:
        conn.close()


def update_admin_avatar(admin_id: str, avatar_url: str) -> bool:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'UPDATE admin_users SET avatar_url = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
            (avatar_url, admin_id)
        )
        conn.commit()
        return cursor.rowcount > 0
    except sqlite3.Error:
        return False
    finally:
        conn.close()


def get_admin_avatar(admin_id: str) -> str | None:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT avatar_url FROM admin_users WHERE id = ?', (admin_id,))
        row = cursor.fetchone()
        return row[0] if row else None
    except sqlite3.Error:
        return None
    finally:
        conn.close()


def init_default_super_admin(env_username: str, env_password: str) -> dict:
    """启动时确保至少有一个超管。
    - 若 admin_users 表已有任意超管，则不做任何事
    - 若表为空（无任何管理员），则用环境变量创建初始超管
    - 若有普通管理员但无超管，则升级最早创建的普通管理员为超管（异常恢复场景）
    返回 {"created": bool, "admin": {...}|None, "message": str}
    """
    if not env_username or not env_password:
        return {"created": False, "admin": None, "message": "环境变量未配置管理员账号"}

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('SELECT COUNT(*) FROM admin_users WHERE is_super = 1')
        super_count = cursor.fetchone()[0]
        if super_count > 0:
            return {"created": False, "admin": None, "message": "已存在超管，跳过初始化"}

        cursor.execute('SELECT COUNT(*) FROM admin_users')
        total = cursor.fetchone()[0]
        if total == 0:
            conn.close()
            result = create_admin_user(env_username, env_password, display_name=env_username, is_super=True)
            if result.get("success"):
                return {"created": True, "admin": result["admin"], "message": "已创建初始超管"}
            return {"created": False, "admin": None, "message": result.get("error", "创建失败")}
        else:
            # 有管理员但无超管，升级最早的一个
            cursor.execute('SELECT id FROM admin_users ORDER BY created_at ASC LIMIT 1')
            row = cursor.fetchone()
            if row:
                cursor.execute('UPDATE admin_users SET is_super = 1 WHERE id = ?', (row[0],))
                conn.commit()
                admin = get_admin_by_id(row[0])
                return {"created": False, "admin": admin, "message": "已将最早的管理员升级为超管"}
            return {"created": False, "admin": None, "message": "无管理员可升级"}
    except sqlite3.Error as e:
        return {"created": False, "admin": None, "message": f"初始化失败: {e}"}
    finally:
        if conn:
            conn.close()