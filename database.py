import sqlite3
import os
import hashlib
import uuid
from datetime import datetime
from pathlib import Path

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
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            email TEXT UNIQUE,
            session_token TEXT UNIQUE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            last_login TIMESTAMP
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS user_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            session_id TEXT UNIQUE NOT NULL,
            title TEXT DEFAULT '新对话',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
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
            user_id INTEGER NOT NULL,
            query TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
        )
    ''')
    
    conn.commit()
    conn.close()

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def register_user(username, password, email=None):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'INSERT INTO users (username, password_hash, email) VALUES (?, ?, ?)',
            (username, hash_password(password), email)
        )
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def authenticate_user(username, password):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'SELECT id, username FROM users WHERE username = ? AND password_hash = ?',
        (username, hash_password(password))
    )
    user = cursor.fetchone()
    conn.close()
    if user:
        return {'id': user[0], 'username': user[1]}
    return None

def generate_session_token():
    return str(uuid.uuid4())

def set_session_token(user_id, token):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'UPDATE users SET session_token = ?, last_login = ? WHERE id = ?',
        (token, datetime.now(), user_id)
    )
    conn.commit()
    conn.close()

def get_user_by_token(token):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'SELECT id, username FROM users WHERE session_token = ?',
        (token,)
    )
    user = cursor.fetchone()
    conn.close()
    if user:
        return {'id': user[0], 'username': user[1]}
    return None

def clear_session_token(token):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('UPDATE users SET session_token = NULL WHERE session_token = ?', (token,))
    conn.commit()
    conn.close()

def save_user_chat_session(user_id, session_id, title='新对话'):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute(
            'INSERT OR REPLACE INTO user_sessions (user_id, session_id, title) VALUES (?, ?, ?)',
            (user_id, session_id, title)
        )
        conn.commit()
        return cursor.lastrowid
    finally:
        conn.close()

def get_user_sessions(user_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
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
    conn.close()
    return sessions

def save_chat_message(user_session_id, role, content):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'INSERT INTO chat_messages (user_session_id, role, content) VALUES (?, ?, ?)',
        (user_session_id, role, content)
    )
    conn.commit()
    conn.close()

def get_chat_messages(user_session_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
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
    conn.close()
    return messages

def save_search_history(user_id, query):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'INSERT INTO search_history (user_id, query) VALUES (?, ?)',
        (user_id, query)
    )
    conn.commit()
    conn.close()

def get_search_history(user_id, limit=50):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
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
    conn.close()
    return history

def delete_user_session(user_id, session_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'DELETE FROM user_sessions WHERE user_id = ? AND session_id = ?',
        (user_id, session_id)
    )
    conn.commit()
    conn.close()

def update_session_title(user_id, session_id, title):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'UPDATE user_sessions SET title = ? WHERE user_id = ? AND session_id = ?',
        (title, user_id, session_id)
    )
    conn.commit()
    conn.close()

def get_user_session_pk(user_id, session_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'SELECT id FROM user_sessions WHERE user_id = ? AND session_id = ? LIMIT 1',
        (user_id, session_id)
    )
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else None

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
        # 取最早一条用户消息 + 最晚一条 AI/assistant 消息
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

        # 从标题/首条用户消息/最后 AI 消息推断模块
        mod = None
        for txt in (r['title'] or "", first_user, last_asst):
            mod = _guess_module_id_from_text(txt)
            if mod:
                break

        # 收集所有 AI 回复中的文件徽章
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

        # 预览：优先用户首条需求摘要 + AI 最后一条摘要
        user_preview = (first_user or "").replace("\n", " ").strip()
        ai_preview = (last_asst or "").replace("\n", " ").strip()
        # 去掉 [文件:xxx.md] 标记让预览更干净
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
    conn.close()
    return result

def delete_session_by_pk(user_id, pk_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(
        'DELETE FROM user_sessions WHERE user_id = ? AND id = ?',
        (user_id, pk_id)
    )
    conn.commit()
    conn.close()

init_db()