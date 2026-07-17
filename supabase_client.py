import re as _re_module
import threading
import time
from typing import Optional

# L3: 将 create_client 导入统一放到文件顶部，便于依赖管理与可读性
from supabase import create_client

def _validate_email(email: str) -> bool:
    if not email or not isinstance(email, str):
        return False
    return bool(_re_module.match(r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$', email))

# L2: 常见弱口令黑名单，拦截高频弱密码
_WEAK_PASSWORDS = {'password1', '12345678', 'abc12345', 'qwerty12', 'iloveyou1', 'admin1234'}

def _validate_password_strength(password: str) -> str:
    if not password or len(password) < 8:
        return "密码长度至少8位"
    if not _re_module.search(r'[A-Za-z]', password) or not _re_module.search(r'[0-9]', password):
        return "密码需包含字母和数字"
    # L2: 拦截常见弱口令
    if password.lower() in _WEAK_PASSWORDS:
        return "密码过于简单，请更换"
    return ""

def _safe_error(e: Exception, fallback: str = "操作失败") -> str:
    err_str = str(e).lower()
    if 'rate limit' in err_str or '429' in err_str:
        return "请求过于频繁，请稍后重试"
    if 'invalid credentials' in err_str or 'invalid login' in err_str:
        return "用户名或密码错误"
    # M4: 注册失败统一返回模糊提示，防止邮箱枚举攻击
    if 'email already' in err_str or 'already registered' in err_str:
        return "注册失败，请检查邮箱或稍后重试"
    if 'email not confirmed' in err_str:
        return "邮箱未验证，请先查看验证邮件"
    return fallback


_supabase: Optional[object] = None
_supabase_admin: Optional[object] = None

# M7: 初始化锁，防止并发场景下重复创建全局单例客户端
_init_lock = threading.Lock()

def init_supabase(url: str, anon_key: str, service_role_key: str = None):
    global _supabase, _supabase_admin
    with _init_lock:
        # M7: 双重检查，避免并发时重复创建客户端
        if _supabase is not None and _supabase_admin is not None:
            return _supabase
        # H3: service_role_key 是高危凭证，仅在内存中持有，切勿打印或写入日志/异常信息
        _supabase = create_client(url, anon_key)
        if service_role_key:
            _supabase_admin = create_client(url, service_role_key)
        return _supabase

def get_supabase():
    if _supabase is None:
        raise RuntimeError("Supabase client not initialized. Call init_supabase() first.")
    return _supabase

def get_supabase_admin():
    # H3: admin 客户端持有 service_role_key，调用方切勿泄露其内部状态或返回值
    if _supabase_admin is not None:
        return _supabase_admin
    return None

def sign_up(email: str, password: str, username: str = None):
    if not _validate_email(email):
        return {"success": False, "error": "请输入有效的邮箱地址"}
    pwd_err = _validate_password_strength(password)
    if pwd_err:
        return {"success": False, "error": pwd_err}
    supabase = get_supabase()
    try:
        data = {"email": email, "password": password}
        if username:
            data["options"] = {"data": {"username": username}}
        result = supabase.auth.sign_up(data)
        if result.user:
            return {
                "success": True,
                "user": {
                    "id": result.user.id,
                    "email": result.user.email,
                    "username": result.user.user_metadata.get("username") if result.user.user_metadata else None
                }
            }
        return {"success": False, "error": "注册失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "注册失败，请稍后重试")}

def sign_in(email: str, password: str):
    supabase = get_supabase()
    try:
        result = supabase.auth.sign_in_with_password({"email": email, "password": password})
        if result.user and result.session:
            return {
                "success": True,
                "user": {
                    "id": result.user.id,
                    "email": result.user.email,
                    "username": result.user.user_metadata.get("username") if result.user.user_metadata else None
                },
                "session": {
                    "access_token": result.session.access_token,
                    "refresh_token": result.session.refresh_token,
                    "expires_at": result.session.expires_at
                }
            }
        return {"success": False, "error": "登录失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "登录失败")}

def sign_out():
    supabase = get_supabase()
    try:
        supabase.auth.sign_out()
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "操作失败")}

def get_user(access_token: str):
    supabase = get_supabase()
    try:
        result = supabase.auth.get_user(access_token)
        if result.user:
            return {
                "success": True,
                "user": {
                    "id": result.user.id,
                    "email": result.user.email,
                    "username": result.user.user_metadata.get("username") if result.user.user_metadata else None
                }
            }
        return {"success": False, "error": "用户不存在"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "获取用户信息失败")}

def update_user_metadata(user_id: str, metadata: dict):
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        # H2: 先读取现有 user_metadata，合并后再写入，避免整体覆盖丢失字段
        user_resp = supabase.auth.admin.get_user_by_id(user_id)
        if not user_resp or not user_resp.user:
            return {"success": False, "error": "用户不存在"}
        existing = user_resp.user.user_metadata or {}
        # 合并：新 metadata 覆盖旧的同名字段，其余字段保留
        merged = {**existing, **metadata}
        result = supabase.auth.admin.update_user_by_id(user_id, {"data": merged})
        if result.user:
            return {"success": True, "user": {"id": result.user.id}}
        return {"success": False, "error": "更新失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "更新失败")}

def get_all_users(page: int = 1, per_page: int = 1000):
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置，请联系管理员"}
    try:
        # M5: 支持分页参数，向后兼容（不传参数时默认拉取前 1000 条）
        result = supabase.auth.admin.list_users(page=page, per_page=per_page)
        users = []
        user_list = result if isinstance(result, list) else getattr(result, 'users', [])
        for user in user_list:
            users.append({
                "id": user.id,
                "email": user.email,
                "username": user.user_metadata.get("username") if user.user_metadata else None,
                "created_at": user.created_at,
                "last_sign_in_at": user.last_sign_in_at,
                "role": user.role
            })
        return {"success": True, "users": users}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "获取用户列表失败")}

def delete_user(user_id: str):
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        result = supabase.auth.admin.delete_user(user_id)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "删除用户失败")}

def create_admin_user(email: str, password: str, username: str = None):
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    pwd_err = _validate_password_strength(password)
    if pwd_err:
        return {"success": False, "error": pwd_err}
    if not _validate_email(email):
        return {"success": False, "error": "请输入有效的邮箱地址"}
    try:
        result = supabase.auth.admin.create_user({
            "email": email,
            "password": password,
            "email_confirm": True,
            "user_metadata": {"username": username} if username else {}
        })
        if result.user:
            return {
                "success": True,
                "user": {
                    "id": result.user.id,
                    "email": result.user.email,
                    "username": result.user.user_metadata.get("username") if result.user.user_metadata else None
                }
            }
        return {"success": False, "error": "创建失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "创建用户失败")}

def get_email_by_username(username: str):
    # TODO M6: 当前为 O(n) 全表扫描，用户量增大后存在性能问题。
    # 建议未来用独立 profiles 表 + username 索引替代。
    supabase = get_supabase_admin()
    if supabase is not None:
        try:
            result = supabase.auth.admin.list_users()
            user_list = result if isinstance(result, list) else getattr(result, 'users', [])
            for user in user_list:
                user_username = user.user_metadata.get("username") if user.user_metadata else None
                if user_username and user_username.lower() == username.lower():
                    return {"success": True, "email": user.email}
            return {"success": False, "error": "用户名不存在"}
        except Exception as e:
            return {"success": False, "error": _safe_error(e, "查找失败")}
    else:
        return {"success": False, "error": "服务未配置"}

def username_exists(username: str):
    # TODO M6: 当前为 O(n) 全表扫描，用户量增大后存在性能问题。
    # 建议未来用独立 profiles 表 + username 唯一索引替代。
    supabase = get_supabase_admin()
    if supabase is not None:
        try:
            result = supabase.auth.admin.list_users()
            user_list = result if isinstance(result, list) else getattr(result, 'users', [])
            for user in user_list:
                user_username = user.user_metadata.get("username") if user.user_metadata else None
                if user_username and user_username.lower() == username.lower():
                    return {"success": True, "exists": True, "user_id": user.id}
            return {"success": True, "exists": False}
        except Exception as e:
            return {"success": False, "error": _safe_error(e, "查找失败")}
    else:
        return {"success": False, "error": "服务未配置"}

# M1: 邮件发送内存速率限制（按邮箱维度，每小时最多 3 次），防止滥用与邮件轰炸
_email_send_log: dict = {}  # email -> [timestamp, ...]

def _check_email_rate_limit(email: str, max_per_hour: int = 3) -> bool:
    now = time.time()
    window = 3600
    logs = _email_send_log.get(email, [])
    # 清理超出 1 小时窗口的过期记录
    logs = [t for t in logs if now - t < window]
    if len(logs) >= max_per_hour:
        _email_send_log[email] = logs
        return False
    logs.append(now)
    _email_send_log[email] = logs
    return True

def resend_verification_email(email: str):
    # M1: 速率限制
    if not _check_email_rate_limit(email):
        return {"success": False, "error": "请求过于频繁，请稍后再试"}
    supabase = get_supabase()
    try:
        result = supabase.auth.resend({"type": "signup", "email": email})
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "发送失败")}

def send_password_reset_email(email: str):
    # M1: 速率限制
    if not _check_email_rate_limit(email):
        return {"success": False, "error": "请求过于频繁，请稍后再试"}
    supabase = get_supabase()
    try:
        result = supabase.auth.reset_password_email(email)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "发送失败")}

def update_username(user_id: str, new_username: str):
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        result = supabase.auth.admin.update_user_by_id(user_id, {"data": {"username": new_username}})
        if result.user:
            return {"success": True, "username": new_username}
        return {"success": False, "error": "更新失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "更新失败")}

# H1: set_session 会污染全局单例的 session 状态，加锁保证并发安全
_session_lock = threading.Lock()

def update_user_password(access_token: str, new_password: str):
    # TODO M3: 应验证旧密码或要求 recent login，需产品决策后补充
    pwd_err = _validate_password_strength(new_password)
    if pwd_err:
        return {"success": False, "error": pwd_err}
    supabase = get_supabase()
    # H1: 临界区内串行执行 set_session/update_user，避免并发请求互相覆盖 session 导致越权
    with _session_lock:
        try:
            supabase.auth.set_session(access_token)
            result = supabase.auth.update_user({"password": new_password})
            if result.user:
                return {"success": True}
            return {"success": False, "error": "更新失败"}
        except Exception as e:
            return {"success": False, "error": _safe_error(e, "密码更新失败")}

def update_user_avatar(user_id: str, avatar_url: str):
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        result = supabase.auth.admin.update_user_by_id(user_id, {"data": {"avatar_url": avatar_url}})
        if result.user:
            return {"success": True, "avatar_url": avatar_url}
        return {"success": False, "error": "更新失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "更新失败")}

def get_user_by_id(user_id: str):
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        result = supabase.auth.admin.get_user_by_id(user_id)
        if result.user:
            user = result.user
            return {
                "success": True,
                "user": {
                    "id": user.id,
                    "email": user.email,
                    "username": user.user_metadata.get("username") if user.user_metadata else None,
                    "avatar_url": user.user_metadata.get("avatar_url") if user.user_metadata else None,
                    "created_at": user.created_at,
                    "last_sign_in_at": user.last_sign_in_at,
                    "role": user.role,
                    "email_confirmed_at": user.email_confirmed_at
                }
            }
        return {"success": False, "error": "用户不存在"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "获取用户失败")}


def update_user_email(user_id: str, new_email: str):
    if not _validate_email(new_email):
        return {"success": False, "error": "请输入有效的邮箱地址"}
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        # M2: 强制 email_confirm=False，并触发验证邮件，确保新邮箱经过验证
        result = supabase.auth.admin.update_user_by_id(user_id, {
            "email": new_email,
            "email_confirm": False
        })
        if result.user:
            # 触发新邮箱的验证邮件（失败不阻塞主流程）
            try:
                get_supabase().auth.resend({"email": new_email, "type": "signup"})
            except Exception:
                pass
            return {"success": True, "email": result.user.email}
        return {"success": False, "error": "更新失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "更新失败")}


def update_user_role(user_id: str, role: str):
    valid_roles = {"authenticated", "admin"}
    if role not in valid_roles:
        return {"success": False, "error": "无效的角色"}
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        result = supabase.auth.admin.update_user_by_id(user_id, {"role": role})
        if result.user:
            return {"success": True, "role": result.user.role}
        return {"success": False, "error": "更新失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "更新失败")}


def admin_update_user_metadata(user_id: str, metadata: dict):
    # L1: 复用 update_user_metadata 的合并逻辑，避免重复实现导致行为不一致
    return update_user_metadata(user_id, metadata)


def admin_create_user(email: str, password: str, username: str = None, role: str = "authenticated"):
    pwd_err = _validate_password_strength(password)
    if pwd_err:
        return {"success": False, "error": pwd_err}
    if not _validate_email(email):
        return {"success": False, "error": "请输入有效的邮箱地址"}
    valid_roles = {"authenticated", "admin"}
    if role not in valid_roles:
        role = "authenticated"
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        result = supabase.auth.admin.create_user({
            "email": email,
            "password": password,
            "email_confirm": True,
            "role": role,
            "user_metadata": {"username": username} if username else {}
        })
        if result.user:
            return {
                "success": True,
                "user": {
                    "id": result.user.id,
                    "email": result.user.email,
                    "username": result.user.user_metadata.get("username") if result.user.user_metadata else None,
                    "role": result.user.role,
                    "created_at": result.user.created_at
                }
            }
        return {"success": False, "error": "创建失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "创建用户失败")}


def admin_update_user_password(user_id: str, new_password: str):
    pwd_err = _validate_password_strength(new_password)
    if pwd_err:
        return {"success": False, "error": pwd_err}
    supabase = get_supabase_admin()
    if supabase is None:
        return {"success": False, "error": "服务未配置"}
    try:
        result = supabase.auth.admin.update_user_by_id(user_id, {"password": new_password})
        if result.user:
            return {"success": True}
        return {"success": False, "error": "更新失败"}
    except Exception as e:
        return {"success": False, "error": _safe_error(e, "密码更新失败")}
