from flask import request, jsonify
from app.supabase_client import get_supabase_client
from app.config import Config
import jwt
from datetime import datetime, timedelta
from functools import wraps

def register_user(email: str, password: str, username: str = None):
    supabase = get_supabase_client()
    try:
        response = supabase.auth.sign_up({
            "email": email,
            "password": password,
            "options": {
                "data": {
                    "username": username or email.split("@")[0]
                }
            }
        })
        if response.user:
            custom_token = generate_custom_jwt(response.user.id, response.user.email, username)
            return {
                "status": "success",
                "user_id": response.user.id,
                "email": response.user.email,
                "username": username or email.split("@")[0],
                "access_token": custom_token,
                "supabase_token": response.session.access_token if response.session else None
            }
        return {"status": "error", "message": "注册失败"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

def login_user(email: str, password: str):
    supabase = get_supabase_client()
    try:
        response = supabase.auth.sign_in_with_password({
            "email": email,
            "password": password
        })
        if response.user:
            username = response.user.user_metadata.get("username", email.split("@")[0])
            custom_token = generate_custom_jwt(response.user.id, response.user.email, username)
            return {
                "status": "success",
                "user_id": response.user.id,
                "email": response.user.email,
                "username": username,
                "access_token": custom_token,
                "supabase_token": response.session.access_token
            }
        return {"status": "error", "message": "登录失败"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

def logout_user(access_token: str = None):
    supabase = get_supabase_client()
    try:
        if access_token:
            supabase.auth.sign_out()
        return {"status": "success"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

def update_user_profile(user_id: str, username: str = None, email: str = None):
    """
    更新用户个人资料
    """
    supabase = get_supabase_client()
    try:
        updates = {}
        
        if username:
            updates["data"] = {"username": username}
        
        if email:
            updates["email"] = email
        
        if not updates:
            return {"status": "error", "message": "没有提供任何更新内容"}
        
        response = supabase.auth.update_user(updates)
        
        if response.user:
            return {
                "status": "success",
                "user_id": response.user.id,
                "email": response.user.email,
                "username": response.user.user_metadata.get("username", response.user.email.split("@")[0])
            }
        return {"status": "error", "message": "更新失败"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

def get_current_user(token: str = None):
    if not token:
        auth_header = request.headers.get("Authorization", "") if request else ""
        if not auth_header.startswith("Bearer "):
            return None
        token = auth_header[7:]
    try:
        payload = jwt.decode(token, Config.JWT_SECRET_KEY, algorithms=["HS256"])
        return {
            "id": payload.get("user_id"),
            "email": payload.get("email"),
            "username": payload.get("username")
        }
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None

def generate_custom_jwt(user_id: str, email: str, username: str = None):
    payload = {
        "user_id": user_id,
        "email": email,
        "username": username or email.split("@")[0],
        "exp": datetime.now() + timedelta(hours=Config.JWT_EXPIRE_HOURS)
    }
    return jwt.encode(payload, Config.JWT_SECRET_KEY, algorithm="HS256")

def auth_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        user = get_current_user()
        if not user:
            return jsonify({"error": "请先登录"}), 401
        return f(user, *args, **kwargs)
    return decorated_function