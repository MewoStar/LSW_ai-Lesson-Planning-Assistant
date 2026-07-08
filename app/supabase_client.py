from supabase import create_client, Client
from app.config import Config

def get_supabase_client() -> Client:
    if not Config.SUPABASE_URL or not Config.SUPABASE_ANON_KEY:
        raise RuntimeError("SUPABASE_URL 和 SUPABASE_ANON_KEY 必须设置")
    return create_client(Config.SUPABASE_URL, Config.SUPABASE_ANON_KEY)

def get_supabase_admin_client() -> Client:
    if not Config.SUPABASE_URL or not Config.SUPABASE_SERVICE_ROLE_KEY:
        raise RuntimeError("SUPABASE_URL 和 SUPABASE_SERVICE_ROLE_KEY 必须设置")
    return create_client(Config.SUPABASE_URL, Config.SUPABASE_SERVICE_ROLE_KEY)