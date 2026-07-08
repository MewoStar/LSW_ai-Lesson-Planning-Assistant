import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
    SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")
    SUPABASE_SERVICE_ROLE_KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
    
    OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
    OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.deepseek.com/v1")
    DEFAULT_MODEL = os.environ.get("DEFAULT_MODEL", "deepseek-chat")
    DEFAULT_TEMPERATURE = float(os.environ.get("DEFAULT_TEMPERATURE", "0.7"))
    DEFAULT_MAX_TOKENS = int(os.environ.get("DEFAULT_MAX_TOKENS", "3072"))
    
    JWT_SECRET_KEY = os.environ.get("JWT_SECRET_KEY", "beike-jwt-secret-key-2026")
    JWT_EXPIRE_HOURS = int(os.environ.get("JWT_EXPIRE_HOURS", "24"))
    
    @classmethod
    def validate(cls):
        required = ["SUPABASE_URL", "SUPABASE_ANON_KEY", "OPENAI_API_KEY"]
        missing = [k for k in required if not getattr(cls, k)]
        if missing:
            raise ValueError(f"缺少必要环境变量: {', '.join(missing)}")