from flask import Flask, jsonify
from app.supabase_client import get_supabase_client
import os
import time

app = Flask(__name__)

@app.route("/api/health", methods=["GET"])
def health():
    start_time = time.time()
    
    env_check = {
        "SUPABASE_URL": "✅" if os.environ.get("SUPABASE_URL") else "❌",
        "SUPABASE_ANON_KEY": "✅" if os.environ.get("SUPABASE_ANON_KEY") else "❌",
        "OPENAI_API_KEY": "✅" if os.environ.get("OPENAI_API_KEY") else "❌",
        "JWT_SECRET_KEY": "✅" if os.environ.get("JWT_SECRET_KEY") else "❌"
    }
    
    db_status = "✅"
    db_error = None
    try:
        supabase = get_supabase_client()
        response = supabase.table("search_history").select("id").limit(1).execute()
        if response.data is not None:
            db_status = "✅"
    except Exception as e:
        db_status = "❌"
        db_error = str(e)[:100]
    
    response_time = round((time.time() - start_time) * 1000, 2)
    
    all_healthy = all(v == "✅" for v in env_check.values()) and db_status == "✅"
    
    return jsonify({
        "status": "healthy" if all_healthy else "unhealthy",
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "response_time_ms": response_time,
        "environment": env_check,
        "database": {
            "status": db_status,
            "error": db_error
        }
    })

def handler(event, context):
    return app(event, context)