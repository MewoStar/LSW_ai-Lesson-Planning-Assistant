from flask import Flask, request, jsonify
from app.auth import auth_required
from app.supabase_client import get_supabase_client

app = Flask(__name__)

@app.route("/api/history", methods=["GET"])
@auth_required
def get_history(user):
    try:
        supabase = get_supabase_client()
        response = supabase.table("search_history") \
            .select("*") \
            .eq("user_id", user["id"]) \
            .order("created_at", desc=True) \
            .limit(50) \
            .execute()
        return jsonify({"history": response.data or []})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/history/<history_id>", methods=["DELETE"])
@auth_required
def delete_history(user, history_id):
    try:
        supabase = get_supabase_client()
        response = supabase.table("search_history") \
            .select("id") \
            .eq("id", history_id) \
            .eq("user_id", user["id"]) \
            .execute()
        
        if not response.data:
            return jsonify({"error": "记录不存在或无权访问"}), 404
        
        supabase.table("search_history") \
            .delete() \
            .eq("id", history_id) \
            .execute()
        
        return jsonify({"status": "ok"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handler(event, context):
    return app(event, context)