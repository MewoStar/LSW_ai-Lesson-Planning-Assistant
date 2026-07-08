from flask import Flask, request, jsonify
from app.auth import login_user

app = Flask(__name__)

@app.route("/api/login", methods=["POST"])
def login():
    try:
        data = request.json
        if not data or not data.get("email") or not data.get("password"):
            return jsonify({"error": "邮箱和密码不能为空"}), 400
        
        email = data.get("email")
        password = data.get("password")
        
        result = login_user(email, password)
        if result["status"] == "success":
            return jsonify(result)
        return jsonify({"error": result["message"]}), 401
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handler(event, context):
    return app(event, context)