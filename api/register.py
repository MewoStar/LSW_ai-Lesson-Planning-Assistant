from flask import Flask, request, jsonify
from app.auth import register_user

app = Flask(__name__)

@app.route("/api/register", methods=["POST"])
def register():
    try:
        data = request.json
        if not data or not data.get("email") or not data.get("password"):
            return jsonify({"error": "邮箱和密码不能为空"}), 400
        
        email = data.get("email")
        password = data.get("password")
        username = data.get("username")
        
        if len(password) < 6:
            return jsonify({"error": "密码长度至少6位"}), 400
        
        result = register_user(email, password, username)
        if result["status"] == "success":
            return jsonify(result), 201
        return jsonify({"error": result["message"]}), 409
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handler(event, context):
    return app(event, context)