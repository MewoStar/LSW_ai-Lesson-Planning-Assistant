from flask import Flask, request, jsonify, make_response
from app.auth import auth_required, logout_user, update_user_profile, get_current_user

app = Flask(__name__)

@app.route("/api/user/profile", methods=["GET"])
@auth_required
def get_profile(user):
    """
    获取当前用户的个人资料
    """
    try:
        return jsonify({
            "status": "success",
            "data": {
                "user_id": user["id"],
                "email": user["email"],
                "username": user["username"]
            }
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/user/profile", methods=["PUT"])
@auth_required
def update_profile(user):
    """
    更新用户个人资料（修改昵称等）
    """
    try:
        data = request.json
        
        if not data or not isinstance(data, dict):
            return jsonify({"error": "无效的请求体格式"}), 400
        
        username = data.get("username", "").strip()
        
        if not username:
            return jsonify({"error": "昵称不能为空"}), 400
        
        if len(username) < 2 or len(username) > 30:
            return jsonify({"error": "昵称长度必须在2-30个字符之间"}), 400
        
        result = update_user_profile(user["id"], username=username)
        
        if result["status"] == "success":
            return jsonify({
                "status": "success",
                "message": "昵称修改成功",
                "data": {
                    "user_id": result["user_id"],
                    "email": result["email"],
                    "username": result["username"]
                }
            })
        return jsonify({"error": result["message"]}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/user/logout", methods=["POST"])
def logout():
    """
    用户退出登录
    """
    try:
        user = get_current_user()
        
        if not user:
            return jsonify({"error": "用户未登录"}), 401
        
        result = logout_user()
        
        if result["status"] == "success":
            response = make_response(jsonify({
                "status": "success",
                "message": "退出登录成功"
            }))
            response.set_cookie('session_token', '', expires=0)
            return response
        
        return jsonify({"error": result["message"]}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handler(event, context):
    return app(event, context)