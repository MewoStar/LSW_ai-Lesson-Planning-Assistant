from flask import Flask, request, jsonify
from app.auth import auth_required
from app.chat_service import ChatService

app = Flask(__name__)

@app.route("/api/chat", methods=["POST"])
@auth_required
def chat(user):
    try:
        data = request.json
        if not data or not isinstance(data, dict):
            return jsonify({"error": "无效的请求体格式"}), 400
        
        message = data.get("message", "").strip()
        if not message:
            return jsonify({"error": "消息内容不能为空"}), 400
        
        if len(message) > 8000:
            return jsonify({"error": "消息内容过长（最大8000字符）"}), 400
        
        session_id = data.get("session_id", "default")
        if not isinstance(session_id, str) or len(session_id) > 64:
            session_id = "default"
        
        try:
            temperature = float(data.get("temperature", 0.7))
            temperature = max(0.0, min(2.0, temperature))
        except (ValueError, TypeError):
            temperature = 0.7
        
        try:
            max_tokens = int(data.get("max_tokens", 3072))
            max_tokens = max(1, min(16384, max_tokens))
        except (ValueError, TypeError):
            max_tokens = 3072
        
        model = data.get("model")
        
        service = ChatService()
        result = service.chat_completion(
            user_id=user["id"],
            query=message,
            session_id=session_id,
            model=model,
            temperature=temperature,
            max_tokens=max_tokens
        )
        
        if result["status"] == "success":
            return jsonify({
                "reply": result["reply"],
                "model": result["model"],
                "usage": result.get("usage", {}),
                "lesson_plan_id": result.get("lesson_plan_id")
            })
        return jsonify({"error": result["message"]}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def handler(event, context):
    return app(event, context)