import json
import os

def handler(event, context):
    try:
        headers = event.get("headers", {})
        auth_header = headers.get("Authorization", headers.get("authorization", ""))
        
        if not auth_header.startswith("Bearer "):
            return {
                "statusCode": 401,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": "请先登录"}
            }
        
        token = auth_header[7:]
        
        from app.auth import get_current_user
        user = get_current_user(token)
        
        if not user:
            return {
                "statusCode": 401,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": "请先登录"}
            }
        
        body = json.loads(event.get("body", "{}")) if event.get("body") else {}
        
        if not body or not isinstance(body, dict):
            return {
                "statusCode": 400,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": "无效的请求体格式"}
            }
        
        message = body.get("message", "").strip()
        if not message:
            return {
                "statusCode": 400,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": "消息内容不能为空"}
            }
        
        if len(message) > 8000:
            return {
                "statusCode": 400,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": "消息内容过长（最大8000字符）"}
            }
        
        session_id = body.get("session_id", "default")
        if not isinstance(session_id, str) or len(session_id) > 64:
            session_id = "default"
        
        try:
            temperature = float(body.get("temperature", 0.7))
            temperature = max(0.0, min(2.0, temperature))
        except (ValueError, TypeError):
            temperature = 0.7
        
        try:
            max_tokens = int(body.get("max_tokens", 3072))
            max_tokens = max(1, min(16384, max_tokens))
        except (ValueError, TypeError):
            max_tokens = 3072
        
        model = body.get("model")
        
        from app.chat_service import ChatService
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
            return {
                "statusCode": 200,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {
                    "reply": result["reply"],
                    "model": result["model"],
                    "usage": result.get("usage", {}),
                    "lesson_plan_id": result.get("lesson_plan_id")
                }
            }
        return {
            "statusCode": 500,
            "headers": {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*"
            },
            "body": {"error": result["message"]}
        }
    except Exception as e:
        return {
            "statusCode": 500,
            "headers": {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*"
            },
            "body": {"error": str(e)}
        }