import json

def handler(event, context):
    try:
        body = json.loads(event.get("body", "{}")) if event.get("body") else {}
        
        if not body or not body.get("email") or not body.get("password"):
            return {
                "statusCode": 400,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": "邮箱和密码不能为空"}
            }
        
        email = body.get("email")
        password = body.get("password")
        username = body.get("username")
        
        if len(password) < 6:
            return {
                "statusCode": 400,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": "密码长度至少6位"}
            }
        
        from app.auth import register_user
        result = register_user(email, password, username)
        
        if result["status"] == "success":
            return {
                "statusCode": 201,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": result
            }
        return {
            "statusCode": 409,
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