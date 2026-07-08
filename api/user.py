import json

def handler(event, context):
    try:
        headers = event.get("headers", {})
        auth_header = headers.get("Authorization", headers.get("authorization", ""))
        
        method = event.get("httpMethod", "GET")
        
        if method == "POST" and event.get("path", "").endswith("/logout"):
            if not auth_header.startswith("Bearer "):
                return {
                    "statusCode": 401,
                    "headers": {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    },
                    "body": {"error": "用户未登录"}
                }
            
            token = auth_header[7:]
            
            from app.auth import get_current_user, logout_user
            user = get_current_user(token)
            
            if not user:
                return {
                    "statusCode": 401,
                    "headers": {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    },
                    "body": {"error": "用户未登录"}
                }
            
            result = logout_user()
            
            if result["status"] == "success":
                return {
                    "statusCode": 200,
                    "headers": {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    },
                    "body": {"status": "success", "message": "退出登录成功"}
                }
            return {
                "statusCode": 500,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": result["message"]}
            }
        
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
        
        if method == "GET":
            return {
                "statusCode": 200,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {
                    "status": "success",
                    "data": {
                        "user_id": user["id"],
                        "email": user["email"],
                        "username": user["username"]
                    }
                }
            }
        
        elif method == "PUT":
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
            
            username = body.get("username", "").strip()
            
            if not username:
                return {
                    "statusCode": 400,
                    "headers": {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    },
                    "body": {"error": "昵称不能为空"}
                }
            
            if len(username) < 2 or len(username) > 30:
                return {
                    "statusCode": 400,
                    "headers": {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    },
                    "body": {"error": "昵称长度必须在2-30个字符之间"}
                }
            
            from app.auth import update_user_profile
            result = update_user_profile(user["id"], username=username)
            
            if result["status"] == "success":
                return {
                    "statusCode": 200,
                    "headers": {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    },
                    "body": {
                        "status": "success",
                        "message": "昵称修改成功",
                        "data": {
                            "user_id": result["user_id"],
                            "email": result["email"],
                            "username": result["username"]
                        }
                    }
                }
            return {
                "statusCode": 400,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": result["message"]}
            }
        
        else:
            return {
                "statusCode": 405,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"error": "不支持的请求方法"}
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