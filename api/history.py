import json

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
        
        method = event.get("httpMethod", "GET")
        
        if method == "GET":
            from app.supabase_client import get_supabase_client
            supabase = get_supabase_client()
            response = supabase.table("search_history") \
                .select("*") \
                .eq("user_id", user["id"]) \
                .order("created_at", desc=True) \
                .limit(50) \
                .execute()
            return {
                "statusCode": 200,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"history": response.data or []}
            }
        
        elif method == "DELETE":
            path = event.get("path", "")
            history_id = path.split("/")[-1] if "/" in path else ""
            
            if not history_id:
                return {
                    "statusCode": 400,
                    "headers": {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    },
                    "body": {"error": "缺少历史记录ID"}
                }
            
            from app.supabase_client import get_supabase_client
            supabase = get_supabase_client()
            
            response = supabase.table("search_history") \
                .select("id") \
                .eq("id", history_id) \
                .eq("user_id", user["id"]) \
                .execute()
            
            if not response.data:
                return {
                    "statusCode": 404,
                    "headers": {
                        "Content-Type": "application/json",
                        "Access-Control-Allow-Origin": "*"
                    },
                    "body": {"error": "记录不存在或无权访问"}
                }
            
            supabase.table("search_history") \
                .delete() \
                .eq("id", history_id) \
                .execute()
            
            return {
                "statusCode": 200,
                "headers": {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*"
                },
                "body": {"status": "ok"}
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