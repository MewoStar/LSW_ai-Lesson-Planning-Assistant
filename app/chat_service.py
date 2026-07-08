from openai import OpenAI
from app.config import Config
from app.supabase_client import get_supabase_client
import re

class ChatService:
    def __init__(self):
        self.client = OpenAI(
            api_key=Config.OPENAI_API_KEY,
            base_url=Config.OPENAI_BASE_URL
        )
    
    def chat_completion(self, user_id: str, query: str, session_id: str = "default",
                       model: str = None, temperature: float = None,
                       max_tokens: int = None) -> dict:
        model = model or Config.DEFAULT_MODEL
        temperature = temperature or Config.DEFAULT_TEMPERATURE
        max_tokens = max_tokens or Config.DEFAULT_MAX_TOKENS
        
        system_prompt = self._get_system_prompt()
        messages = [{"role": "system", "content": system_prompt}] + [
            {"role": "user", "content": query}
        ]
        
        try:
            response = self.client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens
            )
            
            reply = response.choices[0].message.content
            usage = response.usage
            
            self._save_search_history(user_id, query, model)
            lesson_plan = self._parse_and_save_lesson_plan(user_id, query, reply)
            
            return {
                "status": "success",
                "reply": reply,
                "model": response.model,
                "usage": {
                    "prompt_tokens": usage.prompt_tokens,
                    "completion_tokens": usage.completion_tokens
                },
                "lesson_plan_id": lesson_plan.get("id") if lesson_plan else None
            }
        
        except Exception as e:
            error_msg = str(e)
            error_types = {
                "rate_limit": "API 请求过于频繁，请稍后重试",
                "api_key": "API 密钥无效，请检查配置",
                "network": "网络连接异常，请检查网络设置",
                "timeout": "请求超时，请稍后重试"
            }
            lower_error = error_msg.lower()
            for error_type, friendly_msg in error_types.items():
                if error_type in lower_error:
                    return {"status": "error", "message": friendly_msg}
            return {"status": "error", "message": "服务暂时不可用，请稍后重试"}
    
    def _get_system_prompt(self) -> str:
        return """你是专业的「AI 备课助手」，服务对象是中小学 / 职业院校的一线教师。
你的任务是根据教师的需求，按规范格式生成高质量的教学资源。

## 核心能力
1. **PPT 课件生成**：输出结构化 PPT 大纲
2. **教案生成**：45 分钟标准格式教案
3. **习题生成**：基础/提高/拓展三级难度分层习题
4. **作业分析**：作业批改数据分析与辅导建议
5. **复习提纲**：期末/单元复习资料

## 输出规则
- 必须用 `[文件:xxx.md] ... [/文件]` 或 `[PPT:xxx.pptx]` 包裹可保存内容
- Markdown 格式，标题层级清晰
- 重点用 **粗体**，信息用表格整理
- 中文输出，专业简洁"""
    
    def _save_search_history(self, user_id: str, query: str, model: str):
        supabase = get_supabase_client()
        try:
            supabase.table("search_history").insert({
                "user_id": user_id,
                "query": query,
                "model": model
            }).execute()
        except Exception:
            pass
    
    def _parse_and_save_lesson_plan(self, user_id: str, query: str, reply: str) -> dict:
        content_type = self._detect_content_type(query, reply)
        title = self._extract_title(query, reply)
        
        if not title or not content_type:
            return {}
        
        supabase = get_supabase_client()
        try:
            response = supabase.table("lesson_plans").insert({
                "user_id": user_id,
                "title": title,
                "content": reply,
                "content_type": content_type,
                "model": Config.DEFAULT_MODEL
            }).execute()
            if response.data:
                return response.data[0]
        except Exception:
            pass
        return {}
    
    def _detect_content_type(self, query: str, reply: str) -> str:
        lower = (query + " " + reply).lower()
        if "ppt" in lower or "课件" in lower or "幻灯片" in lower:
            return "ppt_outline"
        elif "教案" in lower or "教学设计" in lower:
            return "lesson_plan"
        elif "习题" in lower or "练习" in lower or "试卷" in lower:
            return "exercises"
        elif "复习" in lower or "提纲" in lower:
            return "review_notes"
        elif "作业分析" in lower or "错因" in lower:
            return "analysis_report"
        return "lesson_plan"
    
    def _extract_title(self, query: str, reply: str) -> str:
        title_match = re.search(r"\[文件:\s*([^\]]+)\]", reply)
        if title_match:
            return title_match.group(1)[:100]
        return query[:50] or "未命名"