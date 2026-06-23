# ============================================
# 备课助手 Web 版
# 启动: python web_app.py
# 访问: http://localhost:5000
# ============================================

import re
import os
import sys
import yaml
import random
import threading
import queue
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, Response
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import time
import json


def get_base_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent


def get_data_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


BASE_DIR = get_base_dir()
DATA_DIR = get_data_dir()

template_dir = str(BASE_DIR / "templates")
app = Flask(__name__, template_folder=template_dir)

# 读配置
CONFIG_PATH = BASE_DIR / "config.yaml"
cfg = {}
if CONFIG_PATH.exists():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

API_KEY = cfg.get("api_key", "sk-xxx")
BASE_URL = cfg.get("base_url", "https://api.deepseek.com/v1")
MODEL = cfg.get("model", "deepseek-chat")
TEMP = cfg.get("temperature", 0.7)
MAX_TOKENS = cfg.get("max_tokens", 3072)
MAX_HISTORY = cfg.get("max_history", 6)
OUTPUT_DIR = cfg.get("output_dir", ".")
if not os.path.isabs(OUTPUT_DIR):
    OUTPUT_DIR = os.path.abspath(os.path.join(str(DATA_DIR), OUTPUT_DIR))
else:
    OUTPUT_DIR = os.path.abspath(OUTPUT_DIR)

os.makedirs(OUTPUT_DIR, exist_ok=True)

SYSTEM_PROMPT = """你是资深教育专家，帮教师备课。

## 能力
1. **教案设计**：课题、课型、学段、教学目标、重难点、教学过程、板书、作业
2. **课件大纲**：PPT大纲，每页标题+要点+配图建议
3. **习题试卷**：分层出题（基础/提高/拓展），附答案解析
4. **教学资源**：复习提纲、导学案、知识点总结、实验方案

## 输出要求
- Markdown 格式，标题层级清晰
- 重点用 **粗体**，信息用表格整理
- 需保存的文件内容用 [文件:xxx.md] ... [/文件] 包裹
- 中文输出，专业、简洁、有条理"""

# 存会话（简单内存存储）
sessions: dict[str, list[dict]] = {}


def trim_history(messages):
    if not messages or len(messages) <= MAX_HISTORY + 1:
        return messages
    system_msg = messages[0]
    recent = messages[-(MAX_HISTORY):]
    return [system_msg] + recent


def markdown_to_word(markdown_text: str, title: str = "教案") -> BytesIO:
    """将 Markdown 文本转换为 Word 文档"""
    doc = Document()

    # 设置默认字体
    style = doc.styles['Normal']
    font = style.font
    font.name = '微软雅黑'
    font.size = Pt(11)

    # 添加标题
    title_para = doc.add_paragraph()
    title_run = title_para.add_run(title)
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(31, 78, 121)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    lines = markdown_text.split('\n')
    in_code_block = False
    code_content = []

    for line in lines:
        # 代码块处理
        if line.strip().startswith('```'):
            if in_code_block:
                # 结束代码块
                if code_content:
                    code_para = doc.add_paragraph()
                    code_run = code_para.add_run('\n'.join(code_content))
                    code_run.font.name = 'Courier New'
                    code_run.font.size = Pt(9)
                    code_run.font.color.rgb = RGBColor(128, 128, 128)
                    code_para.paragraph_format.left_indent = Pt(20)
                code_content = []
                in_code_block = False
            else:
                in_code_block = True
            continue

        if in_code_block:
            code_content.append(line)
            continue

        # 标题处理
        if line.startswith('# '):
            doc.add_heading(line[2:], level=1)
        elif line.startswith('## '):
            doc.add_heading(line[3:], level=2)
        elif line.startswith('### '):
            doc.add_heading(line[4:], level=3)
        elif line.startswith('#### '):
            doc.add_heading(line[5:], level=4)
        # 粗体处理
        elif '**' in line:
            parts = line.split('**')
            para = doc.add_paragraph()
            for i, part in enumerate(parts):
                if i % 2 == 1:  # 奇数位置是粗体
                    run = para.add_run(part)
                    run.bold = True
                else:
                    para.add_run(part)
        # 列表处理
        elif line.strip().startswith('- '):
            doc.add_paragraph(line[2:], style='List Bullet')
        elif line.strip().startswith('* '):
            doc.add_paragraph(line[2:], style='List Bullet')
        elif re.match(r'^\d+\.\s', line.strip()):
            doc.add_paragraph(re.sub(r'^\d+\.\s', '', line.strip()), style='List Number')
        # 表格处理
        elif '|' in line and line.strip():
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if len(cells) >= 2:
                # 检查是否是分隔线
                if not all(re.match(r'^-+$', cell) for cell in cells):
                    # 这里简化处理，实际需要更复杂的表格解析
                    para = doc.add_paragraph()
                    para.add_run(line)
        # 引用块处理
        elif line.strip().startswith('>'):
            para = doc.add_paragraph(line[1:].strip())
            para.paragraph_format.left_indent = Pt(20)
            para.runs[0].font.color.rgb = RGBColor(128, 128, 128)
        # 分割线处理
        elif line.strip() == '---':
            doc.add_paragraph('_' * 50)
        # 普通段落
        elif line.strip():
            doc.add_paragraph(line)

    # 保存到内存
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


@app.route("/")
def index():
    return render_template("index.html", model=MODEL)


@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.json
    if not data or not isinstance(data, dict):
        return jsonify({"error": "Invalid request body"}), 400

    session_id = str(data.get("session_id", "default"))[:64]
    user_message = data.get("message", "").strip()
    
    if not user_message:
        return jsonify({"error": "Message cannot be empty"}), 400
    
    if len(user_message) > 8000:
        return jsonify({"error": "Message too long (max 8000 characters)"}), 400

    _model = data.get("model", MODEL)
    try:
        _temp = max(0.0, min(2.0, float(data.get("temperature", TEMP))))
        _max_tokens = max(1, min(16384, int(data.get("max_tokens", MAX_TOKENS))))
    except (ValueError, TypeError):
        return jsonify({"error": "Invalid temperature or max_tokens value"}), 400

    if session_id not in sessions:
        sessions[session_id] = [{"role": "system", "content": SYSTEM_PROMPT}]

    sessions[session_id].append({"role": "user", "content": user_message})

    try:
        client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
        msgs = trim_history(sessions[session_id])
        response = client.chat.completions.create(
            model=_model,
            messages=msgs,
            temperature=_temp,
            max_tokens=_max_tokens,
        )
    except Exception as e:
        return jsonify({"error": f"API request failed: {str(e)}"}), 500

    reply = response.choices[0].message.content
    usage = response.usage

    sessions[session_id].append({"role": "assistant", "content": reply})

    # 提取文件
    files = re.findall(r"\[文件:\s*([^\]]+)\](.*?)\[/文件\]", reply, re.DOTALL)
    saved_files = []
    for filename, content in files:
        filepath = os.path.join(OUTPUT_DIR, filename)
        clean = content.strip()
        if clean.startswith("```"):
            clean = re.sub(r"^```\w*\s*\n", "", clean)
            clean = re.sub(r"\n```\s*$", "", clean)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(clean)
        saved_files.append(filename)

    # 清理回复中的文件标记
    display = reply
    if files:
        display = re.sub(r"\[文件:.*?\[/文件\]", "", reply, flags=re.DOTALL).strip()

    return jsonify({
        "reply": display,
        "model": response.model,
        "usage": {
            "prompt_tokens": usage.prompt_tokens,
            "completion_tokens": usage.completion_tokens,
        },
        "saved_files": saved_files,
    })


def extract_keywords(message):
    keywords = []
    subjects = ["物理", "数学", "语文", "英语", "化学", "生物", "历史", "地理", "政治", "信息技术", "通用技术", "体育", "音乐", "美术", "科学", "道德与法治", "思想品德"]
    for subj in subjects:
        if subj in message:
            keywords.append(subj)
            break
    levels = ["小学", "初中", "高中", "大学", "职业院校", "中职", "高职", "一年级", "二年级", "三年级", "四年级", "五年级", "六年级", "初一", "初二", "初三", "高一", "高二", "高三"]
    for level in levels:
        if level in message:
            keywords.append(level)
            break
    types = ["教案", "课件", "习题", "试卷", "复习", "提纲", "实验", "实训", "说课稿", "导学案", "任务单", "教学设计"]
    for t in types:
        if t in message:
            keywords.append(t)
            break
    if not keywords:
        keywords = ["教学", "备课"]
    return keywords


knowledge_base = {
    "物理": ["牛顿运动定律", "能量守恒", "电磁感应", "光学", "热力学", "力学", "相对论", "量子物理", "波动", "磁场"],
    "数学": ["函数", "方程", "几何", "概率统计", "数列", "三角函数", "向量", "导数", "积分", "不等式"],
    "语文": ["文言文", "现代文", "诗词鉴赏", "写作", "阅读", "修辞手法", "表现手法", "作文", "记叙文", "议论文"],
    "英语": ["语法", "词汇", "阅读", "写作", "听力", "口语", "时态", "句型", "翻译", "完形填空"],
    "化学": ["元素周期表", "化学反应", "有机化学", "无机化学", "化学平衡", "电化学", "化学实验", "化学键"],
    "生物": ["细胞", "遗传", "生态", "光合作用", "呼吸作用", "进化论", "微生物", "人体生理"],
    "历史": ["中国古代史", "中国近代史", "世界史", "历史事件", "历史人物", "朝代", "战争", "改革"],
    "地理": ["自然地理", "人文地理", "气候", "地形", "地图", "环境保护", "区域地理", "人口城市"],
    "政治": ["经济生活", "政治生活", "文化生活", "哲学生活", "法律", "道德", "国情"],
    "default": ["教学目标", "教学重难点", "教学过程", "教学方法", "板书设计", "课后作业", "教学反思", "学情分析"]
}


def generate_thinking_phases(user_message):
    keywords = extract_keywords(user_message)
    subject = keywords[0] if keywords else "default"
    kbs = knowledge_base.get(subject, knowledge_base["default"])
    selected_kb = random.sample(kbs, min(4, len(kbs)))

    phases = [
        {"title": "📋 需求分析阶段", "items": [
            f"解析用户需求：{user_message[:25]}{'...' if len(user_message) > 25 else ''}",
            f"识别学科与学段：{'、'.join(keywords)}",
            "明确输出类型与格式要求",
        ]},
        {"title": "🔍 知识检索阶段", "items": [
            "在知识库中检索相关教学资源...",
            f"匹配知识点：{selected_kb[0]}、{selected_kb[1]}",
            "查找课程标准与教学大纲要求",
            f"筛选参考资料：{selected_kb[2]}相关案例",
        ]},
        {"title": "🧠 内容规划阶段", "items": [
            "梳理知识体系与逻辑结构",
            "设计三维教学目标",
            "确定教学重点与难点",
            "规划教学环节与时间分配",
            "选择教学方法与策略",
        ]},
        {"title": "✍️ 内容生成阶段", "items": [
            "组织语言，生成详细内容...",
            "优化表述，确保专业准确",
            "调整结构，保证逻辑清晰",
            "检查内容完整性与合理性",
        ]},
    ]
    return phases


def thinking_to_text(phases, phase_idx, item_idx):
    lines = []
    for i, phase in enumerate(phases):
        if i > phase_idx:
            break
        lines.append(f"【{phase['title']}】")
        for j, item in enumerate(phase['items']):
            if i == phase_idx and j > item_idx:
                break
            if i < phase_idx or (i == phase_idx and j < item_idx):
                icon = "✓"
            elif i == phase_idx and j == item_idx:
                icon = "⏳"
            else:
                icon = "○"
            lines.append(f"  {icon} {item}")
        lines.append("")
    return "\n".join(lines)


@app.route("/api/chat/stream", methods=["POST"])
def chat_stream():
    data = request.json
    if not data or not isinstance(data, dict):
        return Response(
            f"data: {json.dumps({'type': 'error', 'content': 'Invalid request body'}, ensure_ascii=False)}\n\n",
            mimetype="text/event-stream"
        )

    session_id = str(data.get("session_id", "default"))[:64]
    user_message = data.get("message", "").strip()
    
    if not user_message:
        return Response(
            f"data: {json.dumps({'type': 'error', 'content': 'Message cannot be empty'}, ensure_ascii=False)}\n\n",
            mimetype="text/event-stream"
        )
    
    if len(user_message) > 8000:
        return Response(
            f"data: {json.dumps({'type': 'error', 'content': 'Message too long (max 8000 characters)'}, ensure_ascii=False)}\n\n",
            mimetype="text/event-stream"
        )

    _model = data.get("model", MODEL)
    try:
        _temp = max(0.0, min(2.0, float(data.get("temperature", TEMP))))
        _max_tokens = max(1, min(16384, int(data.get("max_tokens", MAX_TOKENS))))
    except (ValueError, TypeError):
        return Response(
            f"data: {json.dumps({'type': 'error', 'content': 'Invalid temperature or max_tokens value'}, ensure_ascii=False)}\n\n",
            mimetype="text/event-stream"
        )

    if session_id not in sessions:
        sessions[session_id] = [{"role": "system", "content": SYSTEM_PROMPT}]

    sessions[session_id].append({"role": "user", "content": user_message})

    def generate():
        thinking_phases = generate_thinking_phases(user_message)
        thinking_done = False
        ai_done = False
        ai_error = None
        full_reply = ""
        ai_queue = queue.Queue()
        first_content_received = False

        def ai_worker():
            nonlocal full_reply, ai_done, ai_error
            try:
                client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
                msgs = trim_history(sessions[session_id])
                stream = client.chat.completions.create(
                    model=_model,
                    messages=msgs,
                    temperature=_temp,
                    max_tokens=_max_tokens,
                    stream=True,
                )
                for chunk in stream:
                    if chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        full_reply += content
                        ai_queue.put(("content", content))
                ai_queue.put(("done", None))
                ai_done = True
            except Exception as e:
                ai_error = str(e)
                ai_queue.put(("error", str(e)))
                ai_done = True

        ai_thread = threading.Thread(target=ai_worker)
        ai_thread.start()

        total_thinking_time = 0
        min_thinking_time = 1.2
        max_thinking_time = 3.0

        for phase_idx, phase in enumerate(thinking_phases):
            if thinking_done:
                break
            for item_idx, item in enumerate(phase["items"]):
                if thinking_done:
                    break

                thinking_text = thinking_to_text(thinking_phases, phase_idx, item_idx)
                yield f"data: {json.dumps({'type': 'thinking', 'content': thinking_text}, ensure_ascii=False)}\n\n"

                sleep_time = random.uniform(0.25, 0.4)
                if total_thinking_time + sleep_time > max_thinking_time and phase_idx >= 2:
                    sleep_time = 0.1
                total_thinking_time += sleep_time
                time.sleep(sleep_time)

                if total_thinking_time >= min_thinking_time:
                    while not ai_queue.empty():
                        msg_type, msg_data = ai_queue.get()
                        if msg_type == "content":
                            first_content_received = True
                            break
                        elif msg_type == "done":
                            first_content_received = True
                            ai_done = True
                            break
                        elif msg_type == "error":
                            first_content_received = True
                            ai_done = True
                            break

                if first_content_received and total_thinking_time >= min_thinking_time:
                    thinking_done = True
                    final_thinking = thinking_to_text(thinking_phases, len(thinking_phases) - 1, len(thinking_phases[-1]["items"]) - 1)
                    yield f"data: {json.dumps({'type': 'thinking_done', 'content': final_thinking}, ensure_ascii=False)}\n\n"
                    break

        if not thinking_done:
            final_thinking = thinking_to_text(thinking_phases, len(thinking_phases) - 1, len(thinking_phases[-1]["items"]) - 1)
            yield f"data: {json.dumps({'type': 'thinking_done', 'content': final_thinking}, ensure_ascii=False)}\n\n"
            thinking_done = True

        if ai_error:
            yield f"data: {json.dumps({'type': 'error', 'content': ai_error}, ensure_ascii=False)}\n\n"
            return

        while not ai_done or not ai_queue.empty():
            try:
                msg_type, msg_data = ai_queue.get(timeout=0.1)
                if msg_type == "content":
                    yield f"data: {json.dumps({'type': 'content', 'content': msg_data}, ensure_ascii=False)}\n\n"
                elif msg_type == "done":
                    ai_done = True
            except queue.Empty:
                if not ai_done:
                    time.sleep(0.05)

        sessions[session_id].append({"role": "assistant", "content": full_reply})

        files = re.findall(r"\[文件:\s*([^\]]+)\](.*?)\[/文件\]", full_reply, re.DOTALL)
        saved_files = []
        for filename, file_content in files:
            filepath = os.path.join(OUTPUT_DIR, filename)
            clean = file_content.strip()
            if clean.startswith("```"):
                clean = re.sub(r"^```\w*\s*\n", "", clean)
                clean = re.sub(r"\n```\s*$", "", clean)
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(clean)
            saved_files.append(filename)

        display = full_reply
        if files:
            display = re.sub(r"\[文件:.*?\[/文件\]", "", full_reply, flags=re.DOTALL).strip()

        final_thinking_text = thinking_to_text(thinking_phases, len(thinking_phases) - 1, len(thinking_phases[-1]["items"]) - 1)
        yield f"data: {json.dumps({'type': 'done', 'content': display, 'saved_files': saved_files, 'thinking': final_thinking_text}, ensure_ascii=False)}\n\n"

    return Response(generate(), mimetype="text/event-stream")


@app.route("/api/clear", methods=["POST"])
def clear():
    data = request.json
    session_id = data.get("session_id", "default")
    sessions[session_id] = [{"role": "system", "content": SYSTEM_PROMPT}]
    return jsonify({"status": "ok"})


@app.route("/api/sessions", methods=["GET"])
def list_sessions():
    result = []
    for sid, msgs in sessions.items():
        title = "新对话"
        if len(msgs) > 1:
            first_user = next((m["content"] for m in msgs if m["role"] == "user"), "")
            title = first_user[:20] + "..." if len(first_user) > 20 else first_user
        result.append({
            "id": sid,
            "title": title,
            "message_count": len(msgs) - 1,
        })
    return jsonify({"sessions": result})


@app.route("/api/session/<session_id>", methods=["GET"])
def get_session(session_id):
    """获取单个会话的完整消息列表"""
    if session_id not in sessions:
        return jsonify({"error": "会话不存在"}), 404
    msgs = sessions[session_id]
    title = "新对话"
    user_msgs = [m for m in msgs if m["role"] == "user"]
    if user_msgs:
        first_user = user_msgs[0]["content"]
        title = first_user[:20] + "..." if len(first_user) > 20 else first_user
    return jsonify({
        "id": session_id,
        "title": title,
        "messages": msgs
    })


@app.route("/api/session/<session_id>", methods=["DELETE"])
def delete_session(session_id):
    if session_id in sessions:
        del sessions[session_id]
        return jsonify({"status": "ok"})
    return jsonify({"error": "会话不存在"}), 404


@app.route("/api/sessions/batch", methods=["DELETE"])
def delete_sessions_batch():
    data = request.json
    ids = data.get("ids", [])
    deleted = 0
    for sid in ids:
        if sid in sessions:
            del sessions[sid]
            deleted += 1
    return jsonify({"status": "ok", "deleted": deleted})


@app.route("/api/sessions/all", methods=["DELETE"])
def delete_all_sessions():
    sessions.clear()
    return jsonify({"status": "ok"})


@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify({
        "model": MODEL,
        "base_url": BASE_URL,
        "temperature": TEMP,
        "max_tokens": MAX_TOKENS,
        "output_dir": OUTPUT_DIR,
    })


@app.route("/api/download/<filename>")
def download(filename):
    filepath = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    return jsonify({"error": "文件不存在"}), 404


@app.route("/api/export/word", methods=["POST"])
def export_to_word():
    """将 Markdown 内容导出为 Word 文档，保存到文件并返回下载链接"""
    data = request.json
    if not data or not isinstance(data, dict):
        return jsonify({"error": "Invalid request body"}), 400

    markdown_content = data.get("content", "").strip()
    title = str(data.get("title", "教案")).strip()

    if not markdown_content:
        return jsonify({"error": "内容不能为空"}), 400
    
    if len(markdown_content) > 500000:
        return jsonify({"error": "内容过长，无法导出"}), 400

    try:
        word_buffer = markdown_to_word(markdown_content, title)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        safe_title = re.sub(r'[<>:"/\\|?*]', '', title)[:20]
        safe_title = safe_title if safe_title else "untitled"
        filename = f"教案_{safe_title}_{timestamp}.docx"
        filepath = os.path.join(OUTPUT_DIR, filename)

        with open(filepath, "wb") as f:
            f.write(word_buffer.getvalue())

        download_url = f"/api/download/{filename}"
        return jsonify({"download_url": download_url, "filename": filename})
    except Exception as e:
        import traceback
        return jsonify({"error": f"导出失败: {str(e)}"}), 500


if __name__ == "__main__":
    print("""
============================================
       BeiKe Assistant (Web)
       http://localhost:5000
       Ctrl+C to exit
============================================
""")
    import webbrowser
    threading.Timer(1.5, lambda: webbrowser.open("http://127.0.0.1:5000/")).start()
    app.run(host="0.0.0.0", port=5000, debug=False)
