# ============================================
# 可视化生成模块 - 后端代码
# ============================================
# 用法：将本文件中的 register_visual_routes(app, ...) 集成到目标项目即可
# 依赖：Flask、openai、re、json、dashscope (可选)、requests (可选)
# ============================================

import re
import json
import time
import threading
import traceback
from html import escape as html_escape
from flask import jsonify, render_template, redirect, url_for, request

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False


_VISUAL_RATE_LIMIT = {}
_VISUAL_RATE_LOCK = threading.Lock()

def _get_visual_client_ip():
    if request.headers.get('X-Forwarded-For'):
        forwarded = request.headers.get('X-Forwarded-For')
        parts = forwarded.split(',')
        for part in parts:
            ip = part.strip()
            if ip and not ip.startswith('127.') and not ip.startswith('10.') and not ip.startswith('172.'):
                return ip
        return parts[0].strip() if parts else 'unknown'
    return request.remote_addr or 'unknown'

def _check_visual_rate_limit(ip: str) -> bool:
    now = time.time()
    with _VISUAL_RATE_LOCK:
        expired_keys = [k for k, v in _VISUAL_RATE_LIMIT.items() if all(now - t > 60 for t in v)]
        for k in expired_keys:
            del _VISUAL_RATE_LIMIT[k]
        if ip not in _VISUAL_RATE_LIMIT:
            _VISUAL_RATE_LIMIT[ip] = []
        attempts = _VISUAL_RATE_LIMIT[ip]
        attempts = [t for t in attempts if now - t < 60]
        _VISUAL_RATE_LIMIT[ip] = attempts
        if len(attempts) >= 5:
            return False
        _VISUAL_RATE_LIMIT[ip].append(now)
    return True


# --------------------------------------------
# 集成函数：注册可视化模块所有路由
# --------------------------------------------
def register_visual_routes(
    app,
    get_current_user,
    safe_get_json,
    database,
    openai_client_factory,
    model_name,
    modules_data,
    qwen_image_api_key="",
    image_search_provider="",
    image_search_api_key="",
    static_url_path="/static"
):
    """
    注册可视化生成模块的所有路由。

    调用此函数后，会自动添加：
      - GET  /visual                 : 重定向到 /module/visualization
      - POST /api/visual/generate    : 生成思维导图 / 教学动画

    参数：
      - app: Flask 应用实例
      - get_current_user: 获取当前用户的函数
      - safe_get_json: 安全获取JSON请求体的函数
      - database: 数据库模块
      - openai_client_factory: 返回 OpenAI 客户端的函数
      - model_name: 使用的模型名称
      - modules_data: 全局模块字典（包含 visualization 条目）
      - qwen_image_api_key: 通义万相图片生成API Key（可选）
      - image_search_provider: 图片搜索服务提供商（unsplash/pexels，可选）
      - image_search_api_key: 图片搜索API Key（可选）
      - static_url_path: 静态资源路径
    """

    @app.route("/visual")
    def visual_redirect():
        return redirect(url_for('module_detail', module_id='visualization'))

    @app.route("/api/visual/generate", methods=["POST"])
    def api_visual_generate():
        try:
            user = get_current_user()
            if not user:
                return jsonify({"success": False, "error": "请先登录"}), 401

            ip = _get_visual_client_ip()
            if not _check_visual_rate_limit(ip):
                return jsonify({"success": False, "error": "请求过于频繁，请1分钟后再试"}), 429

            data = safe_get_json()
            if not data:
                return jsonify({"success": False, "error": "无效请求"}), 400

            topic = (data.get("topic") or "").strip()[:500]
            gen_type = data.get("type", "mindmap")
            layout = data.get("layout", "logic")

            if not topic:
                return jsonify({"success": False, "error": "请输入课程主题"}), 400

            client = openai_client_factory()

            if gen_type == "mindmap":
                return _handle_mindmap(client, user, topic, layout, database, model_name)
            elif gen_type == "animation":
                return _handle_animation(
                    client, user, topic, database, model_name,
                    qwen_image_api_key, image_search_provider, image_search_api_key
                )
            else:
                return jsonify({"success": False, "error": "未知的生成类型"}), 400
        except Exception as e:
            traceback.print_exc()
            print(f"[api_visual_generate] 异常: {e}")
            # 不暴露完整异常信息给前端，避免泄露敏感路径/API key
            return jsonify({"success": False, "error": "生成失败，请稍后重试"}), 500

    return app


# --------------------------------------------
# 思维导图生成（支持 3 种布局：逻辑图 / 组织结构图 / 矩形思维导图）
# --------------------------------------------
def _handle_mindmap(client, user, topic, layout, database, model_name):
    # 修复 L2：下方各 layout 分支的 prompt 为超长硬编码字符串，建议后续抽取为外部模板文件或模块级常量以便维护
    if layout == "tree":
        layout_desc = "组织结构图/树状图"
        format_hint = """# 角色定义
你是一位精通组织管理与信息分类的可视化专家。你必须严格区分并遵循「组织结构图」与「树状图」的专属规范，禁止将其与流程图、概念图或发散式思维导图混淆。

# 核心定义
- 组织结构图（Org Chart）：以垂直自上而下为主，展示职位/部门间的汇报、隶属与管理跨度关系，强调权责层级。
- 树状图（Tree Diagram）：可垂直或水平展开，展示事物的分类、组成或分解关系，强调MECE（相互独立、完全穷尽）原则。

# 格式规范（必须100%遵守）
1. 使用 Mermaid flowchart TD（自上而下）格式
2. 节点形状：仅允许矩形方框 []，根节点/高管节点可用加粗边框，禁止圆形、椭圆、菱形、云朵形、圆角矩形
3. 连线样式：仅允许直线 --> 表示父子关系，禁止曲线、波浪线、虚线
4. 布局方向：强制自上而下（TD），同一层级节点水平对齐，上级居中于下级上方
5. 层级表达：通过分支位置体现父子级，最多3-4级深度，跨级连接需标注说明
6. 文字规范：节点内文字≤8字，名词/职位名称优先，禁止动词长句、解释性描述
7. 至少3-5个一级分支，每个分支下有2-4个子节点
8. 同级节点必须严格对齐，禁止斜线连接"""
        output_format = "mermaid"
    elif layout == "rect":
        layout_desc = "矩形思维导图"
        format_hint = """# 角色定义
你是一位精通结构化思维的可视化专家。当用户提到"矩形思维导图"时，你必须将其理解为「以矩形为节点、直线为连接的层级拆解图」，禁止生成传统曲线发散式思维导图。

# 核心定义
矩形思维导图 = 思维导图的层级逻辑 + 逻辑图的视觉规范。
它用于对主题进行MECE拆解、分类归纳或流程梳理，强调结构的严谨性与信息的可读性，而非创意发散。

# 格式规范（必须100%遵守）
1. 使用 Mermaid flowchart LR（左→右）格式，内容较多时可用TD（上→下）
2. 节点形状：所有节点（含中心主题）均为矩形方框 []，禁止圆形、椭圆、云朵形、圆角矩形
3. 连线样式：仅允许直线 --> 表示父子关系，禁止任何曲线、波浪线、手绘线条
4. 布局方向：优先左→右展开（中心主题在最左侧），内容较多时可用上→下；禁止中心辐射状布局
5. 对齐规则：同级节点严格垂直/水平对齐，父子节点间距一致，整体呈现网格化秩序感
6. 文字规范：节点内文字≤8字，关键词/短语优先，禁止长句、解释性描述、emoji
7. 至少3-5个一级分支，每个分支下有2-4个子节点
8. 同级节点必须严格对齐，禁止斜线连接"""
        output_format = "mermaid"
    else:
        layout_desc = "逻辑图/逻辑结构图，以矩形方框为节点、直线/直角折线为连接、从左到右或从上到下呈现层级递进关系的结构化图表，强调因果、流程、分类或组成关系"
        format_hint = """格式要求（必须100%遵守）：
1. 使用 Mermaid flowchart LR（从左到右）格式
2. 节点形状：仅允许矩形方框 []，禁止圆形、椭圆、云朵形、圆角矩形、手绘边框
3. 连线样式：仅允许直线 --> 或直角折线 ---，禁止曲线、波浪线；备选路径用虚线 -.-> 表示
4. 布局方向：仅允许左→右（LR）或上→下（TD），禁止中心辐射、环形、自由散点布局
5. 层级表达：通过分支位置体现父子级，同一层级节点水平/垂直对齐
6. 文字规范：节点内文字≤10字，动宾结构优先，禁止长句、解释性描述"""
        output_format = "mermaid"

    if output_format == "mermaid":
        if layout == "tree":
            prompt = f"""请为课程主题"{topic}"生成一个{layout_desc}的Mermaid流程图。

{format_hint}

# 正例格式锚点（必须模仿此结构）

【组织结构图示例】
flowchart TD
    A[总经理] --> B[市场部]
    A --> C[技术部]
    A --> D[人事部]
    B --> E[品牌组]
    B --> F[推广组]
    C --> G[后端组]
    C --> H[前端组]

【树状图示例】
flowchart TD
    A[电子产品] --> B[手机]
    A --> C[电脑]
    B --> D[智能手机]
    B --> E[功能手机]
    C --> F[笔记本]
    C --> G[台式机]

# 反例（禁止生成）
- 用箭头表示"市场部→技术部"的协作关系（× 这是流程图）
- 中心写"公司"，四周发散出各部门曲线分支（× 传统思维导图）
- 同级节点未对齐、连线为斜线或曲线（× 非标准树状结构）
- 节点内写"负责产品推广与品牌建设"等长句（× 违反文字规范）
- 使用()、{{}}、>等非矩形节点形状（× 违反节点形状规范）

# 执行指令
1. 先判断场景：涉及职位/汇报→组织结构图；涉及分类/拆解→树状图
2. 输出时必须严格按上述正例格式锚点的结构呈现
3. 若课程内容不满足MECE或层级混乱，主动调整使分类相互独立、完全穷尽
4. 强制采用自上而下布局（TD），永远不要使用LR/RL/BT
5. 永远不要添加"如图所示""参见下图"等无法渲染的描述，仅输出结构化文本
6. 节点文字必须≤8字，名词优先，禁止动词长句

直接输出Mermaid代码，不要其他解释，不要代码块包裹

节点命名规则：使用A、B、C等字母作为节点ID，方括号内写中文内容（≤8字）。
连接线规则：仅使用-->表示父子隶属关系，禁止曲线和虚线。
至少3-5个核心节点，形成完整的层级结构。
"""
        elif layout == "rect":
            prompt = f"""请为课程主题"{topic}"生成一个{layout_desc}的Mermaid流程图。

{format_hint}

# 正例格式锚点（必须模仿此结构）

【左→右矩形思维导图示例】
flowchart LR
    A[新媒体运营] --> B[内容生产]
    A --> C[渠道分发]
    A --> D[数据复盘]
    B --> E[选题策划]
    B --> F[脚本撰写]
    B --> G[视觉设计]
    G --> H[封面制作]
    G --> I[排版美化]
    C --> J[微信公众号]
    C --> K[视频号]
    C --> L[小红书]
    D --> M[阅读量分析]
    D --> N[转化率优化]

【上→下矩形思维导图示例（适用于宽层级）】
flowchart TD
    A[年度营销计划] --> B[Q1拉新]
    A --> C[Q2留存]
    A --> D[Q3变现]
    B --> E[社媒投放]
    B --> F[KOL合作]
    C --> G[会员体系]
    D --> H[直播带货]
    D --> I[私域转化]

# 反例（禁止生成）
- 中心写"运营"，四周用曲线发散出"内容""渠道"等分支（× 传统博赞式思维导图）
- 节点为圆角矩形或带阴影立体效果（× 非标准矩形）
- 同级节点未对齐、连线为斜线或自由曲线（× 失去网格秩序）
- 节点内写"负责公众号推文撰写与排版"等长句（× 违反文字规范）
- 使用()、{{}}、>等非矩形节点形状（× 违反节点形状规范）

# 执行指令
1. 自动将需求转换为左→右或上→下的矩形层级结构
2. 输出时必须严格按格式锚点的符号与对齐方式呈现
3. 若用户提供的内容存在层级交叉或非MECE问题，主动指出并建议调整后再输出
4. 优先采用左→右布局（LR），分支较多时可用上→下（TD）
5. 永远不要添加"如图所示""见下图"等无法渲染的描述，仅输出纯结构化文本
6. 节点文字必须≤8字，关键词/短语优先，禁止长句、解释性描述、emoji

直接输出Mermaid代码，不要其他解释，不要代码块包裹

节点命名规则：使用A、B、C等字母作为节点ID，方括号内写中文内容（≤8字）。
连接线规则：仅使用-->表示父子关系，禁止曲线和虚线。
至少3-5个核心节点，形成完整的层级结构。
"""
        else:
            prompt = f"""请为课程主题"{topic}"生成一个{layout_desc}的Mermaid流程图。

{format_hint}

正例（逻辑图）：
flowchart LR
    A[用户注册] --> B[手机号验证] --> C[设置密码] --> D[完成注册]
    B -.-> E[邮箱验证]

反例（禁止生成）：
- 中心写"注册"，四周发散出"手机""邮箱""密码"等曲线分支（× 传统思维导图）
- 用圆角矩形+箭头表示"验证成功/失败"的判断菱形（× 流程图）
- 节点为手绘体圆圈，连线为彩色曲线（× 博赞式思维导图）

执行指令：
1. 先确认内容是否适合逻辑图（若为纯创意发散，主动建议改用其他形式）
2. 输出时必须严格按上述格式规范组织文本结构
3. 默认采用左→右布局（LR）
4. 永远不要添加"如图所示""参见下图"等无法渲染的描述

直接输出Mermaid代码，不要其他解释，不要代码块包裹

节点命名规则：使用A、B、C等字母，或简短中文（不超过4字）。
连接线规则：使用-->表示主要路径，-.->表示备选路径。
至少3-5个核心节点，形成完整的逻辑链。
"""
    # 修复 M1：注：当前所有 layout 分支都设置 output_format='mermaid'，此 markdown 分支为死代码
    # 若需支持 markdown 输出，需增加触发条件（如新增 layout 参数或配置开关）
    else:  # markdown 分支（当前为死代码，保留以备扩展）
        prompt = f"""请为课程主题"{topic}"生成一个{layout_desc}的Markdown思维导图，{format_hint}

直接输出Markdown，不要其他解释，不要代码块包裹

示例格式：
# {topic}
- 核心概念1
  - 细节1
  - 细节2
- 核心概念2
  - 细节1
  - 细节2
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=2000
    )
    # 修复 M2：AI 响应空校验，避免 choices 为空时抛 IndexError
    if not response or not response.choices:
        return jsonify({"success": False, "error": "AI 返回为空，请重试"}), 500
    content = response.choices[0].message.content.strip()
    if not content:
        return jsonify({"success": False, "error": "AI 返回内容为空"}), 500

    content = re.sub(r'^```\w*\s*', '', content)
    content = re.sub(r'\s*```$', '', content)

    if output_format == "markdown" and not content.startswith('#'):
        content = f"# {topic}\n" + content

    try:
        database.save_search_history(user['id'], f"可视化生成-思维导图({layout})：{topic}")
    except Exception:
        pass

    return jsonify({
        "success": True,
        "type": "mindmap",
        "layout": layout,
        "format": output_format,
        "content": content
    })


# --------------------------------------------
# 教学动画生成（支持 SVG+CSS 动画、Qwen 图片生成、图片搜索）
# --------------------------------------------

# 修复 H1：AI 生成 HTML 清理不充分导致 XSS
# 增强清理逻辑：覆盖更多危险标签/属性、on* 事件、javascript: 协议变形、CSS expression()、data:text/html 等
# 注意：保留 <svg> 标签（教学动画依赖 SVG 绘制图形），但清理 SVG 内的危险子元素与外部引用
def _sanitize_anim_html(html):
    if not html:
        return ''
    # 1. 移除危险标签（含内容）—— 不含 svg，因动画功能依赖 SVG
    dangerous_tags = ['script', 'iframe', 'object', 'embed', 'math', 'form', 'input', 'button', 'meta', 'link', 'base', 'foreignobject']
    for tag in dangerous_tags:
        html = re.sub(rf'<{tag}\b[^>]*>[\s\S]*?</{tag}>', '', html, flags=re.IGNORECASE)
        html = re.sub(rf'<{tag}\b[^>]*/?>', '', html, flags=re.IGNORECASE)
    # 2. 移除所有 on* 事件属性（支持双引号、单引号、无引号三种写法）
    html = re.sub(r'\son\w+\s*=\s*"[^"]*"', '', html, flags=re.IGNORECASE)
    html = re.sub(r"\son\w+\s*=\s*'[^']*'", '', html, flags=re.IGNORECASE)
    html = re.sub(r'\son\w+\s*=\s*[^\s>]+', '', html, flags=re.IGNORECASE)
    # 3. 移除 javascript: 协议（含制表符/换行符等变形）
    html = re.sub(r'javascript:', '', html, flags=re.IGNORECASE)
    html = re.sub(r'java\tscript:', '', html, flags=re.IGNORECASE)
    html = re.sub(r'java\nscript:', '', html, flags=re.IGNORECASE)
    # 4. 移除 CSS 中的 expression()（旧版 IE 可执行代码）
    html = re.sub(r'expression\s*\([^)]*\)', '', html, flags=re.IGNORECASE)
    # 5. 移除 data:text/html 协议（可能被用于注入 HTML）
    html = re.sub(r'data:text/html', '', html, flags=re.IGNORECASE)
    # 6. SVG 专用清理：移除外部引用（xlink:href / href 指向 javascript: 或 http(s):）
    html = re.sub(r'\s(xlink:href|href)\s*=\s*"[^"]*"', lambda m: '' if re.search(r'(javascript|https?:)', m.group(0), re.IGNORECASE) else m.group(0), html, flags=re.IGNORECASE)
    html = re.sub(r"\s(xlink:href|href)\s*=\s*'[^']*'", lambda m: '' if re.search(r'(javascript|https?:)', m.group(0), re.IGNORECASE) else m.group(0), html, flags=re.IGNORECASE)
    return html


def _handle_animation(
    client, user, topic, database, model_name,
    qwen_image_api_key="", image_search_provider="", image_search_api_key=""
):
    # 修复 L1：此函数较长，建议后续拆分为 _try_qwen_image、_search_images、_build_animation_prompt、_sanitize_anim_html 等子函数
    # 当前保持单函数结构以降低改动风险
    # Qwen 图片生成（如果配置了）
    if qwen_image_api_key:
        try:
            import dashscope
            from dashscope import ImageSynthesis

            dashscope.api_key = qwen_image_api_key

            anim_prompt = f"""教学主题"{topic}"的教学演示插图，用于课件展示，风格：教育类插画，清晰简洁，适合课堂教学使用"""

            result = ImageSynthesis.call(
                model=ImageSynthesis.Models.wanx_v1,
                prompt=anim_prompt,
                n=1,
                size='1280*720'
            )

            # 修复 M4：注意 ImageSynthesis.call 在某些版本是异步的
            # 若返回 PENDING，需要轮询 task_status 直到 SUCCEEDED 或 FAILED；当前直接回退到图片搜索
            task_status = getattr(result.output, 'task_status', None) if result.output else None
            if result.status_code == 200 and task_status == 'SUCCEEDED' and result.output.results:
                image_url = result.output.results[0].url

                safe_topic = html_escape(topic)
                safe_url = html_escape(image_url)

                anim_html = f'''<div class="classroom-animation" style="width:500px;height:320px;background:transparent;display:flex;align-items:center;justify-content:center;">
                            <img src="{safe_url}" alt="{safe_topic}" style="max-width:100%;max-height:100%;border-radius:8px;" />
                        </div>'''

                try:
                    database.save_search_history(user['id'], f"可视化生成-Qwen图像：{topic}")
                except Exception:
                    pass

                return jsonify({
                    "success": True,
                    "type": "animation",
                    "format": "html",
                    "html": anim_html,
                    "image_url": image_url
                })
            elif result.status_code == 200 and task_status == 'PENDING':
                # TODO: 应轮询等待 task_status 变为 SUCCEEDED，当前直接回退到图片搜索
                print("[visual_module] Qwen 任务未完成(PENDING)，回退到图片搜索")
            else:
                print(f"[visual_module] Qwen-Image生成失败: status={getattr(result, 'status_code', 'unknown')}, output={getattr(result, 'output', 'None')}")
        except Exception as e:
            print(f"[visual_module] Qwen-Image调用异常: {e}")

    # 图片搜索（如果配置了）
    search_images = []
    image_search_error = None
    if image_search_provider and REQUESTS_AVAILABLE:
        try:
            search_query = topic
            if image_search_provider.lower() == "unsplash":
                url = f"https://api.unsplash.com/search/photos?query={requests.utils.quote(search_query)}&per_page=3"
                headers = {"Authorization": f"Client-ID {image_search_api_key}"} if image_search_api_key else {}
                response = requests.get(url, headers=headers, timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    if data.get("results"):
                        # 修复 M3：防御 KeyError，跳过无有效 URL 的图片
                        for img in data["results"][:3]:
                            if not isinstance(img, dict):
                                continue
                            urls = img.get("urls", {})
                            regular_url = urls.get("regular")
                            if regular_url:
                                search_images.append(regular_url)
                elif response.status_code == 401:
                    image_search_error = "图片搜索 API Key 无效"
                elif response.status_code == 429:
                    image_search_error = "图片搜索请求频率超限"
                else:
                    image_search_error = f"图片搜索返回错误 {response.status_code}"
            elif image_search_provider.lower() == "pexels":
                url = f"https://api.pexels.com/v1/search?query={requests.utils.quote(search_query)}&per_page=3"
                headers = {"Authorization": image_search_api_key} if image_search_api_key else {}
                response = requests.get(url, headers=headers, timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    if data.get("photos"):
                        # 修复 M3：防御 KeyError，跳过无有效 URL 的图片
                        for img in data["photos"][:3]:
                            if not isinstance(img, dict):
                                continue
                            src = img.get("src", {})
                            medium_url = src.get("medium")
                            if medium_url:
                                search_images.append(medium_url)
                elif response.status_code == 401:
                    image_search_error = "图片搜索 API Key 无效"
                elif response.status_code == 429:
                    image_search_error = "图片搜索请求频率超限"
                else:
                    image_search_error = f"图片搜索返回错误 {response.status_code}"

            if search_images:
                print(f"[visual_module] 搜索到{len(search_images)}张图片")
        except Exception as e:
            image_search_error = f"图片搜索失败: {str(e)}"
            print(f"[visual_module] 图片搜索失败: {e}")
    elif image_search_provider and not REQUESTS_AVAILABLE:
        image_search_error = "图片搜索功能不可用（requests 库未安装）"
        print("[visual_module] 图片搜索跳过：requests 库未安装")

    image_prompt = ""
    if search_images:
        image_prompt = f"""

# 可用图片素材（在动画中合理使用这些图片）
以下是搜索到的相关图片URL，请在动画中使用<img>标签引用：
{chr(10).join([f"- 图片{i+1}: {url}" for i, url in enumerate(search_images)])}

使用示例：<img src="图片URL" style="width:100px;height:auto;" />
"""

    # 特殊主题：勾股定理（精确几何演示）
    # 修复 L2：下方动画 prompt 为超长硬编码字符串，建议后续抽取为外部模板文件或模块级常量以便维护
    if "勾股定理" in topic:
        anim_prompt = f"""请为数学课程"{topic}"生成一个科学准确的几何演示动画HTML代码。

# 勾股定理可视化（a² + b² = c²）

## 核心几何原理
直角三角形两直角边分别为a和b，斜边为c，则 a² + b² = c²。
动画必须清晰展示：两个直角边上的正方形面积之和等于斜边上的正方形面积。

## 构图要求

### 1. 直角三角形（底部中央）
- 放置在坐标系第一象限
- 直角顶点在原点(0,0)附近
- 底边a沿x轴方向（蓝色）
- 高边b沿y轴方向（红色）
- 斜边c（黑色粗线）
- 直角标记（小正方形或直角符号）

### 2. 三个正方形（必须准确构造）
- **a边上的正方形**：沿x轴，边长=a，蓝色填充，面积标注a²
- **b边上的正方形**：沿y轴，边长=b，红色填充，面积标注b²
- **c边上的正方形**：以斜边c为一条边的正方形，绿色填充，面积标注c²

### 3. c边上正方形的几何构造（非常重要！必须按向量法构造）
设斜边端点为A(x1,y1)和C(x2,y2)：
- 向量 v = (x2-x1, y2-y1)
- 边长 c = sqrt((x2-x1)² + (y2-y1)²)
- 垂直向量 n = (-v.y, v.x) = (y1-y2, x2-x1)
- 正方形四个顶点：A, C, C+n, A+n
- 使用SVG <polygon>或<path>绘制
- **禁止**使用"以中点为中心旋转的轴对齐正方形"的错误方法

## 动画效果（CSS动画，无限循环）
1. **面积闪烁**：三个正方形依次高亮闪烁（a²→b²→c²→a²），配合文字标注
2. **面积累加**：a²和b²的面积块依次"移动"到c²位置，演示a²+b²=c²
3. **勾股公式显示**：底部公式"a² + b² = c²"逐字显示或闪烁高亮
4. **边长标注**：a、b、c三边长依次标注，箭头指示

## 尺寸与布局
- viewBox="0 0 500 320"
- 直角三角形底边a=120px，高b=90px（3:4:5比例），斜边c=150px
- 直角顶点坐标(80, 200)
- 底边终点(200, 200)
- 高边终点(80, 110)
- 斜边端点(80, 110)到(200, 200)
- 文字标注清晰，字体大小12-14px

## 色彩规范
- 底边a及正方形：蓝色 #3498DB
- 高边b及正方形：红色 #E74C3C
- 斜边c及正方形：绿色 #27AE60
- 文字：深灰 #333
- 背景：透明

## 格式规范
1. 输出纯HTML代码，包含在<div class="classroom-animation">中
2. CSS放在<style>标签内，SVG直接内联
3. 动画自动循环（animation-iteration-count: infinite）
4. 背景透明
5. 禁止使用JavaScript
6. SVG必须使用精确的几何计算，不能有视觉误差

## 几何验证检查
1. 三个正方形的面积是否满足a²+b²=c²？
2. c边上的正方形是否以斜边为一条边？（不是旋转的轴对齐正方形）
3. 直角标记是否清晰？
4. 公式"a² + b² = c²"是否准确标注？

直接输出HTML代码，不要其他解释。
""" + image_prompt
    else:
        anim_prompt = f"""请为理科课程主题"{topic}"生成一个具有教学逻辑性的课堂演示动画HTML代码。

# 角色定义
你是一位擅长用SVG+CSS动画演示理科教学逻辑的专业课件动画师。你必须使用内联SVG绘制精细的教学图形，配合CSS动画展示知识点的逻辑关系、过程演变或因果关系。本动画主要服务于理科教学（物理、化学、数学、生物）。

# 核心原则（必须遵守）
- **必须使用SVG内联绘图**来绘制教学图形（分子结构、电路图、几何图形、实验装置、生物细胞等）
- SVG图形必须精细、美观、专业，像教科书插图一样清晰
- **动画必须具有教学逻辑性**，展示因果关系、演变过程、步骤流程或对比关系
- 动画是理科教学演示工具，必须能帮助学生理解"{topic}"的核心概念
- 逻辑清晰：从起点→过程→结果，或从问题→分析→结论的完整逻辑链条
- 背景必须透明，便于嵌入PPT

# SVG绘图要求（非常重要！）
- 所有的教学图形、实验装置、分子结构、几何图形等必须用内联<svg>标签绘制
- SVG必须设置viewBox属性，推荐viewBox="0 0 500 320"
- SVG元素使用stroke和fill设置颜色，stroke-width设置线宽
- 用<circle>画圆、<rect>画矩形、<line>画直线、<path>画曲线、<polygon>/<polyline>画多边形
- 用<text>添加标注文字，设置font-size和fill颜色
- 用<g>标签分组，配合CSS动画让整个组动起来
- 用<defs>和<marker>定义箭头等可复用元素
- SVG图形要专业精细，线条流畅，颜色协调，不要画简陋的火柴人

# 理科SVG绘图示例（主要服务以下学科）

**物理**：用SVG画弹簧+方块（弹簧用<path>的正弦曲线）、斜面+方块、电路图（<rect>电阻+<line>导线+<circle>电池）、光的折射（<line>入射光+折射光+<rect>界面）、磁场线（<path>曲线+箭头）、波动图（<path>正弦波）

**化学**：用SVG画分子结构（<circle>原子+<line>化学键）、烧杯试管（<rect>+<path>液面）、电子云（<circle>不同透明度的圆叠加）、反应方程式（<text>+上下标）、原子结构（<circle>原子核+<ellipse>电子轨道）

**数学**：用SVG画坐标系（<line>坐标轴+<text>刻度）、函数曲线（<path>贝塞尔曲线）、几何图形（<polygon>+标注）、数列变化（<rect>柱状图+动画）、导数/积分示意（<path>曲线+<rect>面积）

**生物**：用SVG画细胞结构（<circle>细胞膜+<circle>细胞核+<rect>线粒体）、DNA双螺旋（<path>两条螺旋线+<line>碱基对）、光合作用（<circle>叶绿体+<path>光能+<text>产物）、细胞分裂（<circle>动态分裂过程）

# 理科教学逻辑动画类型

【类型1：过程演示型】适用于物理实验、化学反应、生物生长等
示例结构：SVG绘制起点状态 → CSS动画过渡 → SVG绘制终态（循环）
例：化学反应H₂+O₂→H₂O、细胞分裂过程、电路通电过程

【类型2：因果关系型】适用于物理原理、化学规律、生物机制等
示例结构：SVG绘制原因/条件 → 动画展示作用机制 → SVG绘制结果/现象
例：力→加速度→速度变化、温度升高→分子运动加快→状态改变

【类型3：对比演示型】适用于概念辨析、正确vs错误、变量对比等
示例结构：SVG左侧画A + SVG右侧画B → 动态高亮差异
例：光合作用vs呼吸作用、串联vs并联电路、酸性vs碱性

【类型4：循环系统型】适用于物质循环、能量循环、生物循环等
示例结构：SVG画各环节 → 依次高亮激活 → 形成循环回路
例：碳循环、水循环、能量流动、细胞呼吸链

【类型5：层级展开型】适用于知识结构、分类体系、公式推导等
示例结构：SVG从中心向外逐层展开
例：生物分类树、数学公式推导步骤、物质分类体系

# 格式规范（必须100%遵守）
1. 输出纯HTML代码，不要markdown代码块包裹
2. 代码必须包含在 <div class="classroom-animation"> 中
3. CSS样式放在 <style> 标签内，SVG图形直接内联
4. 动画必须是自动循环播放的（animation-iteration-count: infinite）
5. 背景色必须设为透明（background: transparent）
6. 整体尺寸限制在 500px 宽 × 320px 高以内
7. 文字使用中文，简短明确（关键词≤6字），字体大小12-16px
8. 可以使用提供的图片素材（通过<img>标签引用），禁止使用其他外部字体、JS库资源
9. 禁止使用JavaScript，仅用CSS animation/transition实现动画
10. SVG图形必须标注教学含义（如受力标注F、加速度标注a等）
11. 逻辑链条必须清晰可见，用动画顺序或位置关系体现因果/过程

# 正例格式锚点

【示例：牛顿第二定律因果链】
<div class="classroom-animation" style="width:500px;height:320px;...">
  <style>
    @keyframes pushForce {{ from {{ transform: translateX(0); }} to {{ transform: translateX(30px); }} }}
    @keyframes moveBlock {{ from {{ transform: translateX(0); }} to {{ transform: translateX(80px); }} }}
    .force-arrow {{ animation: pushForce 2s ease-in-out infinite alternate; }}
    .block {{ animation: moveBlock 2s ease-in-out infinite alternate; }}
  </style>
  <svg viewBox="0 0 500 320" width="500" height="320">
    <!-- 地面 -->
    <line x1="20" y1="220" x2="480" y2="220" stroke="#888" stroke-width="2"/>
    <!-- 方块 -->
    <g class="block">
      <rect x="120" y="170" width="60" height="50" fill="#4A90D9" rx="4"/>
      <text x="150" y="200" text-anchor="middle" fill="white" font-size="14">m</text>
    </g>
    <!-- 力的箭头 -->
    <g class="force-arrow">
      <line x1="60" y1="195" x2="115" y2="195" stroke="#E74C3C" stroke-width="3" marker-end="url(#arrowhead)"/>
      <text x="85" y="185" text-anchor="middle" fill="#E74C3C" font-size="14" font-weight="bold">F</text>
    </g>
    <!-- 箭头定义 -->
    <defs><marker id="arrowhead" markerWidth="10" markerHeight="7" refX="10" refY="3.5" orient="auto"><polygon points="0 0, 10 3.5, 0 7" fill="#E74C3C"/></marker></defs>
    <!-- 公式标注 -->
    <text x="250" y="280" text-anchor="middle" fill="#333" font-size="16" font-weight="bold">F = ma</text>
  </svg>
</div>

【示例：化学反应过程】
<div class="classroom-animation" style="...">
  <style>
    @keyframes bond {{ from {{ opacity:1; }} to {{ opacity:0; }} }}
    @keyframes newBond {{ from {{ opacity:0; }} to {{ opacity:1; }} }}
    .old-bond {{ animation: bond 3s ease-in-out infinite; }}
    .new-bond {{ animation: newBond 3s ease-in-out 1.5s infinite; }}
  </style>
  <svg viewBox="0 0 500 320" width="500" height="320">
    <g class="old-bond">
      <circle cx="150" cy="160" r="25" fill="#E74C3C"/><text x="150" y="165" text-anchor="middle" fill="white" font-size="14">H₂</text>
      <circle cx="350" cy="160" r="25" fill="#3498DB"/><text x="350" y="165" text-anchor="middle" fill="white" font-size="14">O₂</text>
    </g>
    <text x="250" y="165" text-anchor="middle" fill="#333" font-size="20">→</text>
    <g class="new-bond">
      <circle cx="250" cy="160" r="30" fill="#2ECC71"/><text x="250" y="165" text-anchor="middle" fill="white" font-size="14">H₂O</text>
    </g>
    <text x="250" y="270" text-anchor="middle" fill="#333" font-size="14">2H₂+O₂→2H₂O</text>
  </svg>
</div>

# 反例（禁止生成）
- 用CSS div方块画教学图形（× 必须用SVG画精细图形）
- 纯装饰性动画：花瓣飘落、星星闪烁（× 无教学意义）
- 简陋火柴人/粗糙贴图式图形（× SVG要精细专业）
- 纯文字静态展示（× 必须有动画效果）
- 包含JavaScript代码（× 纯CSS动画）
- 背景不透明（× 必须透明）
- 使用未提供的外部图片URL（× 只能使用提供的图片素材）
- 文科类动画（× 本模块专注理科：物理/化学/数学/生物）

# 动画设计检查清单
1. 是否使用了SVG绘制精细教学图形？
2. 这个动画展示了"{topic}"的什么理科核心逻辑？
3. 学生看完能理解哪个知识点？
4. SVG图形是否专业清晰，像教科书插图？
5. 逻辑链条是否清晰（起点→过程→结果）？

直接输出HTML代码，不要其他解释。
""" + image_prompt

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": anim_prompt}],
        temperature=0.7,
        max_tokens=4096
    )
    # 修复 M2：AI 响应空校验，避免 choices 为空时抛 IndexError
    if not response or not response.choices:
        return jsonify({"success": False, "error": "AI 返回为空，请重试"}), 500
    anim_html = response.choices[0].message.content.strip()
    if not anim_html:
        return jsonify({"success": False, "error": "AI 返回内容为空"}), 500

    anim_html = re.sub(r'^```\w*\s*', '', anim_html)
    anim_html = re.sub(r'\s*```$', '', anim_html)

    # 修复 H1：调用增强版 HTML 清理函数，防御 XSS（覆盖危险标签、on* 事件、javascript: 变形、expression()、data:text/html 等）
    anim_html = _sanitize_anim_html(anim_html)

    if 'classroom-animation' not in anim_html:
        safe_topic = html_escape(topic)
        anim_html = f'<div class="classroom-animation" style="width:500px;height:320px;background:transparent;display:flex;align-items:center;justify-content:center;color:#333;font-size:18px;border:2px dashed #ccc;border-radius:12px;">{safe_topic}</div>'

    try:
        database.save_search_history(user['id'], f"可视化生成-教学动画：{topic}")
    except Exception as e:
        print(f"[visual_module] 保存搜索历史失败: {e}")

    return jsonify({
        "success": True,
        "type": "animation",
        "format": "html",
        "html": anim_html
    })
