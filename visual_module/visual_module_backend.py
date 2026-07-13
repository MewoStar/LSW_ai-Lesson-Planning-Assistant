# ============================================
# 可视化生成模块 - 后端代码
# ============================================
# 用法：将本文件中的 register_visual_routes(app, ...) 集成到目标项目即可
# 依赖：Flask、openai、re、json
# ============================================

import re
import json
from flask import jsonify, render_template, redirect, url_for


# --------------------------------------------
# 集成函数：注册可视化模块所有路由
# --------------------------------------------
def register_visual_routes(
    app,
    get_current_user,
    safe_get_json,
    database,
    openai_client_factory,   # 一个返回 OpenAI 客户端的函数
    model_name,              # 使用的模型名称，如 "gpt-4o-mini"
    modules_data,            # 全局模块字典（包含 visualization 条目）
    static_url_path="/static"
):
    """
    注册可视化生成模块的所有路由。
    调用此函数后，会自动添加：
      - GET  /visual                 : 重定向到 /module/visualization
      - GET  /module/visualization   : 渲染模板
      - POST /api/visual/generate    : 生成思维导图 / 教学动画
    """

    @app.route("/visual")
    def visual_redirect():
        return redirect(url_for('module_detail', module_id='visualization'))

    @app.route("/module/<module_id>")
    def module_detail(module_id):
        if module_id not in modules_data:
            return redirect(url_for('index'))
        module = modules_data[module_id]

        if module_id == "visualization":
            user = get_current_user()
            if not user:
                return redirect(url_for('login'))
            all_modules = [{"id": k, "name": v["name"], "emoji": v["emoji"]} for k, v in modules_data.items()]
            return render_template(
                "module_visual.html",
                module=module,
                module_id=module_id,
                model=model_name,
                username=user['username'],
                all_modules=all_modules,
                history_sessions=[]
            )
        return redirect(url_for('index'))

    @app.route("/api/visual/generate", methods=["POST"])
    def api_visual_generate():
        try:
            user = get_current_user()
            if not user:
                return jsonify({"success": False, "error": "请先登录"}), 401

            data = safe_get_json()
            if not data:
                return jsonify({"success": False, "error": "无效请求"}), 400

            topic = data.get("topic", "").strip()
            gen_type = data.get("type", "mindmap")
            layout = data.get("layout", "logic")

            if not topic:
                return jsonify({"success": False, "error": "请输入课程主题"}), 400

            client = openai_client_factory()

            if gen_type == "mindmap":
                return _handle_mindmap(client, user, topic, layout, database, model_name)
            elif gen_type == "animation":
                return _handle_animation(client, user, topic, database, model_name)
            else:
                return jsonify({"success": False, "error": "未知的生成类型"}), 400
        except Exception as e:
            import traceback
            traceback.print_exc()
            return jsonify({"success": False, "error": f"生成失败: {str(e)}"}), 500


# --------------------------------------------
# 思维导图生成（支持 3 种布局：逻辑图 / 组织结构图 / 矩形思维导图）
# --------------------------------------------
def _handle_mindmap(client, user, topic, layout, database, model_name):
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

    prompt = _build_mindmap_prompt(topic, layout_desc, format_hint, layout)

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=2000
    )
    content = response.choices[0].message.content.strip()

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


def _build_mindmap_prompt(topic, layout_desc, format_hint, layout):
    if layout == "tree":
        return f"""请为课程主题"{topic}"生成一个{layout_desc}的Mermaid流程图。

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
        return f"""请为课程主题"{topic}"生成一个{layout_desc}的Mermaid流程图。

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
        return f"""请为课程主题"{topic}"生成一个{layout_desc}的Mermaid流程图。

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


# --------------------------------------------
# 教学动画生成
# --------------------------------------------
def _handle_animation(client, user, topic, database, model_name):
    anim_prompt = f"""请为课程主题"{topic}"生成一个具有教学逻辑性的课堂演示动画HTML代码。

# 角色定义
你是一位擅长用CSS动画演示教学逻辑的专业课件动画师。你的动画必须展示知识点的逻辑关系、过程演变或因果关系，帮助学生理解核心概念。

# 核心原则（必须遵守）
- **动画必须具有教学逻辑性**，展示知识点之间的因果关系、演变过程、步骤流程或对比关系
- 动画是教学演示工具，不是装饰品，必须能帮助学生理解"{topic}"的核心概念
- 展示动态过程：如物理实验过程、化学反应变化、数学推导步骤、历史事件演变等
- 逻辑清晰：从起点→过程→结果，或从问题→分析→结论的完整逻辑链条
- 背景必须透明，便于嵌入PPT或其他场景

# 教学逻辑动画类型（根据主题选择合适类型）

【类型1：过程演示型】
适用于：实验过程、操作步骤、化学反应、生物生长等
示例结构：起点状态 → 变化过程 → 最终状态（循环展示）

【类型2：因果关系型】
适用于：物理原理、历史事件、地理现象等
示例结构：原因/条件 → 作用机制 → 结果/现象

【类型3：对比演示型】
适用于：数学对比、概念辨析、正确vs错误等
示例结构：左侧展示A，右侧展示B，动态对比差异

【类型4：循环系统型】
适用于：生态系统、水循环、能量循环、经济循环等
示例结构：各环节依次激活，形成完整循环回路

【类型5：层级展开型】
适用于：知识结构、分类体系、组织结构等
示例结构：从中心向外逐层展开，展示层级关系

# 学科教学逻辑动画示例（必须模仿这种逻辑性）

**物理/力学**：
- 牛顿定律：小球受力→加速度变化→速度变化→位移变化（因果链）
- 浮力原理：物体入水→排开水量→浮力产生→上浮/下沉（过程）
- 能量守恒：动能→势能→动能的循环转换（循环系统）

**化学**：
- 化学反应：反应物分子→碰撞→化学键断裂重组→产物生成（过程）
- 电解水：水电解→氢气上升→氧气上升→气泡对比（因果+对比）

**数学/几何**：
- 几何证明：已知条件→推导步骤→结论（逻辑链）
- 函数变化：x变化→y变化→曲线移动（因果关系）

**生物**：
- 光合作用：光能→叶绿体→化学反应→氧气+有机物（过程）
- 细胞分裂：染色体复制→排列→分裂→两个子细胞（过程）

**地理**：
- 水循环：蒸发→凝结→降水→径流→蒸发（循环系统）
- 板块运动：板块碰撞→挤压→山脉形成（因果关系）

**历史**：
- 事件演变：背景→导火索→事件爆发→结果→影响（逻辑链）

# 格式规范（必须100%遵守）
1. 输出纯HTML代码，不要markdown代码块包裹
2. 代码必须包含在 <div class="classroom-animation"> 中
3. 所有CSS样式必须内联在 style 属性中或使用 <style> 标签放在div内
4. 动画必须是自动循环播放的（animation-iteration-count: infinite）
5. 背景色必须设为透明（background: transparent）
6. 尺寸限制在 500px 宽 × 320px 高以内
7. 文字使用中文，简短明确（关键词≤6字），字体大小14-18px
8. 禁止使用任何外部图片、字体、JS库资源
9. 纯CSS动画，禁止使用JavaScript
10. 动画元素必须标注教学含义（如"受力"、"加速"、"结果"等关键词）
11. 逻辑链条必须清晰可见，用动画顺序或位置关系体现因果/过程

# 正例格式锚点（必须模仿这种教学逻辑性）

【示例1：牛顿第二定律因果链】
<div class="classroom-animation" style="...">
  <div style="...">F(力)</div>  →动画→ <div style="...">a(加速度)</div> →动画→ <div style="...">v(速度)</div>
  底部标注：F=ma
</div>

【示例2：水循环系统】
<div class="classroom-animation" style="...">
  蒸发(上) →动画→ 凝结(右上) →动画→ 降水(右) →动画→ 径流(下) →循环回蒸发
</div>

【示例3：化学反应过程】
<div class="classroom-animation" style="...">
  H₂O分子 →动画分解→ H₂ + O₂ →动画→ 气泡上升对比
</div>

# 反例（禁止生成）
- 纯装饰性动画：花瓣飘落、星星闪烁（无教学意义）
- 随机动画：与知识点无关的有趣元素
- 纯文字静态展示，没有动画效果
- 动画过于复杂，逻辑链条不清晰
- 包含JavaScript代码
- 背景不透明
- 无法让学生理解"{topic}"的核心逻辑

# 动画设计检查清单（生成前必须确认）
1. 这个动画展示了"{topic}"的什么核心逻辑？
2. 学生看完能理解哪个知识点？
3. 动画元素是否标注了教学含义？
4. 逻辑链条是否清晰（起点→过程→结果）？

直接输出HTML代码，不要其他解释。
"""

    response = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": anim_prompt}],
        temperature=0.7,
        max_tokens=2500
    )
    anim_html = response.choices[0].message.content.strip()

    anim_html = re.sub(r'^```\w*\s*', '', anim_html)
    anim_html = re.sub(r'\s*```$', '', anim_html)

    if 'classroom-animation' not in anim_html:
        anim_html = f'<div class="classroom-animation" style="width:500px;height:320px;background:transparent;display:flex;align-items:center;justify-content:center;color:#333;font-size:18px;border:2px dashed #ccc;border-radius:12px;">{topic}</div>'

    try:
        database.save_search_history(user['id'], f"可视化生成-教学动画：{topic}")
    except Exception:
        pass

    return jsonify({
        "success": True,
        "type": "animation",
        "format": "html",
        "html": anim_html
    })
