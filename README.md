# AI 备课助手

基于 Flask + OpenAI API 的智能备课工具，帮助教师快速生成教案、课件大纲、习题试卷等教学资源。

## 功能特点

- 智能教案设计：自动生成教学目标、重难点、教学过程等
- 课件大纲生成：PPT 结构规划，每页标题+要点+配图建议
- 分层习题试卷：基础/提高/拓展三级难度，附答案解析
- 教学资源整理：复习提纲、导学案、知识点总结、实验方案
- 实时流式输出：思考过程可视化，内容逐字显示
- Word 文档导出：一键下载生成的教案为 Word 格式
- 用户认证系统：支持注册、登录、退出登录
- 个人资料管理：修改昵称等账号设置

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/MewoStar/ai-LSW.git
cd ai-LSW
```

### 2. 配置环境变量

```bash
# 复制示例配置文件
cp .env.example .env

# 编辑 .env，填入你的 API Key
```

支持以下 API：
- **智谱 AI**（推荐，有免费额度）：`https://open.bigmodel.cn/api/paas/v4`
- **DeepSeek**：`https://api.deepseek.com/v1`
- **本地 Ollama**：`http://localhost:11434/v1`

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 启动应用

**方式一：双击启动（Windows）**
```bash
启动备课助手.bat
```

**方式二：命令行启动**
```bash
python web_app.py
```

### 5. 访问应用

浏览器打开：`http://localhost:5000`

## 云端部署

### Vercel 部署

1. 在 GitHub 创建仓库并推送代码
2. 登录 [Vercel](https://vercel.com/dashboard)
3. 导入 GitHub 仓库
4. 在 Vercel 设置环境变量
5. 自动构建部署

### 环境变量清单

| 变量名 | 说明 |
|--------|------|
| SUPABASE_URL | Supabase 项目地址 |
| SUPABASE_ANON_KEY | Supabase 匿名密钥 |
| SUPABASE_SERVICE_ROLE_KEY | Supabase 服务端密钥 |
| OPENAI_API_KEY | OpenAI/DeepSeek API 密钥 |
| OPENAI_BASE_URL | API 基础地址 |
| JWT_SECRET_KEY | JWT 签名密钥 |

## 项目结构

```
ai-LSW/
├── api/                     # Vercel Serverless API 路由
│   ├── login.py             # 登录接口
│   ├── register.py          # 注册接口
│   ├── chat.py              # AI 聊天接口
│   ├── history.py           # 历史记录接口
│   ├── health.py            # 健康检查接口
│   └── user.py              # 用户管理接口
├── app/                     # 核心业务逻辑
│   ├── config.py            # 配置管理
│   ├── supabase_client.py   # Supabase 客户端
│   ├── auth.py              # 认证逻辑
│   └── chat_service.py      # 聊天服务
├── templates/               # 前端页面模板
├── vercel.json              # Vercel 部署配置
├── requirements.txt         # Python 依赖
├── .env                     # 环境变量（本地开发）
├── .env.example             # 环境变量示例
└── web_app.py               # 本地 Flask 应用入口
```

## 数据库设计

### search_history 表（搜索历史）

| 字段 | 类型 | 说明 |
|------|------|------|
| id | UUID | 主键 |
| user_id | UUID | 用户 ID（关联 auth.users） |
| query | TEXT | 搜索/提问内容 |
| model | TEXT | 使用的模型名称 |
| created_at | TIMESTAMP | 创建时间 |

### lesson_plans 表（备课记录）

| 字段 | 类型 | 说明 |
|------|------|------|
| id | UUID | 主键 |
| user_id | UUID | 用户 ID |
| title | TEXT | 标题 |
| content | TEXT | 内容 |
| content_type | TEXT | 内容类型 |
| created_at | TIMESTAMP | 创建时间 |

## API 接口

### 认证接口

| 接口 | 方法 | 说明 |
|------|------|------|
| /api/register | POST | 用户注册 |
| /api/login | POST | 用户登录 |
| /api/user/logout | POST | 用户登出 |

### 用户管理接口

| 接口 | 方法 | 说明 |
|------|------|------|
| /api/user/profile | GET | 获取用户资料 |
| /api/user/profile | PUT | 修改用户资料 |

### 聊天接口

| 接口 | 方法 | 说明 |
|------|------|------|
| /api/chat | POST | AI 聊天 |

### 历史记录接口

| 接口 | 方法 | 说明 |
|------|------|------|
| /api/history | GET | 获取历史记录 |
| /api/history/{id} | DELETE | 删除历史记录 |

### 健康检查

| 接口 | 方法 | 说明 |
|------|------|------|
| /api/health | GET | 健康检查 |

## 协作开发

### 分支管理

```bash
# 创建功能分支
git checkout -b feature/xxx

# 提交更改
git add .
git commit -m "feat: xxx"

# 推送到远程
git push origin feature/xxx

# 在 GitHub 上创建 Pull Request 合并到 main
```

### 代码规范

- 提交信息使用英文，格式：`type: description`
- type 类型：`feat`（新功能）、`fix`（修复）、`docs`（文档）、`style`（格式）、`refactor`（重构）、`test`（测试）、`chore`（构建）

## 注意事项

1. **.env 包含敏感信息，已加入 .gitignore，请勿提交到仓库**
2. 生成的教学资源文件（.md、.docx）不会上传到 GitHub
3. 生产环境使用 Supabase 数据库存储数据

## 技术栈

- **后端**：Python + Flask
- **前端**：HTML + CSS + JavaScript（原生）
- **AI API**：OpenAI 兼容接口（智谱/DeepSeek/Ollama）
- **数据库**：Supabase PostgreSQL
- **认证**：Supabase Auth + JWT
- **部署**：Vercel Serverless Functions
- **文档生成**：python-docx

## License

MIT