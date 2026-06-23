# AI 备课助手

基于 Flask + OpenAI API 的智能备课工具，帮助教师快速生成教案、课件大纲、习题试卷等教学资源。

## 功能特点

- 智能教案设计：自动生成教学目标、重难点、教学过程等
- 课件大纲生成：PPT 结构规划，每页标题+要点+配图建议
- 分层习题试卷：基础/提高/拓展三级难度，附答案解析
- 教学资源整理：复习提纲、导学案、知识点总结、实验方案
- 实时流式输出：思考过程可视化，内容逐字显示
- Word 文档导出：一键下载生成的教案为 Word 格式

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/MewoStar/LSW.git
cd LSW
```

### 2. 配置 API Key

```bash
# 复制示例配置文件
cp config.yaml.example config.yaml

# 编辑 config.yaml，填入你的 API Key
```

支持以下 API：
- **智谱 AI**（推荐，有免费额度）：`https://open.bigmodel.cn/api/paas/v4`
- **DeepSeek**：`https://api.deepseek.com/v1`
- **本地 Ollama**：`http://localhost:11434/v1`

### 3. 安装依赖

```bash
pip install flask openai pyyaml python-docx
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

## 项目结构

```
ai备课助手1.0/
├── web_app.py              # Flask 主应用
├── config.yaml             # 配置文件（需自行创建）
├── config.yaml.example     # 配置示例
├── templates/
│   └── index.html          # 前端页面
├── 启动备课助手.bat        # Windows 启动脚本
├── 启动备课助手.ps1        # PowerShell 启动脚本
└── README.md               # 项目说明
```

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

1. **config.yaml 包含敏感信息，已加入 .gitignore，请勿提交到仓库**
2. 生成的教学资源文件（.md、.docx）不会上传到 GitHub
3. 会话数据存储在内存中，重启服务后会丢失

## 技术栈

- **后端**：Python + Flask
- **前端**：HTML + CSS + JavaScript（原生）
- **AI API**：OpenAI 兼容接口（智谱/DeepSeek/Ollama）
- **文档生成**：python-docx

## License

MIT
