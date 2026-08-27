# ppt-word-gen

> **AI 驱动的可编辑 PPTX 与标准化 Word 仿真测试报告生成服务**  
> 单一轻量服务在同一端口提供 **MCP (Model Context Protocol) Server**、**REST API** 与 **Web Demo**。

---

## 🚀 MCP (Model Context Protocol) 服务

`ppt-word-gen` 原生实现标准 MCP 协议（支持 Streamable HTTP 与 STDIO 双传输通道），赋能各类 Agent 客户端（如 **Claude Code**, **Claude Desktop**, **Cursor**, **Cline**, **OpenWebUI** 等）自主完成文档格式提取、用户确认、异步排队生成、语义多轮修订与安全文件下载。

### 1. 客户端接入配置 (`.mcp.json`)

项目根目录已内置 [`.mcp.json`](.mcp.json)，Agent 工具打开本目录即可自动发现并加载服务。

#### 方式 A：HTTP 传输模式（推荐，适用于容器部署 / 局域网服务 / 本机服务）

在项目根目录创建或编辑 `.mcp.json`：

```json
{
  "mcpServers": {
    "ppt-word-gen": {
      "type": "http",
      "url": "http://127.0.0.1:8000/mcp"
    }
  }
}
```

若服务开启了 Bearer Token 鉴权（配置了 `PPT_WORD_GEN_TOKEN`），可添加请求头：

```json
{
  "mcpServers": {
    "ppt-word-gen": {
      "type": "http",
      "url": "http://YOUR_SERVER_IP:8000/mcp",
      "headers": {
        "Authorization": "Bearer ${PPT_WORD_GEN_TOKEN}"
      }
    }
  }
}
```

使用 Claude Code CLI 一键注册：

```bash
# 项目级注册（写入当前目录 .mcp.json）
claude mcp add --transport http --scope project ppt-word-gen http://127.0.0.1:8000/mcp

# 用户全局注册
claude mcp add --transport http --scope user ppt-word-gen http://127.0.0.1:8000/mcp
```

#### 方式 B：STDIO 传输模式（本机免起 HTTP 端口直接调用）

在客户端配置中直接通过 Python 进程启动：

```json
{
  "mcpServers": {
    "ppt-word-gen": {
      "command": "python",
      "args": ["-m", "ppt_word_gen.mcp_stdio"]
    }
  }
}
```

---

### 2. MCP 核心协作工作流

```text
[用户输入 / 来源资料 / 模板]
             │
             ▼
[可选：upload_file / create_upload_ticket 暂存来源文件与 Word 模板]
             │
             ▼
[preview_word_report_format 提取模板默认格式（字体、字号、行距、多级标题、编号规范）]
             │
             ▼
[向用户展示预览并获取确认 ➔ 获得绑定格式哈希的 confirmation_token]
             │
             ▼
[generate_presentation / generate_word_report 发起异步生成任务]
             │
             ▼
[wait_generation_task / get_generation_task 轮询进度状态]
             │
             ▼
[get_artifact 获取 24 小时有效的安全签名 ResourceLink 下载链接]
             │
             ▼
[可选：revise_presentation / revise_word_report 进行多轮语义增量修订]
```

---

### 3. MCP 工具全集（15 个 Tools）

| 工具名称 | 功能描述 | 关键参数 |
| :--- | :--- | :--- |
| `list_generation_profiles` | 列出可生成的文档类型、支持格式与二进制上传约定 | 无 |
| `preview_word_report_format` | 提取模板默认格式（或内置模板）并应用修改；返回格式信息与 `confirmation_token` | `template_upload_id`, `custom_format` |
| `generate_presentation` | 创建可编辑 PPTX 异步生成任务 | `prompt`, `format`, `source_upload_ids`, `model` 等 |
| `generate_word_report` | 提交 Word 报告任务（**必须携带** `confirmation_token`） | `title`, `report_type`, `confirmation_token`, `source_upload_ids` |
| `get_generation_task` | 查询 PPT 或 Word 异步任务的阶段、百分比进度、错误与产物状态 | `task_id`, `task_type` |
| `wait_generation_task` | 最长等待 55 秒，任务完成立即返回，超时则返回最新状态 | `task_id`, `task_type`, `timeout_seconds` |
| `cancel_generation_task` | 请求取消尚未结束的异步生成任务 | `task_id`, `task_type` |
| `get_artifact` | 任务完成后获取带标准 ResourceLink 的 24 小时安全签名下载直链 | `task_id`, `task_type` |
| `upload_file` | Base64 内联上传小文件（不超过 5MB） | `filename`, `content_base64`, `purpose` |
| `create_upload_ticket` | 为大文件创建一次性 PUT 上传地址（避免大内容占用 Agent 上下文） | `filename`, `purpose`, `max_bytes` |
| `list_business_templates` | 列出内置与企业自定义的 PPT/Word 业务规范模板 | 无 |
| `register_business_template` | 注册企业业务内容与风格规范模板 | `template_id`, `name`, `spec` |
| `delete_business_template` | 删除自定义业务模板（内置模板受保护不可删除） | `template_id` |
| `revise_presentation` | 将已有 PPTX 作为上下文，按自然语言指令生成新版本 | `parent_task_id`, `revision_prompt` |
| `revise_word_report` | 用户确认格式后，将已有 DOCX 作为基础增量生成新版本 | `parent_task_id`, `confirmation_token`, `revision_instructions` |

### 4. MCP 规则与资产资源（4 个 Resources）

Agent 可直接读取规范指南与设计约束：
- `ppt-word://rules/workflow`：MCP 完整交互与格式确认规范
- `ppt-word://rules/presentation`：PPT Master 设计与排版规范
- `ppt-word://rules/word-report`：仿真与工程 Word 报告排版标准
- `ppt-word://templates/catalog`：业务模板目录与规范速查

---

## 📁 项目结构

```text
ppt-word-gen/
├── ppt_word_gen/              # Python 核心业务包
│   ├── app.py                 # FastAPI 入口（Demo、REST API、MCP Streamable HTTP）
│   ├── mcp_server.py          # MCP Server 实现（15 工具、4 资源、双向协议）
│   ├── mcp_stdio.py           # MCP STDIO 运行入口
│   ├── tasks.py               # PPT 异步任务队列与状态机
│   ├── report_tasks.py        # Word 异步任务队列与状态机
│   ├── report_agent.py        # Word 报告生成 Agent
│   ├── report_documents.py    # Word 文档排版、格式继承与 GBK 修复引擎
│   ├── word_format.py         # Word 格式解析、修改与防篡改 Token 校验
│   ├── signed_tokens.py       # HMAC-SHA256 签名下载与上传票据
│   ├── pptmaster.py           # PPT Master 脚本适配与执行
│   ├── upload_store.py        # 临时文件上传与生命周期管理
│   ├── task_store.py          # SQLite 持久化任务存储
│   └── config.py              # 全局配置读取
├── assets/word_templates/     # 内置 Word 模板（CID629 电驱系统联合仿真 v1.5）
├── skills/ai-simulation-report/# 仿真报告标准契约与校验规则
├── static/                    # 前端 Web Demo 页面
├── tests/                     # 41 项单元测试与集成测试
├── Dockerfile                 # 容器构建文件（支持联网/离线两种模式）
├── docker-compose.yml         # 生产/测试容器编排配置
├── build-overlay.ps1          # 企业级解密与离线 wheel 构建脚本
├── requirements.txt           # 依赖清单
├── .env.example               # 环境变量参考模板
└── .mcp.json                  # MCP 配置文件
```

---

## 💻 本地快速启动

### 1. 安装依赖与执行测试

```powershell
# 安装依赖
python -m pip install -r requirements.txt

# 运行全套 41 项自动化测试
python -m unittest discover -s tests -v
```

### 2. 启动服务

```powershell
# 启动 HTTP + MCP 服务（默认监听 0.0.0.0:8000）
python -m ppt_word_gen
```

- **Web Demo 页面**：`http://127.0.0.1:8000/demo`
- **OpenAPI 交互文档**：`http://127.0.0.1:8000/docs`
- **MCP HTTP 端点**：`http://127.0.0.1:8000/mcp`
- **健康检查**：`http://127.0.0.1:8000/health`

---

## 🐳 Docker / WSL 容器化部署

### 1. 标准部署

在 Ubuntu / WSL 终端中：

```bash
cd /path/to/ppt-word-gen
cp .env.example .env.compose

# 按需修改 .env.compose（接入真实模型时设置 MOCK_LLM=0 并填写 API Key）
sudo service docker start
docker compose --env-file .env.compose up -d --build --wait
```

### 2. 企业离线 Wheel 构建

在具备企业数据防泄露（DLP）或内网隔离的环境下，在 Windows PowerShell 中执行：

```powershell
pwsh ./build-overlay.ps1
```

该脚本将自动通过文件流安全复制并调用 `OFFLINE_INSTALL=1` 进行本地 Wheel 离线构建。

### 3. 服务状态检查与停止

```bash
# 查看容器状态与健康指标
docker compose ps
curl http://127.0.0.1:8000/health

# 停止容器（数据保留在 ppt-word-gen-data 卷中）
docker compose stop
```

---

## ⚙️ 环境变量说明

复制 `.env.example` 为 `.env` 或 `.env.compose`：

| 变量名 | 默认值 | 说明 |
| :--- | :--- | :--- |
| `LLM_BASE_URL` | `https://dashscope.aliyuncs.com/compatible-mode/v1` | OpenAI 兼容接口 Base URL |
| `LLM_API_KEY` | 无 | 大语言模型 API Key |
| `LLM_MODEL` | `qwen3.7-plus` | 默认创作模型名称 |
| `MOCK_LLM` | `1` | `1` 无需 API Key 即可全流程自测演示；接入真实模型改为 `0` |
| `PPT_WORD_GEN_TOKEN` | 留空 | Bearer Token 鉴权，留空则无需认证 |
| `PUBLIC_BASE_URL` | `http://127.0.0.1:8000` | 产物下载与大文件上传的基础 URL（局域网调用请设为服务器内网 IP） |
| `DOWNLOAD_SIGNING_SECRET` | 留空（自动生成） | HMAC 签名密钥，用于生成 24 小时临时下载凭证 |
| `MAX_CONCURRENT_TASKS` | `2` | PPT 最大并发处理任务数 |
| `MAX_CONCURRENT_REPORT_TASKS` | `1` | Word 最大并发处理任务数 |
| `MCP_INLINE_UPLOAD_MB` | `5` | MCP Base64 内联上传上限（MB） |
| `MAX_UPLOAD_MB` | `20` | REST / 大文件 Ticket 上传上限（MB） |

---

## 🔌 核心 REST API

| 请求方法 | 路径 | 描述 |
| :--- | :--- | :--- |
| `GET` | `/health` | 服务健康检查（含 SQLite 存储与任务队列状态） |
| `GET` | `/demo` | 内置 Web 演示页面 |
| `GET` | `/api/v1/word-format` | 获取内置 Word 模板默认格式及确认 Token |
| `POST` | `/api/v1/word-format` | 提取自定义模板格式、应用修改并生成确认 Token |
| `POST` | `/api/v1/report-tasks` | 提交 Word 仿真/技术报告生成任务 |
| `GET` | `/api/v1/report-tasks/{id}` | 查询 Word 报告任务进度与详情 |
| `GET` | `/api/v1/report-tasks/{id}/result` | 下载生成的 DOCX 报告 |
| `POST` | `/api/v1/tasks` | 提交 PPT 演示文稿生成任务 |
| `GET` | `/api/v1/tasks/{id}` | 查询 PPT 任务进度与详情 |
| `GET` | `/api/v1/tasks/{id}/result` | 下载生成的 PPTX 文件 |
| `POST` | `/api/v1/uploads` | 暂存二进制来源文件或模板 |
| `PUT` | `/api/v1/upload-tickets/{token}` | 一次性大文件直传接口 |
| `GET` | `/api/v1/artifacts/{token}` | 24 小时 HMAC 签名安全下载接口 |
| `POST` | `/mcp` | MCP Streamable HTTP 协议端点 |

---

## 📄 开源许可

本项目遵循 [MIT License](LICENSE)。
