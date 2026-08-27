---
name: ppt-word-gen
description: 通过 MCP 协议驱动，自动生成可编辑的 PPTX 演示文稿和结构化 DOCX 报告的全功能服务。
---

# PPT / Word 文档生成服务

ppt-word-gen 是一个 MCP Server，提供 15 个工具和 4 个资源，支持：

- **PPT 演示文稿生成**：根据主题/素材自动生成可编辑 PPTX
- **Word 报告生成**：基于证据材料生成结构化 DOCX（交付/验证/手册/技术分析）
- **文件上传与管理**：支持素材上传、模板上传
- **业务模板管理**：注册和复用预设的生成规范
- **修订迭代**：在已生成文档基础上修改生成新版本

---

## MCP 工具一览

### 信息查询

| 工具名 | 功能说明 |
|---|---|
| `list_generation_profiles` | 查询支持的文档类型、格式、上传限制等全局配置信息 |

### 文件上传

| 工具名 | 功能说明 |
|---|---|
| `upload_file` | 通过 MCP 直接上传小文件（Base64 编码，受 MCP_INLINE_UPLOAD_MB 限制） |
| `create_upload_ticket` | 为大文件创建带签名的 PUT 上传地址，客户端通过 HTTP PUT 上传 |

> 上传时需指定 `purpose`：`source`（素材）或 `reference_template`（参考模板）

### PPT 生成

| 工具名 | 功能说明 |
|---|---|
| `generate_presentation` | 根据主题、页数、风格等参数，生成可编辑 PPTX |
| `revise_presentation` | 在已成功生成的 PPTX 基础上，按修订指令生成新版本 |

### Word 报告生成

| 工具名 | 功能说明 |
|---|---|
| `preview_word_report_format` | 预览 Word 报告的默认格式（字体、字号、行距、缩进、编号样式等），支持自定义覆盖 |
| `generate_word_report` | 用户确认格式后，创建 DOCX 生成任务（需传入确认凭证） |
| `revise_word_report` | 在已成功生成的 DOCX 基础上，按修订指令生成新版本 |

### 任务管理

| 工具名 | 功能说明 |
|---|---|
| `get_generation_task` | 查询 PPT 或 Word 异步任务的当前状态 |
| `wait_generation_task` | 短轮询等待任务完成（最长 55 秒），超时后返回最新进度 |
| `cancel_generation_task` | 取消排队中或运行中的任务 |
| `get_artifact` | 任务成功后获取带 24 小时签名的文件下载链接（ResourceLink） |

### 业务模板管理

| 工具名 | 功能说明 |
|---|---|
| `list_business_templates` | 列出已注册的业务模板，可按文档类型过滤 |
| `register_business_template` | 注册新的业务模板（预设生成指令和报告类型） |
| `delete_business_template` | 删除自定义业务模板（内置模板不可删除） |

---

## MCP 资源一览

| 资源 URI | 名称 | 说明 |
|---|---|---|
| `ppt-word://rules/workflow` | generation-workflow | PPT/Word 生成完整工作流指引 |
| `ppt-word://rules/presentation` | presentation-rules | PPT 演示文稿内容规范 |
| `ppt-word://rules/word-report` | word-report-rules | Word 报告内容规范 |
| `ppt-word://templates/catalog` | business-template-catalog | 业务模板目录（JSON 格式） |

---

## 标准使用流程

### 流程一：生成 PPT

```
1. [可选] upload_file / create_upload_ticket
   └─ 上传素材文件（purpose="source"）

2. generate_presentation
   ├─ topic: "AI 仿真平台架构介绍"
   ├─ page_count: 10
   ├─ style: "简洁商务风格，蓝灰色调"
   ├─ canvas_format: "ppt169"
   └─ source_upload_id: "（如有上传素材）"
   → 返回 task_id

3. wait_generation_task(task_type="presentation", task_id=...)
   → 等待任务完成

4. get_artifact(task_type="presentation", task_id=...)
   → 获取 .pptx 下载链接（24 小时有效）

5. [可选] revise_presentation(source_task_id=..., instructions="修改第3页标题")
   → 在上一版基础上修订
```

### 流程二：生成 Word 报告

```
1. [可选] upload_file / create_upload_ticket
   ├─ 上传素材文件（purpose="source"）
   └─ 上传参考模板（purpose="reference_template"，仅 .docx/.dotx）

2. preview_word_report_format
   ├─ template_upload_id: "（如有上传模板）"
   ├─ body_font: "宋体"    ← 可选覆盖
   ├─ body_size_pt: 12     ← 可选覆盖
   └─ ...更多格式参数
   → 返回格式摘要 + confirmation_token
   → 将格式摘要展示给用户确认

3. [用户确认格式]

4. generate_word_report
   ├─ format_confirmation_token: "（上一步返回的凭证）"
   ├─ instructions: "生成 XX 仿真验证报告..."
   ├─ report_type: "validation"
   ├─ title / project_name / author / document_version
   └─ source_upload_id: "（如有上传素材）"
   → 返回 task_id

5. wait_generation_task(task_type="word_report", task_id=...)
   → 等待任务完成

6. get_artifact(task_type="word_report", task_id=...)
   → 获取 .docx 下载链接（24 小时有效）

7. [可选] revise_word_report(source_task_id=..., instructions="补充第四章结论")
   → 在上一版基础上修订
```

---

## Word 报告类型

| 类型 | 适用场景 |
|---|---|
| `delivery` | 交付报告：项目范围、Agent 能力、架构、部署、验收证据、运维指引 |
| `validation` | 验证报告：背景目标、指标、环境版本、执行过程、结果对比、精度结论 |
| `manual` | 操作手册：安装访问、配置说明、操作流程、输入输出、故障排查 |
| `technical` | 技术分析：问题陈述、架构方法、证据分析、权衡风险、建议 |

---

## Word 格式可编辑项

通过 `preview_word_report_format` 可自定义以下排版参数：

| 参数 | 说明 | 示例值 |
|---|---|---|
| `body_font` | 正文字体 | 宋体 |
| `body_size_pt` | 正文字号（磅） | 12 |
| `line_spacing` | 行距倍数 | 1.5 |
| `first_line_indent_chars` | 首行缩进（字符数） | 2 |
| `heading1_font` | 一级标题字体 | 黑体 |
| `heading1_size_pt` | 一级标题字号 | 22 |
| `heading2_font` | 二级标题字体 | 黑体 |
| `heading2_size_pt` | 二级标题字号 | 16 |
| `heading3_font` | 三级标题字体 | 黑体 |
| `heading3_size_pt` | 三级标题字号 | 14 |
| `numbering_style` | 编号样式 | `decimal` 或 `chinese` |

---

## 关键约束

- **异步任务模式**：生成类操作均为异步，提交后需轮询 `wait_generation_task` 等待完成
- **Word 格式确认必须**：生成 Word 报告前，必须先调用 `preview_word_report_format` 获取确认凭证
- **证据原则**：数值、版本、日期、人员、结论等必须来自用户提供的证据，禁止编造
- **幂等性**：支持 `idempotency_key` 防止重复提交
- **签名下载**：文件下载链接带签名，默认 24 小时有效
- **LLM 可覆盖**：生成工具支持传入自定义 model/base_url/api_key/temperature，未传时使用服务端默认配置

---

## 参考文档

- [报告契约规范](references/report-contract.md) — 报告章节结构与输出 Schema 定义
- [验证规则](references/validation-rules.md) — 报告生成合规性检查规则
