# coding=utf-8
"""PPT/Word 生成能力的 MCP 工具层。"""
import asyncio
import base64
import binascii
import functools
import inspect
import json
import logging
import time
import uuid
from pathlib import Path
from typing import Any, Callable, Dict, Generic, Literal, Optional, TypeVar, Union, cast, get_type_hints

from mcp.server import MCPServer
from mcp.types import CallToolResult, ResourceLink, TextContent
from pydantic import BaseModel, ConfigDict, RootModel

from .business_templates import (
    BusinessTemplate,
    BusinessTemplateConflict,
    BusinessTemplateNotFound,
    business_template_store,
)
from .config import (
    MAX_UPLOAD_MB,
    MCP_INLINE_UPLOAD_MB,
    PUBLIC_BASE_URL,
    PUBLIC_URL_PREFIX,
    SIGNED_URL_EXPIRE_SECONDS,
)
from .report_documents import safe_filename, validate_docx_package
from .report_models import REPORT_TYPES, ReportCreate, WordFormatOverrides
from .report_tasks import (
    ReportIdempotencyConflict,
    ReportQueueFull,
    report_manager,
)
from .signed_tokens import sign_claims
from .tasks import IdempotencyConflict, TaskCreate, TaskQueueFull, manager
from .upload_store import UploadNotFound, UploadPurposeMismatch, upload_store
from .word_format import (
    WORD_FORMAT_PROFILE_ID,
    WordFormatConfirmation,
    issue_word_format_confirmation,
    verify_word_format_confirmation,
)


TaskType = Literal["presentation", "word_report"]
DocumentTypeFilter = Literal["all", "presentation", "word_report"]
FINAL_STATUSES = {"success", "failed", "cancelled", "interrupted"}
logger = logging.getLogger("ppt_word_gen.mcp")


class GenerationProfiles(BaseModel):
    presentations: Dict[str, Any]
    word_reports: Dict[str, Any]
    model_overrides: Dict[str, Any]
    uploads: Dict[str, Any]
    workflow: list[str]


class CreatedTask(BaseModel):
    task_id: str
    task_type: TaskType
    status: str
    idempotency_reused: bool
    next_tool: str = "get_generation_task"


class GenerationTask(BaseModel):
    task_id: str
    task_type: TaskType
    status: str
    stage: str = ""
    progress: int = 0
    message: str = ""
    error: Optional[str] = None
    artifact_url: Optional[str] = None
    cancel_requested: bool = False
    created_at: float = 0.0
    updated_at: float = 0.0


class Artifact(BaseModel):
    task_id: str
    task_type: TaskType
    filename: str
    bytes: int
    media_type: str
    download_url: str
    expires_at: int


class UploadReceipt(BaseModel):
    upload_id: str
    filename: str
    purpose: Literal["source", "reference_template"]
    bytes: int
    expires_at: float
    use_in_mcp_as: str


class UploadTicket(BaseModel):
    upload_url: str
    method: str = "PUT"
    content_type: str = "application/octet-stream"
    filename: str
    purpose: Literal["source", "reference_template"]
    max_bytes: int
    expires_at: int


class BusinessTemplateList(BaseModel):
    templates: list[BusinessTemplate]


class DeleteBusinessTemplateResult(BaseModel):
    template_id: str
    deleted: bool


class ToolErrorDetails(BaseModel):
    code: str
    message: str
    retryable: bool = False
    suggested_action: str = ""


class ToolFailure(BaseModel):
    ok: Literal[False] = False
    error: ToolErrorDetails


class MCPToolError(RuntimeError):
    def __init__(self, code: str, message: str, *, retryable: bool = False, action: str = ""):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = retryable
        self.action = action


_Callable = TypeVar("_Callable", bound=Callable[..., Any])
_ResultValue = TypeVar("_ResultValue")


class StructuredToolResult(RootModel[_ResultValue], Generic[_ResultValue]):
    """顶层仍是成功/失败对象，同时满足 MCP outputSchema 必须声明 object。"""

    model_config = ConfigDict(json_schema_extra={"type": "object"})


def _error_result(error: MCPToolError) -> CallToolResult:
    payload = ToolFailure(
        error=ToolErrorDetails(
            code=error.code,
            message=error.message,
            retryable=error.retryable,
            suggested_action=error.action,
        )
    ).model_dump(mode="json")
    return CallToolResult(
        isError=True,
        content=[TextContent(type="text", text=json.dumps(payload, ensure_ascii=False))],
        structuredContent=payload,
    )


def _classify_error(exc: Exception) -> MCPToolError:
    if isinstance(exc, MCPToolError):
        return exc
    if isinstance(exc, (UploadNotFound, UploadPurposeMismatch)):
        return MCPToolError("upload_not_found", str(exc), action="重新上传文件并使用新的 upload_id")
    if isinstance(exc, BusinessTemplateNotFound):
        return MCPToolError("template_not_found", str(exc), action="调用 list_business_templates 获取可用模板")
    if isinstance(exc, BusinessTemplateConflict):
        return MCPToolError("template_conflict", str(exc))
    if isinstance(exc, (TaskQueueFull, ReportQueueFull)):
        return MCPToolError("queue_full", str(exc), retryable=True, action="稍后重试或降低并发")
    if isinstance(exc, (IdempotencyConflict, ReportIdempotencyConflict)):
        return MCPToolError("idempotency_conflict", str(exc), action="更换 idempotency_key")
    if isinstance(exc, (ValueError, binascii.Error)):
        return MCPToolError("invalid_argument", str(exc), action="修正参数后重试")
    logger.exception("unhandled MCP tool error", exc_info=True)
    return MCPToolError("internal_error", "服务内部错误", retryable=True, action="稍后重试；持续失败时查看服务日志")


def structured_errors(function: _Callable) -> _Callable:
    """把工具异常转换为带 code/retryable/action 的 MCP structuredContent。"""
    success_type = get_type_hints(function).get("return", Dict[str, Any])
    result_type = StructuredToolResult[Union[success_type, ToolFailure]]
    if inspect.iscoroutinefunction(function):
        @functools.wraps(function)
        async def async_wrapper(*args, **kwargs):
            try:
                return await function(*args, **kwargs)
            except Exception as exc:  # noqa: BLE001
                return _error_result(_classify_error(exc))

        async_wrapper.__annotations__ = dict(function.__annotations__)
        async_wrapper.__annotations__["return"] = result_type
        async_wrapper.__signature__ = inspect.signature(function).replace(  # type: ignore[attr-defined]
            return_annotation=result_type
        )
        return cast(_Callable, async_wrapper)

    @functools.wraps(function)
    def wrapper(*args, **kwargs):
        try:
            return function(*args, **kwargs)
        except Exception as exc:  # noqa: BLE001
            return _error_result(_classify_error(exc))

    wrapper.__annotations__ = dict(function.__annotations__)
    wrapper.__annotations__["return"] = result_type
    wrapper.__signature__ = inspect.signature(function).replace(  # type: ignore[attr-defined]
        return_annotation=result_type
    )
    return cast(_Callable, wrapper)

mcp_server = MCPServer(
    name="ppt-word-gen",
    title="ppt-word-gen",
    version="0.8.0",
    instructions=(
        "用于生成可编辑 PPTX 和 AI 仿真 Agent 交付类 DOCX。"
        "生成是异步任务：先创建任务，再查询状态，成功后获取 artifact 下载地址。"
        "Word 报告必须先调用 preview_word_report_format，把格式摘要展示给用户并取得明确确认，"
        "再把确认凭证传给 generate_word_report 或 revise_word_report。"
        "小型二进制资料可用 upload_file 直接通过 MCP 上传，大文件先调用 create_upload_ticket。"
        "优先调用 wait_generation_task 等待任务，成功后调用 get_artifact 获取 24 小时签名资源链接。"
    ),
)


def _upload(upload_id: str, purpose: str):
    if not upload_id.strip():
        return None
    data, filename, _ = upload_store.get(upload_id, expected_purpose=purpose)
    return data, filename


def _task_info(task_type: TaskType, task_id: str) -> GenerationTask:
    info = manager.get(task_id) if task_type == "presentation" else report_manager.get(task_id)
    if info is None:
        raise MCPToolError("task_not_found", f"{task_type} 任务不存在: {task_id}")
    payload = info.model_dump(mode="json")
    artifact_url = payload.pop("pptx_url", None) or payload.pop("document_url", None)
    return GenerationTask(task_type=task_type, artifact_url=artifact_url, **payload)


def _artifact_url(path: str) -> str:
    relative = f"{PUBLIC_URL_PREFIX}{path}"
    if PUBLIC_BASE_URL:
        return f"{PUBLIC_BASE_URL}{relative}"
    try:
        from mcp.server.dependencies import get_http_request

        request_base = str(get_http_request().base_url).rstrip("/")
    except (ImportError, RuntimeError):
        request_base = ""
    return f"{request_base}{relative}" if request_base else relative


def _template(template_id: str, document_type: Literal["presentation", "word_report"]):
    if not template_id.strip():
        return None
    item = business_template_store.get(template_id)
    if item.document_type != document_type:
        raise ValueError(f"业务模板 {template_id} 不适用于 {document_type}")
    return item


def _merge_instructions(original: str, business_instructions: str) -> str:
    parts = [item.strip() for item in (business_instructions, original) if item.strip()]
    return "\n\n".join(parts)


def _validate_uploaded_docx(metadata: Dict[str, Any]) -> None:
    suffix = Path(str(metadata["filename"])).suffix.lower()
    if suffix in {".docx", ".dotx"}:
        try:
            validate_docx_package(upload_store.path(str(metadata["upload_id"])))
        except Exception:
            upload_store.delete(str(metadata["upload_id"]))
            raise


def _upload_receipt(metadata: Dict[str, Any]) -> UploadReceipt:
    purpose = str(metadata["purpose"])
    return UploadReceipt(
        **metadata,
        use_in_mcp_as="template_upload_id" if purpose == "reference_template" else "source_upload_id",
    )


@mcp_server.tool()
@structured_errors
def list_generation_profiles() -> GenerationProfiles:
    """列出可生成的文档类型、格式以及二进制上传约定。"""
    return GenerationProfiles(**{
        "presentations": {
            "formats": ["ppt169", "ppt43"],
            "default_page_count": 8,
            "page_count_range": [1, 60],
            "output": "pptx",
        },
        "word_reports": {
            "report_types": REPORT_TYPES,
            "default_format_profile_id": WORD_FORMAT_PROFILE_ID,
            "format_fixed": False,
            "confirmation_required": True,
            "confirmation_tool": "preview_word_report_format",
            "custom_template_supported": True,
            "editable_format": [
                "body_font", "body_size_pt", "line_spacing", "first_line_indent_chars",
                "heading1_font", "heading1_size_pt", "heading2_font", "heading2_size_pt",
                "heading3_font", "heading3_size_pt", "numbering_style",
            ],
            "output": "docx",
        },
        "model_overrides": {
            "tools": ["generate_presentation", "generate_word_report"],
            "parameters": ["model", "base_url", "api_key", "temperature"],
            "optional": True,
            "fallback": "未传或传空值时使用服务端 LLM_* 默认配置",
        },
        "uploads": {
            "endpoint": f"{PUBLIC_URL_PREFIX}/api/v1/uploads",
            "method": "POST multipart/form-data",
            "fields": {"file": "binary", "purpose": "source | reference_template"},
            "mcp_inline_tool": "upload_file",
            "inline_max_mb": MCP_INLINE_UPLOAD_MB,
            "large_file_tool": "create_upload_ticket",
        },
        "workflow": [
            "preview_word_report_format (Word only)",
            "user confirms format (Word only)",
            "create",
            "wait_generation_task",
            "get_artifact",
        ],
    })


@mcp_server.tool()
@structured_errors
def preview_word_report_format(
    template_upload_id: str = "",
    body_font: str = "",
    body_size_pt: Optional[float] = None,
    line_spacing: Optional[float] = None,
    first_line_indent_chars: Optional[float] = None,
    heading1_font: str = "",
    heading1_size_pt: Optional[float] = None,
    heading2_font: str = "",
    heading2_size_pt: Optional[float] = None,
    heading3_font: str = "",
    heading3_size_pt: Optional[float] = None,
    numbering_style: Optional[Literal["decimal", "chinese"]] = None,
) -> WordFormatConfirmation:
    """提取模板默认格式并应用可选修改；把结果展示给用户，确认后再生成。"""
    template_path = None
    if template_upload_id.strip():
        _upload(template_upload_id, "reference_template")
        template_path = upload_store.path(template_upload_id)
    overrides = WordFormatOverrides(
        body_font=body_font,
        body_size_pt=body_size_pt,
        line_spacing=line_spacing,
        first_line_indent_chars=first_line_indent_chars,
        heading1_font=heading1_font,
        heading1_size_pt=heading1_size_pt,
        heading2_font=heading2_font,
        heading2_size_pt=heading2_size_pt,
        heading3_font=heading3_font,
        heading3_size_pt=heading3_size_pt,
        numbering_style=numbering_style,
    )
    return issue_word_format_confirmation(
        template_path=template_path,
        template_upload_id=template_upload_id,
        overrides=overrides,
    )


@mcp_server.tool()
@structured_errors
def generate_presentation(
    topic: str = "",
    page_count: int = 8,
    style: str = "",
    canvas_format: Literal["ppt169", "ppt43"] = "ppt169",
    model: str = "",
    base_url: str = "",
    api_key: str = "",
    temperature: Optional[float] = None,
    source_upload_id: str = "",
    business_template_id: str = "",
    idempotency_key: str = "",
) -> CreatedTask:
    """创建可编辑 PPTX 任务；模型参数可选，留空时使用服务端 LLM_* 默认配置。"""
    upload = _upload(source_upload_id, "source")
    business_template = _template(business_template_id, "presentation")
    if not topic.strip() and upload is None:
        raise ValueError("topic 与 source_upload_id 至少提供一个")
    effective_style = _merge_instructions(
        style,
        business_template.instructions if business_template else "",
    )
    request = TaskCreate(
        topic=topic,
        page_count=page_count,
        style=effective_style,
        format=canvas_format,
        language_model=model,
        language_base_url=base_url,
        language_api_key=api_key,
        language_temperature=temperature,
    )
    task_id, reused = manager.submit(
        request,
        upload,
        idempotency_key=idempotency_key.strip() or None,
    )
    return CreatedTask(
        task_id=task_id,
        task_type="presentation",
        status="pending",
        idempotency_reused=reused,
    )


@mcp_server.tool()
@structured_errors
def generate_word_report(
    format_confirmation_token: str,
    instructions: str = "",
    report_type: Literal["delivery", "validation", "manual", "technical"] = "validation",
    title: str = "",
    project_name: str = "",
    document_version: str = "v1.0",
    author: str = "",
    model: str = "",
    base_url: str = "",
    api_key: str = "",
    temperature: Optional[float] = None,
    source_upload_id: str = "",
    business_template_id: str = "",
    idempotency_key: str = "",
) -> CreatedTask:
    """用户确认模板及可编辑格式后创建 DOCX；确认凭证来自 preview_word_report_format。"""
    verified_format = verify_word_format_confirmation(
        format_confirmation_token,
        template_resolver=upload_store.path,
    )
    source_upload = _upload(source_upload_id, "source")
    template_upload = _upload(verified_format.template_upload_id, "reference_template")
    business_template = _template(business_template_id, "word_report")
    effective_instructions = _merge_instructions(
        instructions,
        business_template.instructions if business_template else "",
    )
    if not effective_instructions and source_upload is None:
        raise ValueError("instructions 与 source_upload_id 至少提供一个")
    request = ReportCreate(
        title=title,
        report_type=(business_template.report_type if business_template and business_template.report_type else report_type),
        instructions=effective_instructions,
        project_name=project_name,
        document_version=document_version,
        author=author,
        language_model=model,
        language_base_url=base_url,
        language_api_key=api_key,
        language_temperature=temperature,
        word_format_profile_id=verified_format.profile_id,
        word_format=verified_format.format,
    )
    task_id, reused = report_manager.submit(
        request,
        source_upload,
        template_upload,
        idempotency_key=idempotency_key.strip() or None,
    )
    return CreatedTask(
        task_id=task_id,
        task_type="word_report",
        status="pending",
        idempotency_reused=reused,
    )


@mcp_server.tool()
@structured_errors
def get_generation_task(task_type: TaskType, task_id: str) -> GenerationTask:
    """查询 PPT 或 Word 异步任务的阶段、进度、错误与结果状态。"""
    return _task_info(task_type, task_id)


@mcp_server.tool()
@structured_errors
def cancel_generation_task(task_type: TaskType, task_id: str) -> GenerationTask:
    """请求取消尚未结束的生成任务。"""
    info = manager.cancel(task_id) if task_type == "presentation" else report_manager.cancel(task_id)
    if info is None:
        raise MCPToolError("task_not_found", f"{task_type} 任务不存在: {task_id}")
    payload = info.model_dump(mode="json")
    artifact_url = payload.pop("pptx_url", None) or payload.pop("document_url", None)
    return GenerationTask(task_type=task_type, artifact_url=artifact_url, **payload)


@mcp_server.tool()
@structured_errors
def get_artifact(task_type: TaskType, task_id: str) -> Artifact:
    """任务成功后返回带标准 ResourceLink 的 24 小时签名下载地址。"""
    info = _task_info(task_type, task_id)
    if info.status != "success":
        raise MCPToolError(
            "task_not_ready",
            f"任务尚未成功，当前状态: {info.status}",
            retryable=info.status not in FINAL_STATUSES,
            action="调用 wait_generation_task 后重试",
        )
    if task_type == "presentation":
        path = manager.get_result_path(task_id)
        media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    else:
        path = report_manager.get_result_path(task_id)
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if not path or not Path(path).is_file():
        raise MCPToolError("artifact_not_found", "产物文件不存在或已被清理")
    result = Path(path)
    token = sign_claims("artifact", {"task_type": task_type, "task_id": task_id})
    url = _artifact_url(f"/api/v1/artifacts/{token}")
    artifact = Artifact(
        task_id=task_id,
        task_type=task_type,
        filename=result.name,
        bytes=result.stat().st_size,
        media_type=media_type,
        download_url=url,
        expires_at=int(time.time()) + SIGNED_URL_EXPIRE_SECONDS,
    )
    return CallToolResult(
        content=[
            ResourceLink(
                name=artifact.filename,
                title=artifact.filename,
                uri=url,
                description="24 小时内有效的生成产物下载链接",
                mimeType=artifact.media_type,
                size=artifact.bytes,
            )
        ],
        structuredContent=artifact.model_dump(mode="json"),
    )


@mcp_server.tool()
@structured_errors
def upload_file(
    filename: str,
    content_base64: str,
    purpose: Literal["source", "reference_template"] = "source",
) -> UploadReceipt:
    """通过 MCP 直接上传小文件；Base64 解码后上限由 MCP_INLINE_UPLOAD_MB 控制。"""
    encoded = content_base64.strip()
    if encoded.startswith("data:"):
        if ";base64," not in encoded:
            raise ValueError("仅支持 Base64 data URL")
        encoded = encoded.split(";base64,", 1)[1]
    try:
        data = base64.b64decode(encoded, validate=True)
    except (ValueError, binascii.Error) as exc:
        raise ValueError("content_base64 不是有效的 Base64") from exc
    limit = MCP_INLINE_UPLOAD_MB * 1024 * 1024
    if len(data) > limit:
        raise MCPToolError(
            "inline_upload_too_large",
            f"MCP 内联上传超过 {MCP_INLINE_UPLOAD_MB}MB 限制",
            action="调用 create_upload_ticket 后使用返回的 PUT 地址上传",
        )
    metadata = upload_store.put(data, filename, purpose)
    _validate_uploaded_docx(metadata)
    return _upload_receipt(metadata)


@mcp_server.tool()
@structured_errors
def create_upload_ticket(
    filename: str,
    purpose: Literal["source", "reference_template"] = "source",
) -> UploadTicket:
    """为大文件创建一次性 PUT 上传地址；无需了解服务内部上传 REST 契约。"""
    clean_name = safe_filename(filename, "upload")
    if purpose == "reference_template" and Path(clean_name).suffix.lower() not in {".docx", ".dotx"}:
        raise ValueError("参考模板必须是 .docx 或 .dotx 文件")
    ticket_id = uuid.uuid4().hex[:16]
    token = sign_claims(
        "upload",
        {
            "ticket_id": ticket_id,
            "filename": clean_name,
            "purpose": purpose,
            "max_bytes": MAX_UPLOAD_MB * 1024 * 1024,
        },
    )
    return UploadTicket(
        upload_url=_artifact_url(f"/api/v1/upload-tickets/{token}"),
        filename=clean_name,
        purpose=purpose,
        max_bytes=MAX_UPLOAD_MB * 1024 * 1024,
        expires_at=int(time.time()) + SIGNED_URL_EXPIRE_SECONDS,
    )


@mcp_server.tool()
@structured_errors
async def wait_generation_task(
    task_type: TaskType,
    task_id: str,
    timeout_seconds: int = 30,
) -> GenerationTask:
    """最长等待 55 秒；任务完成立即返回，否则在超时后返回最新进度。"""
    if not 1 <= timeout_seconds <= 55:
        raise ValueError("timeout_seconds 必须在 1 到 55 之间")
    deadline = asyncio.get_running_loop().time() + timeout_seconds
    current = _task_info(task_type, task_id)
    while current.status not in FINAL_STATUSES and asyncio.get_running_loop().time() < deadline:
        await asyncio.sleep(1)
        current = _task_info(task_type, task_id)
    return current


@mcp_server.tool()
@structured_errors
def list_business_templates(
    document_type: DocumentTypeFilter = "all",
) -> BusinessTemplateList:
    """列出内置和公司自定义的 PPT/Word 业务模板。"""
    return BusinessTemplateList(templates=business_template_store.list(document_type))


@mcp_server.tool()
@structured_errors
def register_business_template(
    name: str,
    document_type: Literal["presentation", "word_report"],
    instructions: str = "",
    description: str = "",
    report_type: Literal["", "delivery", "validation", "manual", "technical"] = "",
) -> BusinessTemplate:
    """注册业务内容与风格说明；Word 版式模板在预览确认工具中选择。"""
    return business_template_store.register(
        name=name,
        document_type=document_type,
        description=description,
        instructions=instructions,
        report_type=report_type,
        file_data=None,
        filename="",
    )


@mcp_server.tool()
@structured_errors
def delete_business_template(template_id: str) -> DeleteBusinessTemplateResult:
    """删除自定义业务模板；内置模板不可删除。"""
    return DeleteBusinessTemplateResult(
        template_id=template_id,
        deleted=business_template_store.delete(template_id),
    )


@mcp_server.tool()
@structured_errors
def revise_presentation(
    source_task_id: str,
    instructions: str,
    page_count: int = 0,
    style: str = "",
    canvas_format: Literal["ppt169", "ppt43"] = "ppt169",
    business_template_id: str = "",
    idempotency_key: str = "",
) -> CreatedTask:
    """把已成功生成的 PPTX 作为上下文，按自然语言要求生成一个新版本。"""
    source = _task_info("presentation", source_task_id)
    if source.status != "success":
        raise MCPToolError("task_not_ready", f"源任务状态为 {source.status}", retryable=True)
    path_value = manager.get_result_path(source_task_id)
    path = Path(path_value) if path_value else None
    if path is None or not path.is_file():
        raise MCPToolError("artifact_not_found", "源 PPTX 不存在或已被清理")
    if not instructions.strip():
        raise ValueError("instructions 不能为空")
    if page_count == 0:
        from pptx import Presentation

        page_count = len(Presentation(str(path)).slides)
    business_template = _template(business_template_id, "presentation")
    request = TaskCreate(
        topic=(
            "请基于来源中的现有演示文稿进行语义化修订。保留未要求修改的事实和结构，"
            "不要虚构数据。\n\n修订要求：" + instructions.strip()
        ),
        page_count=page_count,
        style=_merge_instructions(style, business_template.instructions if business_template else ""),
        format=canvas_format,
    )
    task_id, reused = manager.submit(
        request,
        (path.read_bytes(), path.name),
        idempotency_key=idempotency_key.strip() or None,
    )
    return CreatedTask(
        task_id=task_id,
        task_type="presentation",
        status="pending",
        idempotency_reused=reused,
    )


@mcp_server.tool()
@structured_errors
def revise_word_report(
    format_confirmation_token: str,
    source_task_id: str,
    instructions: str,
    report_type: Literal["delivery", "validation", "manual", "technical"] = "validation",
    title: str = "",
    project_name: str = "",
    document_version: str = "v1.0",
    author: str = "",
    business_template_id: str = "",
    idempotency_key: str = "",
) -> CreatedTask:
    """用户确认模板及格式后，把已有 DOCX 作为来源生成一个新版本。"""
    verified_format = verify_word_format_confirmation(
        format_confirmation_token,
        template_resolver=upload_store.path,
    )
    source = _task_info("word_report", source_task_id)
    if source.status != "success":
        raise MCPToolError("task_not_ready", f"源任务状态为 {source.status}", retryable=True)
    path_value = report_manager.get_result_path(source_task_id)
    path = Path(path_value) if path_value else None
    if path is None or not path.is_file():
        raise MCPToolError("artifact_not_found", "源 DOCX 不存在或已被清理")
    if not instructions.strip():
        raise ValueError("instructions 不能为空")
    business_template = _template(business_template_id, "word_report")
    effective_instructions = _merge_instructions(
        (
            "请把来源文档视为当前版本，只修改明确要求的部分，保留其他事实、证据与章节。"
            "不要虚构数据。\n\n修订要求：" + instructions.strip()
        ),
        business_template.instructions if business_template else "",
    )
    request = ReportCreate(
        title=title,
        report_type=(business_template.report_type if business_template and business_template.report_type else report_type),
        instructions=effective_instructions,
        project_name=project_name,
        document_version=document_version,
        author=author,
        word_format_profile_id=verified_format.profile_id,
        word_format=verified_format.format,
    )
    task_id, reused = report_manager.submit(
        request,
        (path.read_bytes(), path.name),
        _upload(verified_format.template_upload_id, "reference_template"),
        idempotency_key=idempotency_key.strip() or None,
    )
    return CreatedTask(
        task_id=task_id,
        task_type="word_report",
        status="pending",
        idempotency_reused=reused,
    )


@mcp_server.resource(
    "ppt-word://rules/workflow",
    name="generation-workflow",
    title="PPT/Word 生成工作流",
    mime_type="text/markdown",
)
def workflow_resource() -> str:
    return """# 生成工作流

1. 小文件调用 `upload_file`，大文件调用 `create_upload_ticket` 后 PUT 上传。
2. Word 模板先以 `purpose=reference_template` 上传；不传模板时使用内置 CID629 模板。
3. 调用 `preview_word_report_format` 提取模板默认格式，可按用户要求修改字体、字号、行距、缩进和编号。
4. 向用户完整展示返回的格式摘要并取得明确确认，再把 `confirmation_token` 传给生成或修订工具。
5. 调用 `wait_generation_task`，必要时再次等待。
6. 成功后调用 `get_artifact`，使用返回的 ResourceLink；链接默认 24 小时有效。
7. 失败时读取 `structuredContent.error` 中的 code、retryable 和 suggested_action。
"""


@mcp_server.resource(
    "ppt-word://rules/presentation",
    name="presentation-rules",
    title="公司技术汇报规范",
    mime_type="text/markdown",
)
def presentation_rules_resource() -> str:
    return """# PPT 规范

- 一页一个核心结论，优先使用架构图、流程图、对比表和证据图。
- AI 仿真主题应说明业务场景、Agent、模型/工具、仿真链路、接口、验证与风险。
- 不虚构测试结果、性能数字、客户结论或系统能力。
- 默认使用可编辑元素，避免把正文整体渲染成图片。
"""


@mcp_server.resource(
    "ppt-word://rules/word-report",
    name="word-report-rules",
    title="AI 仿真 Agent 文档规范",
    mime_type="text/markdown",
)
def word_report_rules_resource() -> str:
    return """# Word 报告规范

- 交付报告应覆盖范围、架构、部署、接口、测试证据、限制、风险和验收结论。
- 测试报告必须区分来源事实、测试结果和推断，不得补造缺失数据。
- 未上传模板时使用 CID629 v1.5；上传 DOCX/DOTX 时，以该模板样式作为确认默认值。
- 必须同时包含一级和二级标题，采用多级编号，并生成 1–3 级自动目录。
- 确认时可修改正文和 1–3 级标题的字体、字号，以及行距、缩进、编号形式。
- 最终确认凭证同时绑定模板哈希和修改后的格式，模板或格式变化后必须重新确认。
- 语义修订仅改变用户明确要求的部分，并生成新任务，不覆盖旧产物。
"""


@mcp_server.resource(
    "ppt-word://templates/catalog",
    name="business-template-catalog",
    title="业务模板目录",
    mime_type="application/json",
)
def business_template_catalog_resource() -> str:
    return json.dumps(
        {
            "templates": [
                item.model_dump(mode="json") for item in business_template_store.list("all")
            ]
        },
        ensure_ascii=False,
        indent=2,
    )
