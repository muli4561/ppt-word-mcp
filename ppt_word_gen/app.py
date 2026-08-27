# coding=utf-8
"""ppt-word-gen —— PPT / Word 生成服务入口。

单一生成服务：POST 任务 → GET 进度 → GET 结果。
对外入口：8000 提供 Demo、API、健康检查与文件下载。
"""
import logging
import hmac
import tempfile
import time
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Literal, Optional

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field, SecretStr

from .agent import LLMClient
from .config import (
    BASE_DIR,
    MCP_ALLOWED_HOSTS,
    MCP_ALLOWED_ORIGINS,
    MCP_DNS_REBINDING_PROTECTION,
    PUBLIC_URL_PREFIX,
    LLM_API_KEY,
    LLM_BASE_URL,
    LLM_MODEL,
    LLM_TEMPERATURE,
    MAX_GENERATED_IMAGES,
    MAX_UPLOAD_MB,
    MOCK_LLM,
    PPT_WORD_GEN_TOKEN,
    VISION_API_KEY,
    VISION_BACKEND,
    VISION_BASE_URL,
    VISION_ENABLED,
    VISION_IMAGE_SIZE,
    VISION_MODEL,
)
from mcp.server.transport_security import TransportSecuritySettings
from .mcp_server import mcp_server
from .pptmaster import generate_image
from .report_documents import validate_docx_package
from .report_models import ReportCreate, ReportTaskInfo, WordFormatOverrides
from .report_tasks import (
    ReportIdempotencyConflict,
    ReportQueueFull,
    report_manager,
)
from .signed_tokens import InvalidSignedToken, verify_claims
from .tasks import IdempotencyConflict, TaskCreate, TaskInfo, TaskQueueFull, manager
from .upload_store import UploadNotFound, upload_store
from .word_format import (
    WordFormatConfirmation,
    issue_word_format_confirmation,
    verify_word_format_confirmation,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(levelname)s %(message)s")

_mcp_http_app = mcp_server.streamable_http_app(
    json_response=True,
    stateless_http=True,
    transport_security=TransportSecuritySettings(
        enable_dns_rebinding_protection=MCP_DNS_REBINDING_PROTECTION,
        allowed_hosts=MCP_ALLOWED_HOSTS,
        allowed_origins=MCP_ALLOWED_ORIGINS,
    ),
)


@asynccontextmanager
async def lifespan(_app: FastAPI):
    async with mcp_server.session_manager.run():
        try:
            yield
        finally:
            manager.shutdown()
            report_manager.shutdown()


app = FastAPI(
    title="ppt-word-gen",
    version="0.8.0",
    description="PPT / Word 文档生成服务：Demo、REST API 与 MCP",
    lifespan=lifespan,
)

_bearer = HTTPBearer(auto_error=False)

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_STATIC_DIR = BASE_DIR / "static"


def _check_auth(credentials: Optional[HTTPAuthorizationCredentials]) -> None:
    if not PPT_WORD_GEN_TOKEN:
        return
    if credentials is None or credentials.credentials != PPT_WORD_GEN_TOKEN:
        raise HTTPException(status_code=401, detail="无效的 Bearer Token")


class MCPBearerAuthMiddleware:
    """让 MCP 端点复用 PPT_WORD_GEN_TOKEN，不引入额外认证系统。"""

    def __init__(self, asgi_app, expected_token: str):
        self.app = asgi_app
        self.expected_token = expected_token

    async def __call__(self, scope, receive, send):
        if scope.get("type") == "http" and self.expected_token:
            headers = dict(scope.get("headers", []))
            authorization = headers.get(b"authorization", b"").decode("latin1")
            supplied = authorization[7:] if authorization.lower().startswith("bearer ") else ""
            if not hmac.compare_digest(supplied, self.expected_token):
                body = b'{"detail":"invalid bearer token"}'
                await send({
                    "type": "http.response.start",
                    "status": 401,
                    "headers": [(b"content-type", b"application/json"), (b"content-length", str(len(body)).encode())],
                })
                await send({"type": "http.response.body", "body": body})
                return
        await self.app(scope, receive, send)


@app.get("/health")
def health():
    task_stats = manager.stats()
    report_stats = report_manager.stats()
    return {
        "status": "ok",
        "service": "ppt-word-gen",
        "version": app.version,
        "mock_llm": MOCK_LLM,
        "custom_models": True,
        "task_store": "ok" if task_stats["database_ok"] else "error",
        "active_tasks": task_stats["active"],
        "queued_tasks": task_stats["queued"],
        "queue_capacity": task_stats["queue_capacity"],
        "report_task_store": "ok" if report_stats["database_ok"] else "error",
        "active_report_tasks": report_stats["active"],
        "queued_report_tasks": report_stats["queued"],
        "report_queue_capacity": report_stats["queue_capacity"],
        "mcp": {
            "status": "ok",
            "endpoint": f"{PUBLIC_URL_PREFIX}/mcp",
            "signed_artifacts": True,
            "resources": True,
        },
    }


@app.get("/api/v1/models")
def model_catalog():
    """供 Demo 初始化的非敏感模型默认值与能力清单。永不返回 API Key。"""
    return {
        "language": {
            "base_url": LLM_BASE_URL,
            "model": LLM_MODEL,
            "temperature": LLM_TEMPERATURE,
            "server_key_configured": bool(LLM_API_KEY),
            "protocol": "OpenAI-compatible chat completions",
        },
        "vision": {
            "enabled": VISION_ENABLED,
            "backend": VISION_BACKEND,
            "base_url": VISION_BASE_URL,
            "model": VISION_MODEL,
            "image_size": VISION_IMAGE_SIZE,
            "server_key_configured": bool(VISION_API_KEY),
            "backends": ["openai", "gemini", "qwen", "zhipu", "volcengine"],
            "max_images_per_task": MAX_GENERATED_IMAGES,
            "adapter": "PPT Master image_gen.py",
        },
    }


class ModelTestRequest(BaseModel):
    kind: Literal["language", "vision"] = "language"
    base_url: str = Field("", max_length=500)
    api_key: SecretStr = Field(default_factory=lambda: SecretStr(""), repr=False)
    model: str = Field("", max_length=200)
    temperature: float = Field(0.7, ge=0, le=2)
    backend: str = Field("openai", max_length=40)
    image_size: str = Field("1K", max_length=10)


@app.post("/api/v1/models/test")
def test_model(
    request: ModelTestRequest,
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    """按用户显式请求测试一次模型调用；服务自测不会调用此接口。"""
    _check_auth(token)
    started = time.perf_counter()
    try:
        if request.kind == "language":
            validated = TaskCreate(
                language_base_url=request.base_url,
                language_api_key=request.api_key,
                language_model=request.model,
                language_temperature=request.temperature,
            )
            client = LLMClient(
                base_url=validated.language_base_url or LLM_BASE_URL,
                api_key=validated.language_api_key.get_secret_value() or LLM_API_KEY,
                model=validated.language_model or LLM_MODEL,
                temperature=validated.language_temperature,
            )
            message = client.chat([
                {"role": "system", "content": "你是连接测试助手。"},
                {"role": "user", "content": "仅回复 OK"},
            ])
            preview = (getattr(message, "content", "") or "")[:100]
            return {"ok": True, "kind": "language", "latency_ms": round((time.perf_counter() - started) * 1000), "preview": preview}

        validated = TaskCreate(
            vision_enabled=True,
            vision_backend=request.backend,
            vision_base_url=request.base_url,
            vision_api_key=request.api_key,
            vision_model=request.model,
            vision_image_size=request.image_size,
        )
        with tempfile.TemporaryDirectory(prefix="ppt-model-test-") as temp_dir:
            path = generate_image(
                output_dir=Path(temp_dir),
                prompt="Minimal abstract presentation background, navy and cyan geometric shapes, no text",
                backend=validated.vision_backend,
                api_key=validated.vision_api_key.get_secret_value() or VISION_API_KEY,
                base_url=validated.vision_base_url or VISION_BASE_URL,
                model=validated.vision_model or VISION_MODEL,
                image_size=validated.vision_image_size,
                aspect_ratio="16:9",
                filename="connection_test",
            )
            size = path.stat().st_size
        return {"ok": True, "kind": "vision", "latency_ms": round((time.perf_counter() - started) * 1000), "bytes": size}
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"模型连接失败：{str(exc)[:800]}") from None


@app.get("/demo", include_in_schema=False)
def demo_page():
    """演示页：浏览器直开，可生成 PPT 或 Word。"""
    html = _STATIC_DIR / "demo.html"
    if not html.is_file():
        raise HTTPException(status_code=404, detail="demo.html 缺失")
    content = html.read_text(encoding="utf-8").replace("__PUBLIC_URL_PREFIX__", PUBLIC_URL_PREFIX)
    from fastapi.responses import HTMLResponse
    return HTMLResponse(content)


app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")


@app.post("/api/v1/uploads", response_model=dict)
async def create_upload(
    file: UploadFile = File(...),
    purpose: Literal["source", "reference_template"] = Form("source"),
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    """为 MCP 任务暂存二进制资料，返回 24 小时有效的 upload_id。"""
    _check_auth(token)
    data = await file.read()
    if len(data) > MAX_UPLOAD_MB * 1024 * 1024:
        raise HTTPException(status_code=413, detail=f"上传文件超过 {MAX_UPLOAD_MB}MB 限制")
    try:
        metadata = upload_store.put(data, file.filename or "upload", purpose)
        if Path(metadata["filename"]).suffix.lower() in {".docx", ".dotx"}:
            validate_docx_package(upload_store.path(metadata["upload_id"]))
    except ValueError as exc:
        if "metadata" in locals():
            upload_store.delete(metadata["upload_id"])
        raise HTTPException(status_code=422, detail=str(exc)) from None
    metadata["use_in_mcp_as"] = (
        "template_upload_id" if purpose == "reference_template" else "source_upload_id"
    )
    return metadata


@app.get("/api/v1/uploads/{upload_id}", response_model=dict)
def get_upload(
    upload_id: str,
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    _check_auth(token)
    try:
        return upload_store.info(upload_id)
    except UploadNotFound as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from None


@app.delete("/api/v1/uploads/{upload_id}", response_model=dict)
def delete_upload(
    upload_id: str,
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    _check_auth(token)
    try:
        deleted = upload_store.delete(upload_id)
    except UploadNotFound as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from None
    return {"upload_id": upload_id, "deleted": deleted}


@app.put("/api/v1/upload-tickets/{signed_token}", response_model=dict)
async def upload_with_ticket(signed_token: str, request: Request):
    """使用 MCP 创建的一次性签名票据上传大文件。"""
    try:
        claims = verify_claims(signed_token, "upload")
        ticket_id = str(claims["ticket_id"])
        filename = str(claims["filename"])
        purpose = str(claims["purpose"])
        max_bytes = min(int(claims["max_bytes"]), MAX_UPLOAD_MB * 1024 * 1024)
    except (InvalidSignedToken, KeyError, TypeError, ValueError) as exc:
        raise HTTPException(status_code=403, detail=str(exc)) from None
    if purpose not in {"source", "reference_template"}:
        raise HTTPException(status_code=403, detail="上传票据用途无效")
    data = await request.body()
    if not data:
        raise HTTPException(status_code=400, detail="上传文件为空")
    if len(data) > max_bytes:
        raise HTTPException(status_code=413, detail=f"上传文件超过 {max_bytes} 字节限制")
    metadata = None
    try:
        metadata = upload_store.put(data, filename, purpose)
        if Path(str(metadata["filename"])).suffix.lower() in {".docx", ".dotx"}:
            validate_docx_package(upload_store.path(str(metadata["upload_id"])))
        upload_store.consume_ticket(ticket_id)
    except ValueError as exc:
        if metadata is not None:
            upload_store.delete(str(metadata["upload_id"]))
        raise HTTPException(status_code=400, detail=str(exc)) from None
    metadata["use_in_mcp_as"] = (
        "source_upload_id" if purpose == "source" else "template_upload_id"
    )
    return metadata


@app.get("/api/v1/artifacts/{signed_token}")
def download_signed_artifact(signed_token: str):
    """下载 MCP 返回的短效签名产物；URL 自身即授权，不要求额外请求头。"""
    try:
        claims = verify_claims(signed_token, "artifact")
        task_type = str(claims["task_type"])
        task_id = str(claims["task_id"])
    except (InvalidSignedToken, KeyError, TypeError, ValueError) as exc:
        raise HTTPException(status_code=403, detail=str(exc)) from None
    if task_type == "presentation":
        info = manager.get(task_id)
        path_value = manager.get_result_path(task_id)
        media_type = PPTX_MIME
    elif task_type == "word_report":
        info = report_manager.get(task_id)
        path_value = report_manager.get_result_path(task_id)
        media_type = DOCX_MIME
    else:
        raise HTTPException(status_code=403, detail="产物类型无效")
    if info is None or info.status != "success":
        raise HTTPException(status_code=404, detail="产物不存在")
    path = Path(path_value) if path_value else None
    if path is None or not path.is_file():
        raise HTTPException(status_code=404, detail="产物不存在或已清理")
    return FileResponse(path, media_type=media_type, filename=path.name)


@app.post("/api/v1/tasks", response_model=dict)
async def create_task(
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
    idempotency_key: Optional[str] = Header(None, alias="Idempotency-Key"),
    topic: str = Form(""),
    page_count: int = Form(8),
    style: str = Form(""),
    format: str = Form("ppt169"),
    language_base_url: str = Form(""),
    language_api_key: str = Form(""),
    language_model: str = Form(""),
    language_temperature: Optional[float] = Form(None),
    vision_enabled: bool = Form(False),
    vision_backend: str = Form("openai"),
    vision_base_url: str = Form(""),
    vision_api_key: str = Form(""),
    vision_model: str = Form(""),
    vision_image_size: str = Form("1K"),
    source_file: Optional[UploadFile] = File(None),
):
    _check_auth(token)
    req = TaskCreate(
        topic=topic,
        page_count=page_count,
        style=style,
        format=format,
        language_base_url=language_base_url,
        language_api_key=language_api_key,
        language_model=language_model,
        language_temperature=language_temperature,
        vision_enabled=vision_enabled,
        vision_backend=vision_backend,
        vision_base_url=vision_base_url,
        vision_api_key=vision_api_key,
        vision_model=vision_model,
        vision_image_size=vision_image_size,
    )
    upload = None
    if source_file is not None and source_file.filename:
        data = await source_file.read()
        if len(data) > MAX_UPLOAD_MB * 1024 * 1024:
            raise HTTPException(status_code=413, detail=f"上传文件超过 {MAX_UPLOAD_MB}MB 限制")
        upload = (data, source_file.filename)
    try:
        task_id, reused = manager.submit(req, upload, idempotency_key=idempotency_key)
    except TaskQueueFull as exc:
        raise HTTPException(
            status_code=429,
            detail=str(exc),
            headers={"Retry-After": "30"},
        ) from None
    except IdempotencyConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from None
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from None
    return {
        "task_id": task_id,
        "status": "pending",
        "poll": f"{PUBLIC_URL_PREFIX}/api/v1/tasks/{task_id}",
        "idempotency_reused": reused,
    }


@app.get("/api/v1/tasks/{task_id}", response_model=TaskInfo)
def get_task(task_id: str, token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer)):
    _check_auth(token)
    info = manager.get(task_id)
    if info is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    return info


@app.post("/api/v1/tasks/{task_id}/cancel", response_model=TaskInfo)
def cancel_task(task_id: str, token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer)):
    _check_auth(token)
    info = manager.cancel(task_id)
    if info is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    return info


@app.get("/api/v1/tasks/{task_id}/result")
def download_result(task_id: str, token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer)):
    _check_auth(token)
    info = manager.get(task_id)
    if info is None:
        raise HTTPException(status_code=404, detail="任务不存在")
    if info.status != "success":
        raise HTTPException(status_code=400, detail=f"任务尚未完成（status={info.status}）")
    pptx_abs = manager.get_result_path(task_id)
    if not pptx_abs or not Path(pptx_abs).is_file():
        raise HTTPException(status_code=404, detail="结果文件不存在或已清理")
    return FileResponse(pptx_abs, media_type=PPTX_MIME, filename=Path(pptx_abs).name)


# ---------------------------------------------------------------- Demo-only Word report route

@app.get("/api/v1/word-format", response_model=WordFormatConfirmation)
def get_word_format_confirmation(
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    """返回内置模板的默认格式和短效确认凭证；不会创建任务。"""
    _check_auth(token)
    return issue_word_format_confirmation()


class WordFormatPreviewRequest(BaseModel):
    template_upload_id: str = Field("", max_length=64)
    format: WordFormatOverrides = Field(default_factory=WordFormatOverrides)


@app.post("/api/v1/word-format", response_model=WordFormatConfirmation)
def preview_word_format(
    request: WordFormatPreviewRequest,
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    """从上传模板提取默认格式，并用用户修改后的值签发确认凭证。"""
    _check_auth(token)
    try:
        template_path = None
        if request.template_upload_id.strip():
            upload_store.get(request.template_upload_id, "reference_template")
            template_path = upload_store.path(request.template_upload_id)
        return issue_word_format_confirmation(
            template_path=template_path,
            template_upload_id=request.template_upload_id,
            overrides=request.format,
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from None


@app.post("/api/v1/report-tasks", response_model=dict)
async def create_report_task(
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
    idempotency_key: Optional[str] = Header(None, alias="Idempotency-Key"),
    title: str = Form(""),
    report_type: Literal["delivery", "validation", "manual", "technical"] = Form("validation"),
    instructions: str = Form(""),
    project_name: str = Form(""),
    document_version: str = Form("v1.0"),
    author: str = Form(""),
    language_base_url: str = Form(""),
    language_api_key: str = Form(""),
    language_model: str = Form(""),
    language_temperature: Optional[float] = Form(None),
    format_confirmation_token: str = Form(...),
    source_file: Optional[UploadFile] = File(None),
):
    """Demo 专用：用户确认模板及可编辑格式后生成 DOCX。"""
    _check_auth(token)
    try:
        verified_format = verify_word_format_confirmation(
            format_confirmation_token,
            template_resolver=upload_store.path,
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from None
    request = ReportCreate(
        title=title,
        report_type=report_type,
        instructions=instructions,
        project_name=project_name,
        document_version=document_version,
        author=author,
        language_base_url=language_base_url,
        language_api_key=language_api_key,
        language_model=language_model,
        language_temperature=language_temperature,
        word_format_profile_id=verified_format.profile_id,
        word_format=verified_format.format,
    )

    async def read_limited(upload: Optional[UploadFile], label: str):
        if upload is None or not upload.filename:
            return None
        data = await upload.read()
        if len(data) > MAX_UPLOAD_MB * 1024 * 1024:
            raise HTTPException(status_code=413, detail=f"{label}超过 {MAX_UPLOAD_MB}MB 限制")
        return data, upload.filename

    source_upload = await read_limited(source_file, "来源资料")
    template_upload = None
    if verified_format.template_upload_id:
        data, filename, _ = upload_store.get(
            verified_format.template_upload_id,
            "reference_template",
        )
        template_upload = (data, filename)
    try:
        task_id, reused = report_manager.submit(
            request,
            source_upload,
            template_upload,
            idempotency_key=idempotency_key,
        )
    except ReportQueueFull as exc:
        raise HTTPException(
            status_code=429,
            detail=str(exc),
            headers={"Retry-After": "30"},
        ) from None
    except ReportIdempotencyConflict as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from None
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from None
    return {
        "task_id": task_id,
        "status": "pending",
        "poll": f"{PUBLIC_URL_PREFIX}/api/v1/report-tasks/{task_id}",
        "idempotency_reused": reused,
    }


@app.get("/api/v1/report-tasks/{task_id}", response_model=ReportTaskInfo)
def get_report_task(
    task_id: str,
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    _check_auth(token)
    info = report_manager.get(task_id)
    if info is None:
        raise HTTPException(status_code=404, detail="报告任务不存在")
    return info


@app.post("/api/v1/report-tasks/{task_id}/cancel", response_model=ReportTaskInfo)
def cancel_report_task(
    task_id: str,
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    _check_auth(token)
    info = report_manager.cancel(task_id)
    if info is None:
        raise HTTPException(status_code=404, detail="报告任务不存在")
    return info


@app.get("/api/v1/report-tasks/{task_id}/result")
def download_report_result(
    task_id: str,
    token: Optional[HTTPAuthorizationCredentials] = Depends(_bearer),
):
    _check_auth(token)
    info = report_manager.get(task_id)
    if info is None:
        raise HTTPException(status_code=404, detail="报告任务不存在")
    if info.status != "success":
        raise HTTPException(status_code=400, detail=f"报告尚未完成（status={info.status}）")
    path = report_manager.get_result_path(task_id)
    if not path or not Path(path).is_file():
        raise HTTPException(status_code=404, detail="结果文件不存在或已清理")
    return FileResponse(path, media_type=DOCX_MIME, filename=Path(path).name)


# 必须最后挂载：Mount("/") 保持 MCP 公共地址为 /mcp，且不遮挡上面的 REST/Demo 路由。
app.mount("/", MCPBearerAuthMiddleware(_mcp_http_app, PPT_WORD_GEN_TOKEN), name="mcp")
