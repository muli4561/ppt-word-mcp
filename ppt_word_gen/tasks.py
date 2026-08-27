# coding=utf-8
"""任务队列：SQLite 持久化状态 + 有界内存工作队列。"""
import hashlib
import json
import logging
import queue
import threading
import time
import uuid
from typing import Dict, Optional, Tuple

from pydantic import BaseModel, Field, SecretStr, field_validator
from urllib.parse import urlsplit

from .agent import AgentLoop, MockAgent, ToolContext
from .config import (
    MAX_CONCURRENT_TASKS,
    MAX_QUEUED_TASKS,
    MOCK_LLM,
    TASK_CLEANUP_INTERVAL_SECONDS,
    TASK_DB_PATH,
    TASK_EXPIRE_HOURS,
    PUBLIC_URL_PREFIX,
)
from .pptmaster import TaskWorkspace, convert_source, import_sources, init_project
from .task_store import FINAL_STATUSES, DuplicateIdempotencyKey, TaskStore

logger = logging.getLogger("ppt_word_gen.tasks")

STATUS_PENDING = "pending"
STATUS_RUNNING = "running"
STATUS_SUCCESS = "success"
STATUS_FAILED = "failed"
STATUS_CANCELLING = "cancelling"
STATUS_CANCELLED = "cancelled"
STATUS_INTERRUPTED = "interrupted"


class TaskQueueFull(RuntimeError):
    pass


class IdempotencyConflict(RuntimeError):
    pass


class TaskCancelled(RuntimeError):
    pass


class TaskCreate(BaseModel):
    """提交生成任务的参数（首期：文本/文档输入）。"""

    topic: str = Field("", description="PPT 主题或直接粘贴的正文内容")
    page_count: int = Field(8, ge=1, le=60, description="目标页数")
    style: str = Field("", description="风格偏好（商务/科技/杂志/新闻/极简…），留空由模型决定")
    format: str = Field("ppt169", description="画布格式：ppt169 / ppt43 / 竖屏等")

    # 语言模型：为空时使用服务端 LLM_* 默认值。SecretStr 避免日志/异常回显密钥。
    language_base_url: str = Field("", max_length=500)
    language_api_key: SecretStr = Field(default_factory=lambda: SecretStr(""), repr=False)
    language_model: str = Field("", max_length=200)
    language_temperature: Optional[float] = Field(None, ge=0, le=2)

    # 视觉模型：复用 PPT Master image_gen.py 的 provider 后端。
    vision_enabled: bool = False
    vision_backend: str = Field("openai", max_length=40)
    vision_base_url: str = Field("", max_length=500)
    vision_api_key: SecretStr = Field(default_factory=lambda: SecretStr(""), repr=False)
    vision_model: str = Field("", max_length=200)
    vision_image_size: str = Field("1K", max_length=10)

    @field_validator("language_base_url", "vision_base_url")
    @classmethod
    def validate_model_url(cls, value: str) -> str:
        value = value.strip().rstrip("/")
        if not value:
            return value
        parsed = urlsplit(value)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise ValueError("模型地址必须是有效的 http/https URL")
        if parsed.username or parsed.password:
            raise ValueError("模型地址中不能包含用户名或密码")
        return value

    @field_validator("vision_backend")
    @classmethod
    def normalize_vision_backend(cls, value: str) -> str:
        value = value.strip().lower()
        supported = {"openai", "gemini", "qwen", "zhipu", "volcengine"}
        if value not in supported:
            raise ValueError(f"视觉模型供应商仅支持: {', '.join(sorted(supported))}")
        return value

    @field_validator("vision_image_size")
    @classmethod
    def validate_image_size(cls, value: str) -> str:
        normalized = value.strip().upper()
        aliases = {"512PX": "512px", "1K": "1K", "2K": "2K", "4K": "4K"}
        if normalized not in aliases:
            raise ValueError("视觉模型尺寸仅支持 512px / 1K / 2K / 4K")
        return aliases[normalized]


class TaskInfo(BaseModel):
    task_id: str
    status: str  # pending | running | cancelling | success | failed | cancelled | interrupted
    stage: str = ""
    progress: int = Field(0, ge=0, le=100)
    message: str = ""
    error: Optional[str] = None
    pptx_url: Optional[str] = None
    cancel_requested: bool = False
    created_at: float = 0.0
    updated_at: float = 0.0


class TaskRecord:
    def __init__(self, task_id: str, store: TaskStore):
        self.task_id = task_id
        self._store = store
        self._lock = threading.Lock()
        self.status = STATUS_PENDING
        self.stage = "排队中"
        self.progress = 0
        self.message = ""
        self.error: Optional[str] = None
        self.pptx_url: Optional[str] = None
        self.pptx_abs: Optional[str] = None
        self.project_dir: Optional[str] = None
        self.cancel_requested = False
        self.created_at = time.time()
        self.updated_at = self.created_at

    def set(self, status: str, stage: str = "", progress: int = None, message: str = "",
            error: Optional[str] = None, pptx_url: Optional[str] = None,
            pptx_abs: Optional[str] = None, project_dir: Optional[str] = None) -> None:
        with self._lock:
            if self.cancel_requested and status in {STATUS_PENDING, STATUS_RUNNING}:
                raise TaskCancelled("task cancellation requested")
            self.status = status
            if stage:
                self.stage = stage
            if progress is not None:
                self.progress = progress
            if message:
                self.message = message
            if error is not None:
                self.error = error
            if pptx_url is not None:
                self.pptx_url = pptx_url
            if pptx_abs is not None:
                self.pptx_abs = pptx_abs
            if project_dir is not None:
                self.project_dir = project_dir
            self.updated_at = time.time()
            fields = {
                "status": self.status,
                "stage": self.stage,
                "progress": self.progress,
                "message": self.message,
                "error": self.error,
                "pptx_url": self.pptx_url,
                "pptx_abs": self.pptx_abs,
                "project_dir": self.project_dir,
                "cancel_requested": self.cancel_requested,
                "updated_at": self.updated_at,
            }
        self._store.update(self.task_id, **fields)

    def request_cancel(self) -> None:
        with self._lock:
            if self.status in FINAL_STATUSES:
                return
            self.cancel_requested = True
            if self.status == STATUS_PENDING:
                self.status = STATUS_CANCELLED
                self.stage = "已取消"
                self.message = "任务已在排队阶段取消"
            else:
                self.status = STATUS_CANCELLING
                self.stage = "正在取消"
                self.message = "当前步骤结束后停止"
            self.updated_at = time.time()
            fields = {
                "status": self.status,
                "stage": self.stage,
                "message": self.message,
                "cancel_requested": True,
                "updated_at": self.updated_at,
            }
        self._store.update(self.task_id, **fields)

    def raise_if_cancelled(self) -> None:
        with self._lock:
            if self.cancel_requested:
                raise TaskCancelled("task cancellation requested")

    def to_info(self) -> TaskInfo:
        with self._lock:
            return TaskInfo(
                task_id=self.task_id,
                status=self.status,
                stage=self.stage,
                progress=self.progress,
                message=self.message,
                error=self.error,
                pptx_url=self.pptx_url,
                cancel_requested=self.cancel_requested,
                created_at=self.created_at,
                updated_at=self.updated_at,
            )


class TaskManager:
    def __init__(
        self,
        max_workers: int = MAX_CONCURRENT_TASKS,
        max_queued: int = MAX_QUEUED_TASKS,
        store: Optional[TaskStore] = None,
    ):
        self.store = store or TaskStore(TASK_DB_PATH)
        interrupted = self.store.mark_incomplete_interrupted()
        if interrupted:
            logger.warning("marked %s incomplete task(s) as interrupted", interrupted)
        self.store.delete_expired(time.time() - TASK_EXPIRE_HOURS * 3600)

        self._queue: queue.Queue = queue.Queue(maxsize=max_queued)
        self._active: Dict[str, TaskRecord] = {}
        self._active_lock = threading.Lock()
        self._submit_lock = threading.Lock()
        self._stop_event = threading.Event()
        self._accepting = True
        self._workers = []
        for index in range(max_workers):
            thread = threading.Thread(
                target=self._worker_loop,
                name=f"ppt-worker-{index + 1}",
                daemon=True,
            )
            thread.start()
            self._workers.append(thread)
        self._cleanup_thread = threading.Thread(
            target=self._cleanup_loop,
            name="ppt-task-cleanup",
            daemon=True,
        )
        self._cleanup_thread.start()

    @staticmethod
    def _request_hash(req: TaskCreate, upload: Optional[Tuple[bytes, str]]) -> str:
        payload = req.model_dump(mode="json")
        # 幂等判断需要区分不同密钥，但数据库中不能保存任务级明文密钥。
        payload["language_api_key"] = hashlib.sha256(
            req.language_api_key.get_secret_value().encode("utf-8")
        ).hexdigest()
        payload["vision_api_key"] = hashlib.sha256(
            req.vision_api_key.get_secret_value().encode("utf-8")
        ).hexdigest()
        if upload is not None:
            payload["upload"] = {
                "filename": upload[1],
                "sha256": hashlib.sha256(upload[0]).hexdigest(),
                "bytes": len(upload[0]),
            }
        raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()

    def submit(
        self,
        req: TaskCreate,
        upload: Optional[Tuple[bytes, str]] = None,
        idempotency_key: Optional[str] = None,
    ) -> Tuple[str, bool]:
        if not self._accepting:
            raise TaskQueueFull("服务正在停止，暂不接收新任务")
        key = (idempotency_key or "").strip() or None
        if key and len(key) > 128:
            raise ValueError("Idempotency-Key 不能超过 128 个字符")
        request_hash = self._request_hash(req, upload)

        with self._submit_lock:
            if key:
                existing = self.store.find_by_idempotency_key(key)
                if existing:
                    if existing.get("request_hash") != request_hash:
                        raise IdempotencyConflict("Idempotency-Key 已用于不同请求")
                    return str(existing["task_id"]), True
            if self._queue.full():
                raise TaskQueueFull(f"任务队列已满（最多等待 {self._queue.maxsize} 个）")

            task_id = uuid.uuid4().hex[:12]
            record = TaskRecord(task_id, self.store)
            try:
                self.store.create(
                    task_id,
                    idempotency_key=key,
                    request_hash=request_hash,
                )
            except DuplicateIdempotencyKey:
                existing = self.store.find_by_idempotency_key(key or "")
                if existing and existing.get("request_hash") == request_hash:
                    return str(existing["task_id"]), True
                raise IdempotencyConflict("Idempotency-Key 已被占用") from None
            with self._active_lock:
                self._active[task_id] = record
            self._queue.put_nowait((record, req, upload))
            return task_id, False

    def get(self, task_id: str) -> Optional[TaskInfo]:
        row = self.store.get(task_id)
        return TaskInfo(**row) if row else None

    def get_result_path(self, task_id: str) -> Optional[str]:
        row = self.store.get(task_id)
        return str(row["pptx_abs"]) if row and row.get("pptx_abs") else None

    def cancel(self, task_id: str) -> Optional[TaskInfo]:
        current = self.get(task_id)
        if current is None or current.status in FINAL_STATUSES:
            return current
        with self._active_lock:
            record = self._active.get(task_id)
        if record is not None:
            record.request_cancel()
        else:
            self.store.update(
                task_id,
                status=STATUS_CANCELLED,
                stage="已取消",
                message="任务已取消",
                cancel_requested=True,
            )
        return self.get(task_id)

    def stats(self) -> Dict:
        with self._active_lock:
            active = len(self._active)
        return {
            "database_ok": self.store.ping(),
            "active": active,
            "queued": self._queue.qsize(),
            "queue_capacity": self._queue.maxsize,
            "statuses": self.store.counts(),
        }

    def _worker_loop(self) -> None:
        while not self._stop_event.is_set():
            try:
                job = self._queue.get(timeout=1)
            except queue.Empty:
                continue
            if job is None:
                self._queue.task_done()
                break
            record, req, upload = job
            try:
                self._run(record, req, upload)
            finally:
                with self._active_lock:
                    self._active.pop(record.task_id, None)
                self._queue.task_done()

    # ---------------------------------------------------------------- worker

    def _run(self, record: TaskRecord, req: TaskCreate, upload: Optional[Tuple[bytes, str]]) -> None:
        task_id = record.task_id
        ws: Optional[TaskWorkspace] = None
        try:
            record.raise_if_cancelled()
            record.set(STATUS_RUNNING, stage="初始化工程", progress=5)
            ws = TaskWorkspace(task_id)

            def progress(stage: str, pct: int, msg: str) -> None:
                record.raise_if_cancelled()
                record.set(STATUS_RUNNING, stage=stage, progress=pct, message=msg)

            # 1) 源文档（PDF/DOCX/MD…）→ source_to_md 转换
            if upload is not None:
                progress("转换源文档", 10, f"转换上传文件 {upload[1]}…")
                src = ws.save_upload(upload[0], upload[1])
                convert_source(str(src), cwd=ws.staging_dir)

            # 2) 初始化工程（--quick-generate 创建 svg_output/；实际目录名带格式与日期后缀）
            progress("初始化工程", 15, "创建工程结构…")
            ws.project_dir = init_project(ws.project_dir, fmt=req.format)
            record.set(
                STATUS_RUNNING,
                stage="初始化工程",
                progress=15,
                project_dir=str(ws.project_dir),
            )

            # 3) 导入源文档
            sources = list(ws.staging_dir.iterdir()) if ws.staging_dir.is_dir() else []
            if sources:
                progress("导入资料", 18, f"导入 {len(sources)} 个源文件…")
                import_sources(ws.project_dir, sources)

            # 4) Agent 创作（真实 LLM 或 Mock）
            ctx = ToolContext(ws, req, progress)
            agent = MockAgent(ctx, progress) if MOCK_LLM else AgentLoop(ctx, progress)
            agent.run()

            if ctx.pptx_path is None or not ctx.pptx_path.is_file():
                raise RuntimeError("未找到导出的 PPTX 文件")

            record.raise_if_cancelled()
            record.set(
                STATUS_SUCCESS,
                stage="完成",
                progress=100,
                message=ctx.finish_message or f"生成完成：{ctx.pptx_path.name}",
                pptx_url=f"{PUBLIC_URL_PREFIX}/api/v1/tasks/{task_id}/result",
                pptx_abs=str(ctx.pptx_path),
                project_dir=str(ws.project_dir),
            )
            logger.info("task %s success: %s", task_id, ctx.pptx_path.name)
        except TaskCancelled:
            logger.info("task %s cancelled", task_id)
            record.set(
                STATUS_CANCELLED,
                stage="已取消",
                message="任务已取消",
                error="用户取消任务",
            )
        except Exception as exc:  # noqa: BLE001
            logger.exception("task %s failed", task_id)
            if record.cancel_requested:
                record.set(
                    STATUS_CANCELLED,
                    stage="已取消",
                    message="任务已取消",
                    error="用户取消任务",
                )
            else:
                record.set(STATUS_FAILED, stage="失败", error=str(exc)[:2000])
        finally:
            if ws is not None:
                ws.cleanup_staging()

    def _cleanup_loop(self) -> None:
        while not self._stop_event.wait(TASK_CLEANUP_INTERVAL_SECONDS):
            try:
                removed = self.store.delete_expired(time.time() - TASK_EXPIRE_HOURS * 3600)
                if removed:
                    logger.info("expired %s task metadata row(s)", removed)
            except Exception:  # noqa: BLE001
                logger.exception("task metadata cleanup failed")

    def shutdown(self) -> None:
        self._accepting = False
        self._stop_event.set()


manager = TaskManager()
