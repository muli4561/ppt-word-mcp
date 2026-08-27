# coding=utf-8
"""Word 报告任务队列；与 PPT 任务使用独立状态库。"""
import hashlib
import json
import logging
import queue
import threading
import time
import uuid
from pathlib import Path
from typing import Dict, Optional, Tuple

from .config import (
    PUBLIC_URL_PREFIX,
    MAX_CONCURRENT_REPORT_TASKS,
    MAX_QUEUED_REPORT_TASKS,
    MOCK_LLM,
    REPORT_OUTPUT_DIR,
    REPORT_TASK_DB_PATH,
    TASK_CLEANUP_INTERVAL_SECONDS,
    TASK_EXPIRE_HOURS,
    WORD_REPORT_TEMPLATE_PATH,
)
from .report_agent import ReportAgent, build_mock_spec
from .report_documents import (
    build_evidence_manifest,
    extract_source,
    render_report,
    validate_docx_package,
    validate_rendered_report,
    write_json,
    write_upload_atomic,
)
from .report_models import ReportCreate, ReportTaskInfo
from .task_store import FINAL_STATUSES, DuplicateIdempotencyKey, TaskStore

logger = logging.getLogger("ppt_word_gen.report_tasks")

Upload = Tuple[bytes, str]
STATUS_PENDING = "pending"
STATUS_RUNNING = "running"
STATUS_SUCCESS = "success"
STATUS_FAILED = "failed"
STATUS_CANCELLING = "cancelling"
STATUS_CANCELLED = "cancelled"


class ReportQueueFull(RuntimeError):
    pass


class ReportIdempotencyConflict(RuntimeError):
    pass


class ReportCancelled(RuntimeError):
    pass


class ReportTaskRecord:
    def __init__(self, task_id: str, store: TaskStore):
        self.task_id = task_id
        self._store = store
        self._lock = threading.Lock()
        self.status = STATUS_PENDING
        self.stage = "排队中"
        self.progress = 0
        self.message = ""
        self.error: Optional[str] = None
        self.document_url: Optional[str] = None
        self.document_abs: Optional[str] = None
        self.project_dir: Optional[str] = None
        self.cancel_requested = False
        self.created_at = time.time()
        self.updated_at = self.created_at

    def set(
        self,
        status: str,
        *,
        stage: str = "",
        progress: Optional[int] = None,
        message: str = "",
        error: Optional[str] = None,
        document_url: Optional[str] = None,
        document_abs: Optional[str] = None,
        project_dir: Optional[str] = None,
    ) -> None:
        with self._lock:
            if self.cancel_requested and status in {STATUS_PENDING, STATUS_RUNNING}:
                raise ReportCancelled("report task cancellation requested")
            self.status = status
            if stage:
                self.stage = stage
            if progress is not None:
                self.progress = progress
            if message:
                self.message = message
            if error is not None:
                self.error = error
            if document_url is not None:
                self.document_url = document_url
            if document_abs is not None:
                self.document_abs = document_abs
            if project_dir is not None:
                self.project_dir = project_dir
            self.updated_at = time.time()
            # 使用独立 report_tasks.db；复用 TaskStore 的旧字段名仅是内部存储兼容。
            fields = {
                "status": self.status,
                "stage": self.stage,
                "progress": self.progress,
                "message": self.message,
                "error": self.error,
                "pptx_url": self.document_url,
                "pptx_abs": self.document_abs,
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
                raise ReportCancelled("report task cancellation requested")


class ReportTaskManager:
    def __init__(
        self,
        max_workers: int = MAX_CONCURRENT_REPORT_TASKS,
        max_queued: int = MAX_QUEUED_REPORT_TASKS,
        store: Optional[TaskStore] = None,
        output_dir: Path = REPORT_OUTPUT_DIR,
        mock_llm: bool = MOCK_LLM,
    ):
        self.store = store or TaskStore(REPORT_TASK_DB_PATH)
        interrupted = self.store.mark_incomplete_interrupted()
        if interrupted:
            logger.warning("marked %s incomplete report task(s) as interrupted", interrupted)
        self.store.delete_expired(time.time() - TASK_EXPIRE_HOURS * 3600)
        self.output_dir = Path(output_dir).resolve()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.mock_llm = mock_llm
        self._queue: queue.Queue = queue.Queue(maxsize=max_queued)
        self._active: Dict[str, ReportTaskRecord] = {}
        self._active_lock = threading.Lock()
        self._submit_lock = threading.Lock()
        self._stop_event = threading.Event()
        self._accepting = True
        self._workers = []
        for index in range(max_workers):
            thread = threading.Thread(
                target=self._worker_loop,
                name=f"report-worker-{index + 1}",
                daemon=True,
            )
            thread.start()
            self._workers.append(thread)
        self._cleanup_thread = threading.Thread(
            target=self._cleanup_loop,
            name="report-task-cleanup",
            daemon=True,
        )
        self._cleanup_thread.start()

    @staticmethod
    def _request_hash(
        request: ReportCreate,
        source_upload: Optional[Upload],
        template_upload: Optional[Upload],
    ) -> str:
        payload = request.model_dump(mode="json")
        payload["language_api_key"] = hashlib.sha256(
            request.language_api_key.get_secret_value().encode("utf-8")
        ).hexdigest()
        for key, upload in (("source", source_upload), ("template", template_upload)):
            if upload:
                payload[key] = {
                    "filename": upload[1],
                    "bytes": len(upload[0]),
                    "sha256": hashlib.sha256(upload[0]).hexdigest(),
                }
        raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()

    def submit(
        self,
        request: ReportCreate,
        source_upload: Optional[Upload] = None,
        template_upload: Optional[Upload] = None,
        idempotency_key: Optional[str] = None,
    ) -> Tuple[str, bool]:
        if not self._accepting:
            raise ReportQueueFull("服务正在停止，暂不接收新任务")
        if not request.instructions and source_upload is None:
            raise ValueError("请填写报告要求或上传来源资料")
        key = (idempotency_key or "").strip() or None
        if key and len(key) > 128:
            raise ValueError("Idempotency-Key 不能超过 128 个字符")
        request_hash = self._request_hash(request, source_upload, template_upload)
        with self._submit_lock:
            if key:
                existing = self.store.find_by_idempotency_key(key)
                if existing:
                    if existing.get("request_hash") != request_hash:
                        raise ReportIdempotencyConflict("Idempotency-Key 已用于不同请求")
                    return str(existing["task_id"]), True
            if self._queue.full():
                raise ReportQueueFull(f"报告任务队列已满（最多等待 {self._queue.maxsize} 个）")

            task_id = uuid.uuid4().hex[:12]
            record = ReportTaskRecord(task_id, self.store)
            try:
                self.store.create(task_id, idempotency_key=key, request_hash=request_hash)
            except DuplicateIdempotencyKey:
                existing = self.store.find_by_idempotency_key(key or "")
                if existing and existing.get("request_hash") == request_hash:
                    return str(existing["task_id"]), True
                raise ReportIdempotencyConflict("Idempotency-Key 已被占用") from None
            with self._active_lock:
                self._active[task_id] = record
            self._queue.put_nowait((record, request, source_upload, template_upload))
            return task_id, False

    @staticmethod
    def _info_from_row(row: Dict) -> ReportTaskInfo:
        return ReportTaskInfo(
            task_id=row["task_id"],
            status=row["status"],
            stage=row.get("stage", ""),
            progress=row.get("progress", 0),
            message=row.get("message", ""),
            error=row.get("error"),
            document_url=row.get("pptx_url"),
            cancel_requested=row.get("cancel_requested", False),
            created_at=row.get("created_at", 0),
            updated_at=row.get("updated_at", 0),
        )

    def get(self, task_id: str) -> Optional[ReportTaskInfo]:
        row = self.store.get(task_id)
        return self._info_from_row(row) if row else None

    def get_result_path(self, task_id: str) -> Optional[str]:
        row = self.store.get(task_id)
        return str(row["pptx_abs"]) if row and row.get("pptx_abs") else None

    def cancel(self, task_id: str) -> Optional[ReportTaskInfo]:
        current = self.get(task_id)
        if current is None or current.status in FINAL_STATUSES:
            return current
        with self._active_lock:
            record = self._active.get(task_id)
        if record:
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
            record, request, source_upload, template_upload = job
            try:
                self._run(record, request, source_upload, template_upload)
            finally:
                with self._active_lock:
                    self._active.pop(record.task_id, None)
                self._queue.task_done()

    def _run(
        self,
        record: ReportTaskRecord,
        request: ReportCreate,
        source_upload: Optional[Upload],
        template_upload: Optional[Upload],
    ) -> None:
        task_id = record.task_id
        job_dir = self.output_dir / task_id
        try:
            record.raise_if_cancelled()
            job_dir.mkdir(parents=True, exist_ok=False)
            record.set(
                STATUS_RUNNING,
                stage="准备资料",
                progress=5,
                message="校验上传文件…",
                project_dir=str(job_dir),
            )
            source_path = None
            if source_upload:
                source_path = write_upload_atomic(
                    job_dir / "sources", source_upload[0], source_upload[1], "source"
                )
            template_path = Path(WORD_REPORT_TEMPLATE_PATH)
            if template_upload:
                template_path = write_upload_atomic(
                    job_dir / "template", template_upload[0], template_upload[1], "template.docx"
                )
                validate_docx_package(template_path)
            else:
                validate_docx_package(template_path)

            record.raise_if_cancelled()
            record.set(STATUS_RUNNING, stage="提取证据", progress=18, message="转换资料并建立证据清单…")
            source_text, images, source_metadata = extract_source(source_path, request.instructions)
            evidence = build_evidence_manifest(source_text, source_metadata)
            write_json(job_dir / "source_metadata.json", source_metadata)
            write_json(job_dir / "evidence_manifest.json", evidence)

            record.raise_if_cancelled()
            record.set(STATUS_RUNNING, stage="编制报告", progress=38, message="生成结构化 ReportSpec…")
            image_names = [path.name for path in images]
            if self.mock_llm:
                spec = build_mock_spec(request, evidence, image_names)
            else:
                spec = ReportAgent(request).run(evidence, image_names)
            write_json(job_dir / "report_spec.json", spec.model_dump(mode="json"))

            record.raise_if_cancelled()
            record.set(STATUS_RUNNING, stage="渲染 Word", progress=72, message="应用参考模板并生成 DOCX…")
            output_path = job_dir / f"report-{task_id}.docx"
            render_metadata = render_report(
                spec,
                output_path,
                job_dir,
                template_path,
                strict_reference=template_upload is None,
                format_options=request.word_format,
                format_profile_id=request.word_format_profile_id,
            )

            record.raise_if_cancelled()
            record.set(STATUS_RUNNING, stage="质量检查", progress=90, message="检查包结构、模板变量与证据引用…")
            validation = validate_rendered_report(output_path, spec, evidence, render_metadata)
            write_json(job_dir / "validation.json", validation)

            record.raise_if_cancelled()
            record.set(
                STATUS_SUCCESS,
                stage="完成",
                progress=100,
                message=f"Word 报告已生成：{output_path.name}",
                document_url=f"{PUBLIC_URL_PREFIX}/api/v1/report-tasks/{task_id}/result",
                document_abs=str(output_path),
            )
            logger.info("report task %s success: %s", task_id, output_path.name)
        except ReportCancelled:
            logger.info("report task %s cancelled", task_id)
            record.set(
                STATUS_CANCELLED,
                stage="已取消",
                message="任务已取消",
                error="用户取消任务",
            )
        except Exception as exc:  # noqa: BLE001
            logger.exception("report task %s failed", task_id)
            if record.cancel_requested:
                record.set(
                    STATUS_CANCELLED,
                    stage="已取消",
                    message="任务已取消",
                    error="用户取消任务",
                )
            else:
                record.set(STATUS_FAILED, stage="失败", error=str(exc)[:2000])

    def _cleanup_loop(self) -> None:
        while not self._stop_event.wait(TASK_CLEANUP_INTERVAL_SECONDS):
            try:
                self.store.delete_expired(time.time() - TASK_EXPIRE_HOURS * 3600)
            except Exception:  # noqa: BLE001
                logger.exception("report task metadata cleanup failed")

    def shutdown(self, wait: bool = False) -> None:
        self._accepting = False
        self._stop_event.set()
        if wait:
            for worker in self._workers:
                worker.join(timeout=10)
            self._cleanup_thread.join(timeout=2)


report_manager = ReportTaskManager()
