# coding=utf-8
"""MCP/REST 共用的短期上传文件存储。"""
import json
import re
import shutil
import threading
import time
import uuid
from pathlib import Path
from typing import Dict, Optional, Tuple

from .config import UPLOAD_DIR, UPLOAD_EXPIRE_HOURS
from .report_documents import safe_filename


_ID_RE = re.compile(r"^[0-9a-f]{16}$")
_PURPOSES = {"source", "reference_template"}


class UploadNotFound(ValueError):
    pass


class UploadPurposeMismatch(ValueError):
    pass


class UploadStore:
    def __init__(self, root: Path = UPLOAD_DIR, expire_hours: int = UPLOAD_EXPIRE_HOURS):
        self.root = Path(root).resolve()
        self.expire_seconds = max(3600, int(expire_hours) * 3600)
        self.root.mkdir(parents=True, exist_ok=True)
        self._ticket_lock = threading.Lock()
        self._ticket_dir = self.root / ".consumed_tickets"
        self._ticket_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _validate_id(upload_id: str) -> str:
        value = (upload_id or "").strip().lower()
        if not _ID_RE.fullmatch(value):
            raise UploadNotFound("上传文件不存在")
        return value

    def _directory(self, upload_id: str) -> Path:
        value = self._validate_id(upload_id)
        directory = (self.root / value).resolve()
        if directory.parent != self.root:
            raise UploadNotFound("上传文件不存在")
        return directory

    def put(self, data: bytes, filename: str, purpose: str) -> Dict:
        if purpose not in _PURPOSES:
            raise ValueError("purpose 仅支持 source / reference_template")
        if not data:
            raise ValueError("上传文件为空")
        clean_name = safe_filename(filename, "upload")
        original_suffix = Path(filename).suffix.lower()
        if original_suffix and not Path(clean_name).suffix:
            clean_name = f"upload{original_suffix}"
        if purpose == "reference_template" and Path(clean_name).suffix.lower() not in {".docx", ".dotx"}:
            raise ValueError("参考模板必须是 .docx 或 .dotx 文件")

        self.cleanup_expired()
        upload_id = uuid.uuid4().hex[:16]
        directory = self._directory(upload_id)
        directory.mkdir(parents=False, exist_ok=False)
        path = directory / clean_name
        part = directory / (clean_name + ".part")
        part.write_bytes(data)
        part.replace(path)
        metadata = {
            "upload_id": upload_id,
            "filename": clean_name,
            "purpose": purpose,
            "bytes": len(data),
            "created_at": time.time(),
            "expires_at": time.time() + self.expire_seconds,
        }
        (directory / "metadata.json").write_text(
            json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return metadata

    def get(self, upload_id: str, expected_purpose: Optional[str] = None) -> Tuple[bytes, str, Dict]:
        directory = self._directory(upload_id)
        metadata_path = directory / "metadata.json"
        if not metadata_path.is_file():
            raise UploadNotFound("上传文件不存在或已过期")
        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        if float(metadata.get("expires_at", 0)) <= time.time():
            self.delete(upload_id)
            raise UploadNotFound("上传文件已过期，请重新上传")
        if expected_purpose and metadata.get("purpose") != expected_purpose:
            raise UploadPurposeMismatch(
                f"上传文件用途应为 {expected_purpose}，实际为 {metadata.get('purpose', 'unknown')}"
            )
        path = (directory / str(metadata["filename"])).resolve()
        if path.parent != directory or not path.is_file():
            raise UploadNotFound("上传文件内容不存在")
        return path.read_bytes(), str(metadata["filename"]), metadata

    def info(self, upload_id: str) -> Dict:
        _, _, metadata = self.get(upload_id)
        return metadata

    def path(self, upload_id: str) -> Path:
        """返回已校验的内部路径，仅供服务端预检使用。"""
        _, filename, _ = self.get(upload_id)
        return self._directory(upload_id) / filename

    def delete(self, upload_id: str) -> bool:
        directory = self._directory(upload_id)
        if not directory.is_dir():
            return False
        shutil.rmtree(directory)
        return True

    def cleanup_expired(self) -> int:
        removed = 0
        now = time.time()
        for directory in self.root.iterdir():
            if not directory.is_dir() or not _ID_RE.fullmatch(directory.name):
                continue
            metadata_path = directory / "metadata.json"
            try:
                metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
                expired = float(metadata.get("expires_at", 0)) <= now
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                expired = directory.stat().st_mtime + self.expire_seconds <= now
            if expired:
                shutil.rmtree(directory)
                removed += 1
        return removed

    def consume_ticket(self, ticket_id: str) -> None:
        """原子标记一次性上传票据；重复使用会被拒绝。"""
        value = (ticket_id or "").strip().lower()
        if not _ID_RE.fullmatch(value):
            raise ValueError("上传票据无效")
        marker = self._ticket_dir / value
        with self._ticket_lock:
            try:
                with marker.open("x", encoding="utf-8") as handle:
                    handle.write(str(time.time()))
            except FileExistsError as exc:
                raise ValueError("上传票据已使用") from exc


upload_store = UploadStore()
