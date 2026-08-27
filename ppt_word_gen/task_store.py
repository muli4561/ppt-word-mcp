# coding=utf-8
"""SQLite task metadata store.

Only non-secret task state is persisted. Request-level model API keys stay in
the in-memory job payload and disappear when the worker finishes or restarts.
"""
import sqlite3
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, Optional


FINAL_STATUSES = ("success", "failed", "cancelled", "interrupted")
INCOMPLETE_STATUSES = ("pending", "running", "cancelling")


class DuplicateIdempotencyKey(RuntimeError):
    """Raised when a unique idempotency key already exists."""


class TaskStore:
    _UPDATABLE = {
        "status",
        "stage",
        "progress",
        "message",
        "error",
        "pptx_url",
        "pptx_abs",
        "project_dir",
        "cancel_requested",
        "updated_at",
    }

    def __init__(self, path: Path):
        self.path = Path(path).resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._write_lock = threading.RLock()
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.path), timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA busy_timeout=30000")
        return conn

    @contextmanager
    def _connection(self):
        conn = self._connect()
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _initialize(self) -> None:
        with self._write_lock, self._connection() as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS tasks (
                    task_id TEXT PRIMARY KEY,
                    status TEXT NOT NULL,
                    stage TEXT NOT NULL DEFAULT '',
                    progress INTEGER NOT NULL DEFAULT 0,
                    message TEXT NOT NULL DEFAULT '',
                    error TEXT,
                    pptx_url TEXT,
                    pptx_abs TEXT,
                    project_dir TEXT,
                    cancel_requested INTEGER NOT NULL DEFAULT 0,
                    idempotency_key TEXT,
                    request_hash TEXT,
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS ux_tasks_idempotency_key
                ON tasks(idempotency_key)
                WHERE idempotency_key IS NOT NULL
                """
            )

    @staticmethod
    def _as_dict(row: Optional[sqlite3.Row]) -> Optional[Dict]:
        if row is None:
            return None
        data = dict(row)
        data["cancel_requested"] = bool(data.get("cancel_requested"))
        return data

    def create(
        self,
        task_id: str,
        *,
        idempotency_key: Optional[str] = None,
        request_hash: Optional[str] = None,
    ) -> Dict:
        now = time.time()
        try:
            with self._write_lock, self._connection() as conn:
                conn.execute(
                    """
                    INSERT INTO tasks (
                        task_id, status, stage, progress, message,
                        idempotency_key, request_hash, created_at, updated_at
                    ) VALUES (?, 'pending', '排队中', 0, '', ?, ?, ?, ?)
                    """,
                    (task_id, idempotency_key, request_hash, now, now),
                )
        except sqlite3.IntegrityError as exc:
            if idempotency_key:
                raise DuplicateIdempotencyKey(idempotency_key) from exc
            raise
        result = self.get(task_id)
        if result is None:
            raise RuntimeError(f"任务写入后无法读取: {task_id}")
        return result

    def update(self, task_id: str, **fields) -> bool:
        unknown = set(fields) - self._UPDATABLE
        if unknown:
            raise ValueError(f"不允许更新任务字段: {', '.join(sorted(unknown))}")
        if not fields:
            return False
        fields.setdefault("updated_at", time.time())
        if "cancel_requested" in fields:
            fields["cancel_requested"] = int(bool(fields["cancel_requested"]))
        assignments = ", ".join(f"{name} = ?" for name in fields)
        values = [fields[name] for name in fields]
        with self._write_lock, self._connection() as conn:
            cursor = conn.execute(
                f"UPDATE tasks SET {assignments} WHERE task_id = ?",  # noqa: S608 - column allowlist above
                [*values, task_id],
            )
            return cursor.rowcount > 0

    def get(self, task_id: str) -> Optional[Dict]:
        with self._connection() as conn:
            row = conn.execute("SELECT * FROM tasks WHERE task_id = ?", (task_id,)).fetchone()
        return self._as_dict(row)

    def find_by_idempotency_key(self, key: str) -> Optional[Dict]:
        with self._connection() as conn:
            row = conn.execute(
                "SELECT * FROM tasks WHERE idempotency_key = ?", (key,)
            ).fetchone()
        return self._as_dict(row)

    def mark_incomplete_interrupted(self) -> int:
        placeholders = ",".join("?" for _ in INCOMPLETE_STATUSES)
        now = time.time()
        with self._write_lock, self._connection() as conn:
            cursor = conn.execute(
                f"""
                UPDATE tasks
                SET status = 'interrupted', stage = '服务已重启',
                    message = '任务未完成，可重新提交',
                    error = '服务重启导致任务中断', updated_at = ?
                WHERE status IN ({placeholders})
                """,  # noqa: S608 - placeholders are generated constants
                (now, *INCOMPLETE_STATUSES),
            )
            return cursor.rowcount

    def delete_expired(self, expire_before: float) -> int:
        placeholders = ",".join("?" for _ in FINAL_STATUSES)
        with self._write_lock, self._connection() as conn:
            cursor = conn.execute(
                f"DELETE FROM tasks WHERE status IN ({placeholders}) AND updated_at < ?",  # noqa: S608
                (*FINAL_STATUSES, expire_before),
            )
            return cursor.rowcount

    def counts(self) -> Dict[str, int]:
        with self._connection() as conn:
            rows = conn.execute(
                "SELECT status, COUNT(*) AS count FROM tasks GROUP BY status"
            ).fetchall()
        return {str(row["status"]): int(row["count"]) for row in rows}

    def ping(self) -> bool:
        try:
            with self._connection() as conn:
                return conn.execute("SELECT 1").fetchone()[0] == 1
        except sqlite3.Error:
            return False
