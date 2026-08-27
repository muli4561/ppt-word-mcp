import tempfile
import time
import unittest
from pathlib import Path

from ppt_word_gen.task_store import DuplicateIdempotencyKey, TaskStore


class TaskStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(prefix="ppt-task-store-")
        self.store = TaskStore(Path(self.temp_dir.name) / "tasks.db")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_create_update_and_read(self):
        self.store.create("task-1", idempotency_key="key-1", request_hash="hash-1")
        self.store.update(
            "task-1",
            status="success",
            stage="完成",
            progress=100,
            pptx_abs="/data/result.pptx",
        )

        row = self.store.get("task-1")
        self.assertEqual("success", row["status"])
        self.assertEqual(100, row["progress"])
        self.assertEqual("/data/result.pptx", row["pptx_abs"])
        self.assertEqual("task-1", self.store.find_by_idempotency_key("key-1")["task_id"])

    def test_idempotency_key_is_unique(self):
        self.store.create("task-1", idempotency_key="same-key", request_hash="hash-1")
        with self.assertRaises(DuplicateIdempotencyKey):
            self.store.create("task-2", idempotency_key="same-key", request_hash="hash-2")

    def test_restart_marks_incomplete_tasks_interrupted(self):
        self.store.create("pending-task")
        self.store.create("running-task")
        self.store.update("running-task", status="running", stage="AI 创作")
        self.store.create("success-task")
        self.store.update("success-task", status="success", progress=100)

        changed = self.store.mark_incomplete_interrupted()

        self.assertEqual(2, changed)
        self.assertEqual("interrupted", self.store.get("pending-task")["status"])
        self.assertEqual("interrupted", self.store.get("running-task")["status"])
        self.assertEqual("success", self.store.get("success-task")["status"])

    def test_cleanup_removes_only_expired_final_tasks(self):
        self.store.create("old-success")
        self.store.update(
            "old-success",
            status="success",
            progress=100,
            updated_at=time.time() - 7200,
        )
        self.store.create("old-running")
        self.store.update(
            "old-running",
            status="running",
            updated_at=time.time() - 7200,
        )

        removed = self.store.delete_expired(time.time() - 3600)

        self.assertEqual(1, removed)
        self.assertIsNone(self.store.get("old-success"))
        self.assertIsNotNone(self.store.get("old-running"))


if __name__ == "__main__":
    unittest.main()
