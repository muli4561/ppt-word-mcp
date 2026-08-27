import os
import tempfile
import unittest
from pathlib import Path


# tasks.py 会创建应用级 manager；先把它隔离到临时数据库，避免测试污染运行数据。
_BOOT_TEMP = tempfile.TemporaryDirectory(prefix="ppt-task-manager-boot-")
os.environ["TASK_DB_PATH"] = str(Path(_BOOT_TEMP.name) / "boot.db")

from ppt_word_gen.task_store import TaskStore  # noqa: E402
from ppt_word_gen.tasks import (  # noqa: E402
    IdempotencyConflict,
    TaskCreate,
    TaskManager,
    TaskQueueFull,
    manager as application_manager,
)


class TaskManagerTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(prefix="ppt-task-manager-")
        store = TaskStore(Path(self.temp_dir.name) / "tasks.db")
        # 不启动 worker，验证不会进入 PPT/LLM 生成链路。
        self.manager = TaskManager(max_workers=0, max_queued=1, store=store)

    def tearDown(self):
        self.manager.shutdown()
        self.temp_dir.cleanup()

    def test_idempotent_submit_reuses_same_task(self):
        request = TaskCreate(topic="阶段二测试", page_count=3)

        task_id, reused = self.manager.submit(request, idempotency_key="request-1")
        repeated_id, repeated = self.manager.submit(request, idempotency_key="request-1")

        self.assertFalse(reused)
        self.assertTrue(repeated)
        self.assertEqual(task_id, repeated_id)
        self.assertEqual("pending", self.manager.get(task_id).status)

    def test_same_key_with_different_request_conflicts(self):
        self.manager.submit(TaskCreate(topic="请求 A"), idempotency_key="request-2")

        with self.assertRaises(IdempotencyConflict):
            self.manager.submit(TaskCreate(topic="请求 B"), idempotency_key="request-2")

    def test_secret_change_is_part_of_idempotency_fingerprint(self):
        self.manager.submit(
            TaskCreate(topic="相同主题", language_api_key="secret-a"),
            idempotency_key="request-3",
        )

        with self.assertRaises(IdempotencyConflict):
            self.manager.submit(
                TaskCreate(topic="相同主题", language_api_key="secret-b"),
                idempotency_key="request-3",
            )

    def test_queue_limit_and_pending_cancellation(self):
        task_id, _ = self.manager.submit(TaskCreate(topic="排队任务"))

        with self.assertRaises(TaskQueueFull):
            self.manager.submit(TaskCreate(topic="超出队列"))

        cancelled = self.manager.cancel(task_id)
        self.assertEqual("cancelled", cancelled.status)
        self.assertTrue(cancelled.cancel_requested)


def tearDownModule():
    application_manager.shutdown()
    _BOOT_TEMP.cleanup()


if __name__ == "__main__":
    unittest.main()
