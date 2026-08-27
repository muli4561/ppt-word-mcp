import os
import base64
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from mcp import Client
from mcp.types import ResourceLink


_BOOT_TEMP = tempfile.TemporaryDirectory(prefix="mcp-server-boot-")
os.environ["TASK_DB_PATH"] = str(Path(_BOOT_TEMP.name) / "ppt.db")
os.environ["REPORT_TASK_DB_PATH"] = str(Path(_BOOT_TEMP.name) / "report.db")
os.environ["REPORT_OUTPUT_DIR"] = str(Path(_BOOT_TEMP.name) / "reports")
os.environ["UPLOAD_DIR"] = str(Path(_BOOT_TEMP.name) / "uploads")
os.environ["BUSINESS_TEMPLATE_DIR"] = str(Path(_BOOT_TEMP.name) / "templates")
os.environ["DOWNLOAD_SIGNING_SECRET_FILE"] = str(Path(_BOOT_TEMP.name) / "signing-secret")

import ppt_word_gen.mcp_server as module  # noqa: E402
from ppt_word_gen.report_models import ReportTaskInfo  # noqa: E402
from ppt_word_gen.task_store import TaskStore  # noqa: E402
from ppt_word_gen.tasks import TaskInfo, TaskManager  # noqa: E402


class MCPServerTests(unittest.IsolatedAsyncioTestCase):
    async def test_tools_are_discoverable(self):
        async with Client(module.mcp_server, raise_exceptions=True) as client:
            tools = await client.list_tools()
        names = {tool.name for tool in tools.tools}
        self.assertEqual(
            {
                "list_generation_profiles",
                "preview_word_report_format",
                "generate_presentation",
                "generate_word_report",
                "get_generation_task",
                "cancel_generation_task",
                "get_artifact",
                "upload_file",
                "create_upload_ticket",
                "wait_generation_task",
                "list_business_templates",
                "register_business_template",
                "delete_business_template",
                "revise_presentation",
                "revise_word_report",
            },
            names,
        )

    async def test_resources_are_discoverable(self):
        async with Client(module.mcp_server, raise_exceptions=True) as client:
            resources = await client.list_resources()
        uris = {str(resource.uri) for resource in resources.resources}
        self.assertEqual(
            {
                "ppt-word://rules/workflow",
                "ppt-word://rules/presentation",
                "ppt-word://rules/word-report",
                "ppt-word://templates/catalog",
            },
            uris,
        )

    async def test_inline_upload_and_structured_error(self):
        async with Client(module.mcp_server, raise_exceptions=False) as client:
            upload = await client.call_tool(
                "upload_file",
                {
                    "filename": "evidence.md",
                    "purpose": "source",
                    "content_base64": base64.b64encode(b"# evidence").decode("ascii"),
                },
            )
            missing = await client.call_tool(
                "get_generation_task",
                {"task_type": "presentation", "task_id": "missing"},
            )
        self.assertFalse(upload.is_error)
        self.assertEqual("source_upload_id", upload.structured_content["use_in_mcp_as"])
        self.assertTrue(missing.is_error)
        self.assertEqual("task_not_found", missing.structured_content["error"]["code"])

    async def test_business_template_lifecycle(self):
        async with Client(module.mcp_server, raise_exceptions=True) as client:
            created = await client.call_tool(
                "register_business_template",
                {
                    "name": "仿真周报",
                    "document_type": "presentation",
                    "instructions": "突出本周验证结果和下周计划",
                },
            )
            template_id = created.structured_content["template_id"]
            listed = await client.call_tool(
                "list_business_templates",
                {"document_type": "presentation"},
            )
            deleted = await client.call_tool(
                "delete_business_template",
                {"template_id": template_id},
            )
        ids = {item["template_id"] for item in listed.structured_content["templates"]}
        self.assertIn(template_id, ids)
        self.assertTrue(deleted.structured_content["deleted"])

    async def test_generate_presentation_uses_existing_queue(self):
        with tempfile.TemporaryDirectory(prefix="mcp-ppt-manager-") as directory:
            manager = TaskManager(
                max_workers=0,
                max_queued=2,
                store=TaskStore(Path(directory) / "tasks.db"),
            )
            try:
                with patch.object(module, "manager", manager):
                    async with Client(module.mcp_server, raise_exceptions=True) as client:
                        result = await client.call_tool(
                            "generate_presentation",
                            {
                                "topic": "MCP 架构说明",
                                "page_count": 5,
                                "idempotency_key": "mcp-test-1",
                            },
                        )
                self.assertFalse(result.is_error)
                payload = result.structured_content
                self.assertEqual("presentation", payload["task_type"])
                self.assertEqual("pending", manager.get(payload["task_id"]).status)
            finally:
                manager.shutdown()

    async def test_generate_tools_expose_and_forward_model_overrides(self):
        class FakePptManager:
            submitted = None

            def submit(self, request, upload, idempotency_key=None):
                self.submitted = (request, upload, idempotency_key)
                return "ppt-model-task", False

        class FakeReportManager:
            submitted = None

            def submit(self, request, source_upload, template_upload, idempotency_key=None):
                self.submitted = (request, source_upload, template_upload, idempotency_key)
                return "word-model-task", False

        ppt_manager = FakePptManager()
        report_manager = FakeReportManager()
        with (
            patch.object(module, "manager", ppt_manager),
            patch.object(module, "report_manager", report_manager),
        ):
            async with Client(module.mcp_server, raise_exceptions=True) as client:
                tools = await client.list_tools()
                confirmation = await client.call_tool(
                    "preview_word_report_format",
                    {"body_font": "宋体", "body_size_pt": 12.5, "numbering_style": "chinese"},
                )
                ppt = await client.call_tool(
                    "generate_presentation",
                    {
                        "topic": "模型切换验证",
                        "model": "qwen-plus",
                        "base_url": "https://dashscope.example.com/compatible-mode/v1/",
                        "api_key": "ppt-secret",
                        "temperature": 0.25,
                    },
                )
                word = await client.call_tool(
                    "generate_word_report",
                    {
                        "format_confirmation_token": confirmation.structured_content["confirmation_token"],
                        "instructions": "生成模型切换验证报告",
                        "model": "deepseek-chat",
                        "base_url": "https://api.deepseek.example/v1/",
                        "api_key": "word-secret",
                        "temperature": 0.75,
                    },
                )

        schemas = {tool.name: tool.input_schema for tool in tools.tools}
        for tool_name in ("generate_presentation", "generate_word_report"):
            properties = schemas[tool_name]["properties"]
            self.assertTrue({"model", "base_url", "api_key", "temperature"} <= properties.keys())
            self.assertNotIn("model", schemas[tool_name].get("required", []))
        self.assertIn(
            "format_confirmation_token",
            schemas["generate_word_report"].get("required", []),
        )
        self.assertTrue(
            {"template_upload_id", "body_font", "body_size_pt", "heading1_font", "numbering_style"}
            <= schemas["preview_word_report_format"]["properties"].keys()
        )

        self.assertFalse(ppt.is_error)
        ppt_request = ppt_manager.submitted[0]
        self.assertEqual("qwen-plus", ppt_request.language_model)
        self.assertEqual("https://dashscope.example.com/compatible-mode/v1", ppt_request.language_base_url)
        self.assertEqual("ppt-secret", ppt_request.language_api_key.get_secret_value())
        self.assertEqual(0.25, ppt_request.language_temperature)

        self.assertFalse(word.is_error)
        word_request = report_manager.submitted[0]
        self.assertEqual("deepseek-chat", word_request.language_model)
        self.assertEqual("https://api.deepseek.example/v1", word_request.language_base_url)
        self.assertEqual("word-secret", word_request.language_api_key.get_secret_value())
        self.assertEqual(0.75, word_request.language_temperature)
        self.assertEqual("宋体", word_request.word_format.body_font)
        self.assertEqual(12.5, word_request.word_format.body_size_pt)
        self.assertEqual("chinese", word_request.word_format.numbering_style)

    async def test_artifact_returns_signed_resource_link(self):
        with tempfile.TemporaryDirectory(prefix="mcp-artifact-") as directory:
            path = Path(directory) / "result.pptx"
            path.write_bytes(b"pptx")

            class FakeManager:
                def get(self, task_id):
                    return TaskInfo(task_id=task_id, status="success", progress=100)

                def get_result_path(self, task_id):
                    return str(path)

            with patch.object(module, "manager", FakeManager()):
                async with Client(module.mcp_server, raise_exceptions=True) as client:
                    result = await client.call_tool(
                        "get_artifact",
                        {"task_type": "presentation", "task_id": "source-task"},
                    )
            self.assertFalse(result.is_error)
            self.assertIsInstance(result.content[0], ResourceLink)
            self.assertIn("/api/v1/artifacts/", str(result.content[0].uri))
            self.assertGreater(result.structured_content["expires_at"], 0)

    async def test_semantic_revision_creates_new_tasks(self):
        with tempfile.TemporaryDirectory(prefix="mcp-revise-") as directory:
            pptx_path = Path(directory) / "old.pptx"
            docx_path = Path(directory) / "old.docx"
            pptx_path.write_bytes(b"pptx-source")
            docx_path.write_bytes(b"docx-source")

            class FakePptManager:
                submitted = None

                def get(self, task_id):
                    return TaskInfo(task_id=task_id, status="success", progress=100)

                def get_result_path(self, task_id):
                    return str(pptx_path)

                def submit(self, request, upload, idempotency_key=None):
                    self.submitted = (request, upload, idempotency_key)
                    return "new-ppt-task", False

            class FakeReportManager:
                submitted = None

                def get(self, task_id):
                    return ReportTaskInfo(task_id=task_id, status="success", progress=100)

                def get_result_path(self, task_id):
                    return str(docx_path)

                def submit(self, request, source_upload, template_upload, idempotency_key=None):
                    self.submitted = (request, source_upload, template_upload, idempotency_key)
                    return "new-word-task", False

            ppt_manager = FakePptManager()
            report_manager = FakeReportManager()
            with (
                patch.object(module, "manager", ppt_manager),
                patch.object(module, "report_manager", report_manager),
            ):
                async with Client(module.mcp_server, raise_exceptions=True) as client:
                    confirmation = await client.call_tool("preview_word_report_format", {})
                    ppt = await client.call_tool(
                        "revise_presentation",
                        {
                            "source_task_id": "old-ppt",
                            "instructions": "把结论页改为三条验收结论",
                            "page_count": 6,
                        },
                    )
                    word = await client.call_tool(
                        "revise_word_report",
                        {
                            "format_confirmation_token": confirmation.structured_content["confirmation_token"],
                            "source_task_id": "old-word",
                            "instructions": "补充部署回滚步骤",
                        },
                    )
                    waited = await client.call_tool(
                        "wait_generation_task",
                        {
                            "task_type": "presentation",
                            "task_id": "old-ppt",
                            "timeout_seconds": 1,
                        },
                    )
            self.assertEqual("new-ppt-task", ppt.structured_content["task_id"])
            self.assertEqual("old.pptx", ppt_manager.submitted[1][1])
            self.assertEqual("new-word-task", word.structured_content["task_id"])
            self.assertEqual("old.docx", report_manager.submitted[1][1])
            self.assertEqual("success", waited.structured_content["status"])


def tearDownModule():
    module.manager.shutdown()
    module.report_manager.shutdown()
    _BOOT_TEMP.cleanup()


if __name__ == "__main__":
    unittest.main()
