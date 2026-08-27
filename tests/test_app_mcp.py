import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


_BOOT_TEMP = tempfile.TemporaryDirectory(prefix="app-mcp-boot-")
os.environ["TASK_DB_PATH"] = str(Path(_BOOT_TEMP.name) / "ppt.db")
os.environ["REPORT_TASK_DB_PATH"] = str(Path(_BOOT_TEMP.name) / "report.db")
os.environ["REPORT_OUTPUT_DIR"] = str(Path(_BOOT_TEMP.name) / "reports")
os.environ["UPLOAD_DIR"] = str(Path(_BOOT_TEMP.name) / "uploads")
os.environ["BUSINESS_TEMPLATE_DIR"] = str(Path(_BOOT_TEMP.name) / "templates")
os.environ["DOWNLOAD_SIGNING_SECRET_FILE"] = str(Path(_BOOT_TEMP.name) / "signing-secret")
os.environ["PPT_WORD_GEN_TOKEN"] = ""
os.environ["MOCK_LLM"] = "1"

from fastapi.testclient import TestClient  # noqa: E402
import ppt_word_gen.app as app_module  # noqa: E402
from ppt_word_gen.app import app, manager  # noqa: E402
from ppt_word_gen.config import WORD_REPORT_TEMPLATE_PATH  # noqa: E402
from ppt_word_gen.signed_tokens import sign_claims  # noqa: E402


class AppMCPTests(unittest.TestCase):
    def test_health_upload_and_streamable_http_tools(self):
        with TestClient(app) as client:
            health = client.get("/health")
            self.assertEqual(200, health.status_code)
            self.assertEqual("/mcp", health.json()["mcp"]["endpoint"])

            upload = client.post(
                "/api/v1/uploads",
                data={"purpose": "source"},
                files={"file": ("source.md", b"# evidence", "text/markdown")},
            )
            self.assertEqual(200, upload.status_code, upload.text)
            self.assertEqual("source_upload_id", upload.json()["use_in_mcp_as"])

            response = client.post(
                "/mcp",
                headers={
                    "Accept": "application/json, text/event-stream",
                    "MCP-Protocol-Version": "2026-07-28",
                    "MCP-Method": "tools/list",
                },
                json={
                    "jsonrpc": "2.0",
                    "id": 1,
                    "method": "tools/list",
                    "params": {
                        "_meta": {
                            "io.modelcontextprotocol/protocolVersion": "2026-07-28",
                            "io.modelcontextprotocol/clientCapabilities": {},
                        }
                    },
                },
            )
            self.assertEqual(200, response.status_code, response.text)
            tools = response.json()["result"]["tools"]
            self.assertIn("generate_presentation", {tool["name"] for tool in tools})

    def test_signed_upload_ticket_is_one_time(self):
        token = sign_claims(
            "upload",
            {
                "ticket_id": "0123456789abcdef",
                "filename": "evidence.md",
                "purpose": "source",
                "max_bytes": 1024,
            },
        )
        client = TestClient(app)
        response = client.put(
            f"/api/v1/upload-tickets/{token}",
            content=b"# evidence",
            headers={"Content-Type": "application/octet-stream"},
        )
        self.assertEqual(200, response.status_code, response.text)
        self.assertEqual("source_upload_id", response.json()["use_in_mcp_as"])
        replay = client.put(f"/api/v1/upload-tickets/{token}", content=b"again")
        self.assertEqual(400, replay.status_code)

    def test_word_report_requires_current_format_confirmation(self):
        client = TestClient(app)
        profile = client.get("/api/v1/word-format")
        self.assertEqual(200, profile.status_code, profile.text)
        payload = profile.json()
        self.assertEqual("cid629-joint-simulation-v1.5", payload["profile_id"])
        self.assertIn("一级标题", payload["confirmation_text"])

        missing = client.post(
            "/api/v1/report-tasks",
            data={"instructions": "生成联合仿真测试报告"},
        )
        self.assertEqual(422, missing.status_code)

        with patch.object(app_module.report_manager, "submit", return_value=("word-test", False)):
            accepted = client.post(
                "/api/v1/report-tasks",
                data={
                    "instructions": "生成联合仿真测试报告",
                    "format_confirmation_token": payload["confirmation_token"],
                },
            )
        self.assertEqual(200, accepted.status_code, accepted.text)
        self.assertIn("task_id", accepted.json())

    def test_uploaded_template_supplies_editable_format_defaults(self):
        client = TestClient(app)
        upload = client.post(
            "/api/v1/uploads",
            data={"purpose": "reference_template"},
            files={
                "file": (
                    "custom.docx",
                    WORD_REPORT_TEMPLATE_PATH.read_bytes(),
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            },
        )
        self.assertEqual(200, upload.status_code, upload.text)
        upload_id = upload.json()["upload_id"]
        preview = client.post(
            "/api/v1/word-format",
            json={
                "template_upload_id": upload_id,
                "format": {"body_font": "宋体", "body_size_pt": 12.5, "numbering_style": "chinese"},
            },
        )
        self.assertEqual(200, preview.status_code, preview.text)
        payload = preview.json()
        self.assertEqual("uploaded", payload["template_source"])
        self.assertEqual("宋体", payload["format"]["body_font"])
        self.assertEqual(12.5, payload["format"]["body_size_pt"])
        self.assertEqual("chinese", payload["format"]["numbering_style"])

    def test_signed_artifact_download_needs_no_bearer_header(self):
        task_id = "signedtest01"
        artifact = Path(_BOOT_TEMP.name) / "signed-test.pptx"
        artifact.write_bytes(b"fake-pptx-for-route-test")
        manager.store.create(task_id)
        manager.store.update(
            task_id,
            status="success",
            stage="完成",
            progress=100,
            pptx_abs=str(artifact),
        )
        token = sign_claims(
            "artifact",
            {"task_type": "presentation", "task_id": task_id},
        )
        client = TestClient(app)
        response = client.get(f"/api/v1/artifacts/{token}")
        self.assertEqual(200, response.status_code, response.text)
        self.assertEqual(artifact.read_bytes(), response.content)
        tampered = client.get(f"/api/v1/artifacts/{token}x")
        self.assertEqual(403, tampered.status_code)


def tearDownModule():
    # 其他测试模块与 app 共享已导入的全局 manager/store；进程退出时自动清理。
    pass


if __name__ == "__main__":
    unittest.main()
