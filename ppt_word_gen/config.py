# coding=utf-8
"""全局配置：从 .env / 环境变量读取。"""
import os
from pathlib import Path

from dotenv import load_dotenv

PACKAGE_DIR = Path(__file__).resolve().parent
BASE_DIR = PACKAGE_DIR.parent
load_dotenv(BASE_DIR / ".env")


def _get(key: str, default: str = "") -> str:
    return os.environ.get(key, default)


# ---------- ppt-master 路径 ----------
PPTMASTER_ROOT = Path(_get("PPTMASTER_ROOT", str(BASE_DIR.parent / "ppt-master"))).resolve()
SKILL_DIR = PPTMASTER_ROOT / "skills" / "ppt-master"
SCRIPTS_DIR = SKILL_DIR / "scripts"
PROJECTS_DIR = PPTMASTER_ROOT / "projects"
EXAMPLES_DIR = PPTMASTER_ROOT / "examples"

# ---------- 创作 LLM（OpenAI 兼容）----------
LLM_BASE_URL = _get("LLM_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1").rstrip("/")
LLM_API_KEY = _get("LLM_API_KEY", "")
LLM_MODEL = _get("LLM_MODEL", "qwen3.7-plus")
LLM_TEMPERATURE = float(_get("LLM_TEMPERATURE", "0.7"))

# ---------- 可选视觉模型（沿用 PPT Master image_gen.py 适配层）----------
VISION_ENABLED = _get("VISION_ENABLED", "0") == "1"
VISION_BACKEND = _get("VISION_BACKEND", "openai").strip().lower() or "openai"
VISION_BASE_URL = _get("VISION_BASE_URL", "").rstrip("/")
VISION_API_KEY = _get("VISION_API_KEY", "")
VISION_MODEL = _get("VISION_MODEL", "")
VISION_IMAGE_SIZE = _get("VISION_IMAGE_SIZE", "1K")
MAX_GENERATED_IMAGES = max(1, min(30, int(_get("MAX_GENERATED_IMAGES", "8"))))

# ---------- 服务 ----------
PPT_WORD_GEN_TOKEN = _get("PPT_WORD_GEN_TOKEN", "")
# 对外部署在子路径时设置；直接监听 8000 时保持为空。
PUBLIC_URL_PREFIX = _get("PUBLIC_URL_PREFIX", "").rstrip("/")
# MCP 工具返回下载地址时使用；留空则返回当前服务内的相对路径。
PUBLIC_BASE_URL = _get("PUBLIC_BASE_URL", "").rstrip("/")
UPLOAD_DIR = Path(_get("UPLOAD_DIR", str(BASE_DIR / "data" / "uploads"))).resolve()
UPLOAD_EXPIRE_HOURS = max(1, int(_get("UPLOAD_EXPIRE_HOURS", "24")))
MAX_UPLOAD_MB = max(1, int(_get("MAX_UPLOAD_MB", "20")))
BUSINESS_TEMPLATE_DIR = Path(
    _get("BUSINESS_TEMPLATE_DIR", str(BASE_DIR / "data" / "business_templates"))
).resolve()
# MCP 内联传输采用 Base64，默认限制得比普通 multipart 上传更小，避免大 JSON
# 长时间占用 Agent 上下文；大文件可改用一次性上传票据。
MCP_INLINE_UPLOAD_MB = max(
    1,
    min(MAX_UPLOAD_MB, int(_get("MCP_INLINE_UPLOAD_MB", "5"))),
)
SIGNED_URL_EXPIRE_SECONDS = max(60, int(_get("SIGNED_URL_EXPIRE_SECONDS", "86400")))
DOWNLOAD_SIGNING_SECRET = _get("DOWNLOAD_SIGNING_SECRET", "")
DOWNLOAD_SIGNING_SECRET_FILE = Path(
    _get(
        "DOWNLOAD_SIGNING_SECRET_FILE",
        str(BASE_DIR / "data" / ".download_signing_secret"),
    )
).resolve()
# 内网部署可保持关闭；公网部署建议开启并配置 MCP_ALLOWED_HOSTS。
MCP_DNS_REBINDING_PROTECTION = _get("MCP_DNS_REBINDING_PROTECTION", "0") == "1"
MCP_ALLOWED_HOSTS = [
    item.strip() for item in _get("MCP_ALLOWED_HOSTS", "localhost,localhost:*,127.0.0.1,127.0.0.1:*").split(",")
    if item.strip()
]
MCP_ALLOWED_ORIGINS = [
    item.strip() for item in _get("MCP_ALLOWED_ORIGINS", "").split(",") if item.strip()
]

# ---------- 任务 ----------
MAX_CONCURRENT_TASKS = max(1, int(_get("MAX_CONCURRENT_TASKS", "2")))
MAX_QUEUED_TASKS = max(1, int(_get("MAX_QUEUED_TASKS", "20")))
TASK_EXPIRE_HOURS = max(1, int(_get("TASK_EXPIRE_HOURS", "24")))
TASK_CLEANUP_INTERVAL_SECONDS = max(60, int(_get("TASK_CLEANUP_INTERVAL_SECONDS", "600")))
TASK_DB_PATH = Path(_get("TASK_DB_PATH", str(BASE_DIR / "data" / "tasks.db"))).resolve()
REPORT_TASK_DB_PATH = Path(
    _get("REPORT_TASK_DB_PATH", str(BASE_DIR / "data" / "report_tasks.db"))
).resolve()
REPORT_OUTPUT_DIR = Path(
    _get("REPORT_OUTPUT_DIR", str(BASE_DIR / "data" / "reports"))
).resolve()
WORD_REPORT_TEMPLATE_PATH = Path(
    _get(
        "WORD_REPORT_TEMPLATE_PATH",
        str(BASE_DIR / "assets" / "word_templates" / "cid629-joint-simulation-v1.5.docx"),
    )
).resolve()
MAX_CONCURRENT_REPORT_TASKS = max(1, int(_get("MAX_CONCURRENT_REPORT_TASKS", "1")))
MAX_QUEUED_REPORT_TASKS = max(1, int(_get("MAX_QUEUED_REPORT_TASKS", "10")))
MAX_REPORT_SOURCE_CHARS = max(10000, int(_get("MAX_REPORT_SOURCE_CHARS", "80000")))
MAX_AGENT_STEPS = max(10, int(_get("MAX_AGENT_STEPS", "80")))

# ---------- 自测模式 ----------
MOCK_LLM = _get("MOCK_LLM", "0") == "1"
MOCK_EXAMPLE = _get("MOCK_EXAMPLE", "ppt169_building_effective_agents")
