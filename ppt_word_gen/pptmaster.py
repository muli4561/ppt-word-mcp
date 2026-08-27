# coding=utf-8
"""PPT Master 封装：脚本白名单（subprocess + UTF-8）+ 任务工作区。"""
import json
import os
import re
import shutil
import subprocess
import sys
import uuid
from pathlib import Path
from typing import Dict, List, Optional

from .config import EXAMPLES_DIR, PROJECTS_DIR, SCRIPTS_DIR

# 允许调用的脚本白名单（防注入：Agent 只能通过这些入口操作）
ALLOWED_SCRIPTS = {
    "project_manager.py",
    "source_to_md.py",
    "svg_quality_checker.py",
    "svg_to_pptx.py",
    "finalize_svg.py",
    "analyze_images.py",
    "image_search.py",
    "image_gen.py",
    "icon_sync.py",
}

_SAFE_NAME = re.compile(r"[^0-9A-Za-z_.\-]")


class ScriptError(RuntimeError):
    def __init__(self, script: str, returncode: int, stdout: str, stderr: str):
        self.script = script
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr
        tail = (stdout or "").strip().splitlines()[-12:]
        detail = "\n".join(tail) if tail else (stderr or "").strip()
        super().__init__(f"{script} 退出码 {returncode}: {detail[:2000]}")


# ---------------------------------------------------------------- 脚本执行

def run_script(
    script: str,
    args: List[str],
    cwd: Optional[Path] = None,
    timeout: int = 1800,
    env_extra: Optional[Dict[str, str]] = None,
    check: bool = True,
) -> subprocess.CompletedProcess:
    if script not in ALLOWED_SCRIPTS:
        raise ValueError(f"脚本不在白名单内: {script}")
    cmd = [sys.executable, str(SCRIPTS_DIR / script), *[str(a) for a in args]]
    env = os.environ.copy()
    if env_extra:
        env.update(env_extra)
    proc = subprocess.run(
        cmd,
        cwd=str(cwd or PROJECTS_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=env,
    )
    if check and proc.returncode != 0:
        raise ScriptError(script, proc.returncode, proc.stdout, proc.stderr)
    return proc


# ---------------------------------------------------------------- 高层命令

def init_project(project_dir: Path, fmt: str = "ppt169") -> Path:
    """project_manager.py init <name> --format <fmt> --dir <PROJECTS_DIR> --quick-generate
    显式指定 --dir 使工程落在 PROJECTS_DIR 下（init 会追加 _<格式>_<日期> 后缀）。
    返回实际创建的工程目录。"""
    proc = run_script(
        "project_manager.py",
        ["init", project_dir.name, "--format", fmt, "--dir", str(PROJECTS_DIR), "--quick-generate"],
        cwd=PROJECTS_DIR,
        timeout=300,
    )
    m = re.search(r"(?:Project created|Project initialized):\s*(.+)$", proc.stdout, re.MULTILINE)
    if not m:
        raise RuntimeError(f"无法解析 project_manager init 输出:\n{proc.stdout[-1000:]}")
    created = Path(m.group(1).strip()).resolve()
    if not created.is_dir():
        raise RuntimeError(f"init 声明的目录不存在: {created}")
    return created


def import_sources(project_dir: Path, sources: List[Path]) -> None:
    run_script(
        "project_manager.py",
        ["import-sources", str(project_dir), *[str(s) for s in sources]],
        cwd=PROJECTS_DIR,
        timeout=600,
    )


def convert_source(path_or_url: str, cwd: Path) -> str:
    """source_to_md.py <file_or_url> —— 在 cwd 中生成 <name>.md 及转换说明。"""
    proc = run_script("source_to_md.py", [path_or_url], cwd=cwd, timeout=600, check=False)
    if proc.returncode != 0:
        raise ScriptError("source_to_md.py", proc.returncode, proc.stdout, proc.stderr)
    return proc.stdout


def quality_check(project_dir: Path) -> Dict:
    """svg_quality_checker.py <project> --quick-generate --stage final --json
    返回解析后的 JSON 报告；报告同时落盘到 <project>/validation/svg_quality_report.json。"""
    proc = run_script(
        "svg_quality_checker.py",
        [str(project_dir), "--quick-generate", "--stage", "final", "--json"],
        cwd=PROJECTS_DIR,
        timeout=600,
        check=False,
    )
    report_path = project_dir / "validation" / "svg_quality_report.json"
    if report_path.is_file():
        try:
            return json.loads(report_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            pass
    raise ScriptError("svg_quality_checker.py", proc.returncode, proc.stdout, proc.stderr)


def export_pptx(project_dir: Path, with_notes: bool = False, timeout: int = 1800) -> Path:
    """svg_to_pptx.py <project> --quick-generate (--with-notes|--no-notes)；返回导出的 .pptx 路径。"""
    flag = "--with-notes" if with_notes else "--no-notes"
    proc = run_script(
        "svg_to_pptx.py",
        [str(project_dir), "--quick-generate", flag],
        cwd=PROJECTS_DIR,
        timeout=timeout,
        check=False,
    )
    if proc.returncode != 0:
        raise ScriptError("svg_to_pptx.py", proc.returncode, proc.stdout, proc.stderr)
    exports = project_dir / "exports"
    if exports.is_dir():
        candidates = sorted(exports.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if candidates:
            return candidates[0]
    raise RuntimeError(f"导出成功但未找到 .pptx 文件:\n{proc.stdout[-2000:]}")


_VISION_ENV = {
    "openai": ("OPENAI_API_KEY", "OPENAI_BASE_URL", "OPENAI_MODEL"),
    "gemini": ("GEMINI_API_KEY", "GEMINI_BASE_URL", "GEMINI_MODEL"),
    "qwen": ("QWEN_API_KEY", "QWEN_BASE_URL", "QWEN_MODEL"),
    "zhipu": ("ZHIPU_API_KEY", "ZHIPU_BASE_URL", "ZHIPU_MODEL"),
    "volcengine": ("LAS_API_KEY", "VOLCENGINE_BASE_URL", "VOLCENGINE_MODEL"),
}


def generate_image(
    output_dir: Path,
    prompt: str,
    backend: str,
    api_key: str,
    base_url: str = "",
    model: str = "",
    image_size: str = "1K",
    aspect_ratio: str = "16:9",
    filename: str = "visual",
) -> Path:
    """通过 PPT Master 的 image_gen.py 生成单张图片并返回落盘路径。"""
    backend = backend.strip().lower()
    if backend not in _VISION_ENV:
        raise ValueError(f"不支持的视觉模型供应商: {backend}")
    if not api_key:
        raise ValueError("视觉模型 API Key 未配置")
    if not prompt.strip():
        raise ValueError("图片提示词不能为空")

    output_dir.mkdir(parents=True, exist_ok=True)
    safe_stem = _SAFE_NAME.sub("_", Path(filename).stem).strip("._") or "visual"
    # 每次生成使用唯一文件名，避免覆盖同任务内已有素材。
    safe_stem = f"{safe_stem[:48]}_{uuid.uuid4().hex[:6]}"
    key_env, base_env, model_env = _VISION_ENV[backend]
    env_extra = {"IMAGE_BACKEND": backend, key_env: api_key}
    if base_url:
        env_extra[base_env] = base_url
    if model:
        env_extra[model_env] = model

    args = [
        prompt[:6000], "--backend", backend, "--output", str(output_dir),
        "--filename", safe_stem, "--aspect_ratio", aspect_ratio,
        "--image_size", image_size,
    ]
    if model:
        args.extend(["--model", model])
    run_script("image_gen.py", args, cwd=output_dir, timeout=900, env_extra=env_extra)
    candidates = sorted(
        output_dir.glob(f"{safe_stem}.*"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        raise RuntimeError("视觉模型调用成功，但没有找到生成的图片文件")
    return candidates[0]


# ---------------------------------------------------------------- 任务工作区

class TaskWorkspace:
    """每任务独立项目目录 + 暂存目录（用于上传的源文档）。"""

    def __init__(self, task_id: str):
        self.task_id = task_id
        self.slug = f"task_{task_id[:8]}"
        self.project_dir: Path = PROJECTS_DIR / self.slug
        self.staging_dir: Path = PROJECTS_DIR / f"_staging_{task_id[:8]}"
        self._cleanup_done = False

    def prepare_staging(self) -> None:
        self.staging_dir.mkdir(parents=True, exist_ok=True)

    def save_upload(self, data: bytes, filename: str) -> Path:
        self.prepare_staging()
        safe = _SAFE_NAME.sub("_", Path(filename).name) or "source.bin"
        path = self.staging_dir / safe
        path.write_bytes(data)
        return path

    def cleanup_staging(self) -> None:
        if self._cleanup_done:
            return
        shutil.rmtree(self.staging_dir, ignore_errors=True)
        self._cleanup_done = True

    @property
    def svg_dir(self) -> Path:
        return self.project_dir / "svg_output"

    def find_exported_pptx(self) -> Optional[Path]:
        exports = self.project_dir / "exports"
        if not exports.is_dir():
            return None
        candidates = sorted(exports.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
        return candidates[0] if candidates else None

    # ---- Mock：复制官方示例 deck ----
    def copy_example_deck(self, example_name: str, max_pages: int) -> int:
        """把 examples/<name> 的 SVG 页面（及其引用的 images/ 等资源）复制进工程。返回实际页数。"""
        src = EXAMPLES_DIR / example_name
        if not src.is_dir():
            raise FileNotFoundError(f"示例工程不存在: {src}")
        for sub in ("images", "icons", "sources", "assets"):
            s = src / sub
            if s.is_dir():
                shutil.copytree(s, self.project_dir / sub, dirs_exist_ok=True)
        svg_files = sorted(
            (src / "svg_output").glob("*.svg"),
            key=lambda p: [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", p.stem.lower())],
        )
        if not svg_files:
            raise FileNotFoundError(f"示例工程没有 SVG 页面: {src / 'svg_output'}")
        target = self.svg_dir
        target.mkdir(parents=True, exist_ok=True)
        for p in target.glob("*.svg"):
            p.unlink()
        kept = svg_files[:max_pages]
        for i, p in enumerate(kept, start=1):
            shutil.copy2(p, target / f"{i:02d}_{p.name}")
        return len(kept)
