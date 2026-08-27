# coding=utf-8
"""DeepSeek 工具循环（Quick 模式）：LLM 客户端 + 提示词 + 工具集 + Agent/Mock。"""
import json
import logging
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Callable, Dict, List, Optional

from openai import OpenAI

from .config import (
    LLM_API_KEY,
    LLM_BASE_URL,
    LLM_MODEL,
    LLM_TEMPERATURE,
    MAX_GENERATED_IMAGES,
    MAX_AGENT_STEPS,
    MOCK_EXAMPLE,
    VISION_API_KEY,
    VISION_BACKEND,
    VISION_BASE_URL,
    VISION_IMAGE_SIZE,
    VISION_MODEL,
)
from .pptmaster import TaskWorkspace, export_pptx, generate_image as pptmaster_generate_image, quality_check

logger = logging.getLogger("ppt_word_gen.agent")

MAX_SVG_BYTES = 400 * 1024
Progress = Callable[[str, int, str], None]


# ================================================================ LLM 客户端

class LLMClient:
    def __init__(
        self,
        base_url: str = LLM_BASE_URL,
        api_key: str = LLM_API_KEY,
        model: str = LLM_MODEL,
        temperature: float = LLM_TEMPERATURE,
    ):
        if not api_key:
            raise ValueError("LLM_API_KEY 未配置（可在 .env 中设置，或开启 MOCK_LLM=1 自测）")
        self.client = OpenAI(base_url=base_url, api_key=api_key, timeout=600, max_retries=2)
        self.model = model
        self.temperature = temperature

    def chat(
        self,
        messages: List[Dict],
        tools: Optional[List[Dict]] = None,
        max_tokens: Optional[int] = None,
    ) -> object:
        kwargs: Dict = {"model": self.model, "messages": messages, "temperature": self.temperature}
        if tools:
            kwargs["tools"] = tools
        if max_tokens is not None:
            kwargs["max_tokens"] = max_tokens
        resp = self.client.chat.completions.create(**kwargs)
        return resp.choices[0].message


# ================================================================ 提示词

SYSTEM_PROMPT = """你是 PPT Master 的 Quick 模式执行引擎：一位资深演示设计师 + 原生 SVG 幻灯片作者。
你的任务：根据用户提供的主题或资料，直接创作一份视觉专业、信息清晰的原生 PPT，
经过质量检查后导出为可编辑的 .pptx。全程不需要向用户提问，所有未明确指定的选择由你直接决定。

## 硬性规则
1. 画布：所有页面使用完全一致的画布：<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720" data-pptx-page-role="..." width="1280" height="720">。
   data-pptx-page-role 取值：封面 cover / 目录 toc / 章节 section / 内容 content / 结尾 ending。
2. 页面扁平化：每页是独立、完整的 SVG 文件。禁止 <script>、<style>、网络资源引用；
   禁止把页面内容拆到多个文件或依赖其他页。不要写设计文档/计划文件，不要写 design_spec/spec_lock。
3. 允许的元素：rect、circle、ellipse、line、polyline、polygon、path、g、text、tspan、image、defs（仅限线性/径向渐变与简单 pattern）。
   不要使用 filter、feGaussianBlur、阴影、动画等复杂效果；不要使用 SmartArt 类结构。
   <image> 只能引用项目 images/ 目录中已存在的文件，不要引用网络 URL；默认尽量不依赖图片，用形状与排版表达。
4. 字体：使用 PowerPoint 常见稳定字体，如 Arial、Arial Black、Georgia、Microsoft YaHei（微软雅黑）、Source Han Sans CN（思源黑体）。
   必须为每个 text 显式声明 font-size（px，基于 1280 宽画布）：封面标题 56-72、页标题 36-48、正文 18-24、注释 14-16。
   中文内容使用中文字体，西文/数字用西文字体，可配对设置（如 font-family="Arial, Microsoft YaHei"）。
5. 配色：整份 deck 使用统一色板（3-5 色），确保正文与背景对比度足够（深字浅底或浅字深底）。背景可为纯色或渐变。
6. 排版：内容页四周留 60-80px 安全边距；善用卡片、分区、对齐网格与留白；标题层级清晰；
   正文行高 1.4-1.6（可用 tspan 分行）；避免文字溢出——估算文本宽度：中文字符宽度≈font-size，西文≈0.55×font-size。
7. 图表/数据：简单柱状图、饼图、折线图可用 rect/path 手工绘制并标注数据；表格用 rect+text 绘制；
   不要用图片代替数据图表。
8. 页数：严格遵守用户要求的页数。每页信息量适中，一页一个核心主题。

## 工作流程（必须按顺序执行）
1. 阅读资料：若提示词说明有源文档，用 read_file 读取（一般在 sources/ 或项目根目录的 .md），提炼要点；否则直接规划主题。
2. 规划：确定每页的核心信息与版式（封面/目录/章节/内容/结尾），用文本简要说明，不落盘。
3. 逐页创作：用 write_svg_page 一次写 1~2 页（page_index 从 1 开始），直到全部页写完。
4. 质量检查：调用 run_quality_check。若有 blocking 错误，仔细阅读错误清单，修复对应页面后重跑（最多 3 轮）；warning 可忽略。
5. 导出：质检通过后调用 export_pptx。
6. 完成：调用 finish，附上给用户的简要说明（页数、主题、风格）。"""

VISION_PROMPT = """

## 视觉模型工具
本任务已启用视觉模型。仅在照片、主视觉或复杂插画能明显增强表达时调用 generate_image，不要每页都生成图片。
工具返回项目内 images/ 文件路径；SVG 页面位于 svg_output/，因此应使用 ../images/文件名 作为 <image href="...">。
生成图是位图素材，标题、正文、数据和图表仍应使用 SVG 原生元素，以保证清晰度和可编辑性。"""


def build_system_prompt(vision_enabled: bool) -> str:
    return SYSTEM_PROMPT + (VISION_PROMPT if vision_enabled else "")


def build_user_prompt(topic: str, page_count: int, style: str, fmt: str, source_hint: str = "") -> str:
    lines = [f"请生成一份 {page_count} 页的 PPT（画布格式: {fmt}）。"]
    if style:
        lines.append(f"风格偏好: {style}。")
    if topic:
        lines.append(f"主题/内容：\n{topic[:8000]}")
    if source_hint:
        lines.append(f"已转换的源资料：{source_hint}（用 read_file 工具读取后使用）。")
    lines.append("开始创作，直接调用工具执行，不要提问。")
    return "\n".join(lines)


# ================================================================ 工具集

def tool_schemas(vision_enabled: bool = False) -> List[Dict]:
    schemas = [
        {
            "type": "function",
            "function": {
                "name": "write_svg_page",
                "description": "把一页完整的 SVG 幻灯片写入 svg_output/。每页调用一次，page_index 从 1 开始连续递增。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "page_index": {"type": "integer", "description": "1 起始的页码"},
                        "svg_xml": {"type": "string", "description": "该页完整的 SVG 文档字符串"},
                    },
                    "required": ["page_index", "svg_xml"],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "read_file",
                "description": "读取工程内的文本文件（如转换后的源文档 .md、已写的 SVG 页面）。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "相对工程根目录的路径，如 sources/xx.md 或 svg_output/01_page.svg"},
                    },
                    "required": ["path"],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "list_dir",
                "description": "列出工程内某个目录下的文件。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "path": {"type": "string", "description": "相对工程根目录的目录路径（默认空 = 根目录）"},
                    },
                    "required": [],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "run_quality_check",
                "description": "对 svg_output/ 全部页面运行最终质量检查，返回 blocking 错误清单与汇总。",
                "parameters": {"type": "object", "properties": {}, "required": []},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "export_pptx",
                "description": "在质量检查通过后，将 svg_output/ 导出为原生可编辑的 .pptx。返回导出文件名。",
                "parameters": {"type": "object", "properties": {}, "required": []},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "finish",
                "description": "完成整个生成任务，附上给用户的总结。调用后任务结束。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "message": {"type": "string", "description": "给用户的总结（页数、主题、风格、注意事项）"},
                    },
                    "required": ["message"],
                },
            },
        },
    ]
    if vision_enabled:
        schemas.insert(-1, {
            "type": "function",
            "function": {
                "name": "generate_image",
                "description": "调用用户配置的视觉模型生成一张项目素材图，返回 images/ 下的相对路径。只在确有视觉价值时使用。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "prompt": {"type": "string", "description": "具体、可直接用于图像生成的提示词；不要要求图片内生成长文字"},
                        "filename": {"type": "string", "description": "简短英文文件名，不含目录和扩展名"},
                        "aspect_ratio": {
                            "type": "string",
                            "enum": ["16:9", "4:3", "3:2", "1:1", "9:16"],
                            "description": "图片宽高比",
                        },
                    },
                    "required": ["prompt", "filename", "aspect_ratio"],
                },
            },
        })
    return schemas


class ToolContext:
    """一次生成任务的上下文：工作区 + 任务进度回调。"""

    def __init__(self, workspace: TaskWorkspace, request, progress: Progress):
        self.ws = workspace
        self.req = request
        self.progress = progress
        self.project_dir: Path = workspace.project_dir
        self.finished = False
        self.finish_message = ""
        self.pptx_path: Optional[Path] = None
        self.last_report: Dict = {}
        self._written_pages: set = set()
        self._generated_images = 0

        custom_vision_key = request.vision_api_key.get_secret_value()
        self.vision_enabled = bool(request.vision_enabled)
        self.vision_backend = request.vision_backend or VISION_BACKEND
        self.vision_base_url = request.vision_base_url or VISION_BASE_URL
        self.vision_api_key = custom_vision_key or VISION_API_KEY
        self.vision_model = request.vision_model or VISION_MODEL
        self.vision_image_size = request.vision_image_size or VISION_IMAGE_SIZE

    # ---------------- 工具实现 ----------------

    def write_svg_page(self, page_index: int, svg_xml: str) -> str:
        if not isinstance(page_index, int) or page_index < 1:
            return "错误：page_index 必须是 >=1 的整数"
        if not isinstance(svg_xml, str) or len(svg_xml.encode("utf-8")) > MAX_SVG_BYTES:
            return f"错误：SVG 过大（>{MAX_SVG_BYTES // 1024}KB）或非文本"
        try:
            root = ET.fromstring(svg_xml)
        except ET.ParseError as exc:
            return f"错误：SVG 不是合法 XML：{exc}"
        if root.tag.split('}')[-1] != "svg":
            return "错误：根元素必须是 <svg>"
        svg_dir = self.ws.svg_dir
        svg_dir.mkdir(parents=True, exist_ok=True)
        path = svg_dir / f"{page_index:02d}_page.svg"
        path.write_text(svg_xml, encoding="utf-8")
        self._written_pages.add(page_index)
        return f"OK：已写入第 {page_index} 页 ({path.name})"

    def read_file(self, path: str) -> str:
        full = self._resolve(path)
        if full is None:
            return f"错误：路径越界或不存在: {path}"
        if not full.is_file():
            return f"错误：不是文件: {path}"
        try:
            return full.read_text(encoding="utf-8", errors="replace")[:50000]
        except OSError as exc:
            return f"错误：读取失败 {exc}"

    def list_dir(self, path: str = "") -> str:
        full = self._resolve(path)
        if full is None:
            return f"错误：路径越界: {path}"
        if not full.is_dir():
            return f"错误：不是目录: {path}"
        entries = sorted(full.iterdir(), key=lambda p: (p.is_file(), p.name.lower()))
        return "\n".join(p.name + ("/" if p.is_dir() else "") for p in entries[:200]) or "(空目录)"

    def generate_image(self, prompt: str, filename: str = "visual", aspect_ratio: str = "16:9") -> str:
        if not self.vision_enabled:
            return "错误：本任务未启用视觉模型"
        if not self.vision_api_key:
            return "错误：视觉模型 API Key 未配置"
        if self._generated_images >= MAX_GENERATED_IMAGES:
            return f"错误：本任务最多生成 {MAX_GENERATED_IMAGES} 张图片"
        if aspect_ratio not in {"16:9", "4:3", "3:2", "1:1", "9:16"}:
            return "错误：不支持的图片宽高比"
        self.progress("视觉生成", min(82, 35 + self._generated_images * 4), "生成演示素材图…")
        path = pptmaster_generate_image(
            output_dir=self.project_dir / "images",
            prompt=prompt,
            backend=self.vision_backend,
            api_key=self.vision_api_key,
            base_url=self.vision_base_url,
            model=self.vision_model,
            image_size=self.vision_image_size,
            aspect_ratio=aspect_ratio,
            filename=filename,
        )
        self._generated_images += 1
        return f"OK：图片已生成到 images/{path.name}；在 svg_output 页面中请用 ../images/{path.name} 引用"

    def run_quality_check(self) -> str:
        self.progress("质量检查", 85, "运行 SVG 最终质量检查…")
        report = quality_check(self.project_dir)
        self.last_report = report
        summary = report.get("summary", {})
        errors = summary.get("errors", 0)
        warnings = summary.get("warnings", 0)
        total = summary.get("total", 0)
        blocking = report.get("categories", {}).get("blocking", {})
        if isinstance(blocking, dict):
            blocking_issues = blocking.get("issues", []) or []
            blocking_count = blocking.get("count", len(blocking_issues))
        else:
            blocking_issues = blocking or []
            blocking_count = len(blocking_issues)
        lines = [f"检查完成：共 {total} 个文件，errors={errors}，warnings={warnings}"]
        if blocking_count:
            lines.append(f"blocking 问题 {blocking_count} 条：")
            for item in blocking_issues[:30]:
                if isinstance(item, dict):
                    lines.append(f"- [{item.get('code')}] {item.get('file')}: {str(item.get('message'))[:300]}")
                else:
                    lines.append(f"- {str(item)[:300]}")
        else:
            lines.append("无 blocking 问题，可以导出。")
        return "\n".join(lines)

    def export_pptx(self) -> str:
        if not self.last_report:
            return "错误：请先运行 run_quality_check 且无 blocking 错误，再导出。"
        if self.last_report.get("summary", {}).get("errors", 0) > 0:
            return "错误：质量检查存在 blocking 错误，请先修复。"
        self.progress("导出", 95, "导出 PPTX…")
        try:
            out = export_pptx(self.project_dir)
        except Exception as exc:
            return f"错误：导出失败 {exc}"
        self.pptx_path = out
        return f"OK：已导出 {out.name}"

    def finish(self, message: str) -> str:
        self.finished = True
        self.finish_message = message
        return f"OK：任务完成。{message}"

    # ---------------- 内部 ----------------

    def _resolve(self, rel: str) -> Optional[Path]:
        rel = (rel or "").strip()
        if not rel or rel == ".":
            return self.project_dir
        full = (self.project_dir / rel).resolve()
        try:
            full.relative_to(self.project_dir.resolve())
        except ValueError:
            return None
        return full


def dispatch(ctx: ToolContext, name: str, args) -> str:
    fn: Optional[Callable] = getattr(ctx, name, None)
    if fn is None:
        return f"错误：未知工具 {name}"
    try:
        if isinstance(args, str):
            args = json.loads(args or "{}")
        if not isinstance(args, dict):
            args = {}
        result = fn(**args)
        if isinstance(result, str):
            return result
        return json.dumps(result, ensure_ascii=False, default=str)
    except TypeError as exc:
        return f"错误：工具 {name} 参数错误：{exc}"
    except Exception as exc:  # noqa: BLE001
        return f"错误：工具 {name} 执行异常：{exc}"


# ================================================================ 生成循环

class AgentLoop:
    """真实 LLM function-calling 循环（Quick 模式）。"""

    def __init__(self, ctx: ToolContext, progress: Progress):
        self.ctx = ctx
        self.progress = progress
        req = ctx.req
        self.client = LLMClient(
            base_url=req.language_base_url or LLM_BASE_URL,
            api_key=req.language_api_key.get_secret_value() or LLM_API_KEY,
            model=req.language_model or LLM_MODEL,
            temperature=req.language_temperature if req.language_temperature is not None else LLM_TEMPERATURE,
        )

    def run(self) -> None:
        self.progress("AI 创作", 20, "规划并逐页创作 SVG…")
        source_hint = ""
        sources = self.ctx.ws.project_dir / "sources"
        if sources.is_dir():
            mds = sorted(sources.rglob("*.md"))
            if mds:
                source_hint = mds[0].relative_to(self.ctx.project_dir).as_posix()

        messages = [
            {"role": "system", "content": build_system_prompt(self.ctx.vision_enabled)},
            {
                "role": "user",
                "content": build_user_prompt(
                    topic=self.ctx.req.topic,
                    page_count=self.ctx.req.page_count,
                    style=self.ctx.req.style,
                    fmt=self.ctx.req.format,
                    source_hint=source_hint,
                ),
            },
        ]
        tools = tool_schemas(self.ctx.vision_enabled)

        for step in range(1, MAX_AGENT_STEPS + 1):
            if self.ctx.finished:
                break
            self.progress("AI 创作", min(80, 20 + step), f"LLM 步骤 {step}/{MAX_AGENT_STEPS}…")
            msg = self.client.chat(messages, tools=tools)
            tool_calls = getattr(msg, "tool_calls", None) or []

            if tool_calls:
                assistant_entry = {"role": "assistant", "content": getattr(msg, "content", None) or ""}
                calls = []
                for tc in tool_calls:
                    try:
                        json.loads(tc.function.arguments or "{}")
                    except json.JSONDecodeError:
                        pass
                    calls.append({
                        "id": tc.id, "type": "function",
                        "function": {"name": tc.function.name, "arguments": tc.function.arguments or "{}"},
                    })
                assistant_entry["tool_calls"] = calls
                messages.append(assistant_entry)

                stop = False
                for tc in tool_calls:
                    result = dispatch(self.ctx, tc.function.name, tc.function.arguments)
                    messages.append({"role": "tool", "tool_call_id": tc.id, "content": result})
                    logger.info("[%s] tool %s -> %s", step, tc.function.name, result[:200].replace("\n", " "))
                    if tc.function.name == "finish":
                        stop = True
                if stop:
                    break
            else:
                messages.append({"role": "assistant", "content": getattr(msg, "content", None) or ""})

        # 兜底：循环结束仍未导出则尝试导出
        if self.ctx.pptx_path is None and self.ctx.last_report:
            if self.ctx.last_report.get("summary", {}).get("errors", 0) == 0:
                self.progress("导出", 95, "兜底导出…")
                self.ctx.export_pptx()

        if self.ctx.pptx_path is None:
            raise RuntimeError("未能完成导出：LLM 未通过质检或未调用导出工具")


class MockAgent:
    """自测模式：不使用 LLM，复制官方示例 deck 走完整流水线。"""

    def __init__(self, ctx: ToolContext, progress: Progress):
        self.ctx = ctx
        self.progress = progress

    def run(self) -> None:
        self.progress("AI 创作", 25, f"Mock：复制示例工程 {MOCK_EXAMPLE}…")
        pages = self.ctx.ws.copy_example_deck(MOCK_EXAMPLE, max_pages=self.ctx.req.page_count)
        self.ctx.req.page_count = pages
        report_text = self.ctx.run_quality_check()
        logger.info("mock quality check:\n%s", report_text)
        self.progress("导出", 90, "导出 PPTX…")
        result = self.ctx.export_pptx()
        logger.info("mock export: %s", result)
        if self.ctx.pptx_path is None:
            raise RuntimeError(f"Mock 导出失败: {result}")
        self.ctx.finish_message = (
            f"（Mock 模式）已按示例工程生成 {pages} 页示例 PPT，仅用于流水线自测。"
        )
        self.ctx.finished = True
