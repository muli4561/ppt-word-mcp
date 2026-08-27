# coding=utf-8
"""模板格式提取、用户可编辑格式确认，以及短效确认凭证。"""
from __future__ import annotations

import hashlib
import time
import zipfile
from pathlib import Path
from typing import Callable, Optional
from xml.etree import ElementTree as ET

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from pydantic import BaseModel

from .config import WORD_REPORT_TEMPLATE_PATH
from .report_documents import validate_docx_package
from .report_models import WordFormatOptions, WordFormatOverrides
from .signed_tokens import InvalidSignedToken, sign_claims, verify_claims


WORD_FORMAT_PROFILE_ID = "cid629-joint-simulation-v1.5"
WORD_FORMAT_PROFILE_NAME = "CID629 联合仿真测试报告 v1.5"
WORD_FORMAT_CONFIRM_TTL_SECONDS = 15 * 60


class WordFormatConfirmation(BaseModel):
    profile_id: str
    profile_name: str
    template_source: str
    template_upload_id: str = ""
    reference_filename: str
    format: WordFormatOptions
    rules: list[str]
    confirmation_text: str
    confirmation_token: str
    expires_at: int
    next_tool: str = "generate_word_report"


class VerifiedWordFormat(BaseModel):
    profile_id: str
    template_upload_id: str = ""
    reference_filename: str
    template_sha256: str
    format: WordFormatOptions


def _built_in_template_path() -> Path:
    path = Path(WORD_REPORT_TEMPLATE_PATH)
    if not path.is_file():
        raise RuntimeError(f"内置 Word 格式模板不存在: {path}")
    validate_docx_package(path)
    return path


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _effective_font(style, fallback: str) -> str:
    current = style
    visited = set()
    while current is not None and current.style_id not in visited:
        visited.add(current.style_id)
        rpr = current._element.rPr
        fonts = rpr.rFonts if rpr is not None else None
        if fonts is not None:
            value = fonts.get(qn("w:eastAsia")) or fonts.get(qn("w:ascii"))
            if value:
                return value
        if current.font.name:
            return current.font.name
        current = current.base_style
    return fallback


def _effective_size(style, fallback: float) -> float:
    current = style
    visited = set()
    while current is not None and current.style_id not in visited:
        visited.add(current.style_id)
        if current.font.size is not None:
            return round(float(current.font.size.pt), 2)
        current = current.base_style
    return fallback


def _detect_numbering_style(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/numbering.xml"))
        formats = [
            node.get(qn("w:val"), "")
            for node in root.findall(".//" + qn("w:numFmt"))
        ]
    except (KeyError, ET.ParseError, zipfile.BadZipFile):
        return "decimal"
    return "chinese" if any("chinese" in value.lower() for value in formats) else "decimal"


def extract_word_format(path: Path) -> WordFormatOptions:
    """读取模板的正文和 1–3 级标题格式，缺失项回退到专业报告默认值。"""
    validate_docx_package(path)
    document = Document(path)
    defaults = WordFormatOptions()
    styles = document.styles

    def style(name: str):
        try:
            return styles[name]
        except KeyError:
            return styles["Normal"]

    normal = style("Normal")
    h1, h2, h3 = style("Heading 1"), style("Heading 2"), style("Heading 3")
    paragraph_format = normal.paragraph_format
    spacing = paragraph_format.line_spacing
    if not isinstance(spacing, (int, float)):
        spacing = defaults.line_spacing
    indent = paragraph_format.first_line_indent
    indent_chars = round(indent.cm / 0.3704166667, 2) if indent is not None and indent.cm > 0 else 0.0
    return WordFormatOptions(
        body_font=_effective_font(normal, defaults.body_font),
        body_size_pt=_effective_size(normal, defaults.body_size_pt),
        line_spacing=float(spacing),
        first_line_indent_chars=indent_chars,
        heading1_font=_effective_font(h1, defaults.heading1_font),
        heading1_size_pt=_effective_size(h1, defaults.heading1_size_pt),
        heading2_font=_effective_font(h2, defaults.heading2_font),
        heading2_size_pt=_effective_size(h2, defaults.heading2_size_pt),
        heading3_font=_effective_font(h3, defaults.heading3_font),
        heading3_size_pt=_effective_size(h3, defaults.heading3_size_pt),
        numbering_style=_detect_numbering_style(path),
    )


def _apply_overrides(base: WordFormatOptions, overrides: Optional[WordFormatOverrides]) -> WordFormatOptions:
    if overrides is None:
        return base
    values = overrides.model_dump(exclude_none=True)
    return base.model_copy(update=values)


def _rules(options: WordFormatOptions, template_name: str) -> list[str]:
    numbering = "一. / 一.一." if options.numbering_style == "chinese" else "1. / 1.1."
    return [
        f"版式来源：{template_name}（保留模板页边距、页眉和页脚）",
        f"一级标题：{options.heading1_font} {options.heading1_size_pt:g}pt",
        f"二级标题：{options.heading2_font} {options.heading2_size_pt:g}pt；三级标题：{options.heading3_font} {options.heading3_size_pt:g}pt",
        f"正文：{options.body_font} {options.body_size_pt:g}pt、{options.line_spacing:g} 倍行距、首行缩进 {options.first_line_indent_chars:g} 字符",
        f"标题自动编号：{numbering}",
        "自动生成 1–3 级目录，并在打开 Word 时更新目录域",
    ]


def issue_word_format_confirmation(
    template_path: Optional[Path] = None,
    template_upload_id: str = "",
    overrides: Optional[WordFormatOverrides] = None,
) -> WordFormatConfirmation:
    path = Path(template_path) if template_path is not None else _built_in_template_path()
    validate_docx_package(path)
    digest = _sha256(path)
    custom = bool(template_upload_id.strip())
    profile_id = f"custom-template-{digest[:12]}" if custom else WORD_FORMAT_PROFILE_ID
    profile_name = f"用户模板：{path.name}" if custom else WORD_FORMAT_PROFILE_NAME
    options = _apply_overrides(extract_word_format(path), overrides)
    rules = _rules(options, path.name)
    token = sign_claims(
        "word-format-confirmation",
        {
            "profile_id": profile_id,
            "template_upload_id": template_upload_id.strip(),
            "reference_filename": path.name,
            "template_sha256": digest,
            "format": options.model_dump(mode="json"),
        },
        WORD_FORMAT_CONFIRM_TTL_SECONDS,
    )
    expires_at = int(time.time()) + WORD_FORMAT_CONFIRM_TTL_SECONDS
    lines = "\n".join(f"{index}. {rule}" for index, rule in enumerate(rules, 1))
    return WordFormatConfirmation(
        profile_id=profile_id,
        profile_name=profile_name,
        template_source="uploaded" if custom else "built-in",
        template_upload_id=template_upload_id.strip(),
        reference_filename=path.name,
        format=options,
        rules=rules,
        confirmation_text=f"即将按以下 Word 格式生成报告：\n{lines}\n\n是否确认？",
        confirmation_token=token,
        expires_at=expires_at,
    )


def verify_word_format_confirmation(
    token: str,
    template_resolver: Optional[Callable[[str], Path]] = None,
) -> VerifiedWordFormat:
    value = token.strip()
    if not value:
        raise ValueError("生成 Word 报告前必须先确认模板和格式")
    try:
        claims = verify_claims(value, "word-format-confirmation")
        options = WordFormatOptions.model_validate(claims.get("format"))
    except (InvalidSignedToken, TypeError, ValueError) as exc:
        raise ValueError(f"Word 格式确认已失效：{exc}") from exc

    upload_id = str(claims.get("template_upload_id", "")).strip()
    if upload_id:
        if template_resolver is None:
            raise ValueError("确认凭证使用了上传模板，但当前无法读取该模板")
        path = template_resolver(upload_id)
    else:
        path = _built_in_template_path()
    if _sha256(path) != claims.get("template_sha256"):
        raise ValueError("Word 模板已更新或已被替换，请重新确认")
    return VerifiedWordFormat(
        profile_id=str(claims.get("profile_id", "")),
        template_upload_id=upload_id,
        reference_filename=str(claims.get("reference_filename", path.name)),
        template_sha256=str(claims.get("template_sha256", "")),
        format=options,
    )
