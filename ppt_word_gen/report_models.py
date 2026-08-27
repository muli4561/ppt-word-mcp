# coding=utf-8
"""Word 报告请求、结构化中间产物与任务响应模型。"""
from typing import List, Literal, Optional
from urllib.parse import urlsplit

from pydantic import BaseModel, Field, SecretStr, field_validator, model_validator


REPORT_TYPES = {
    "delivery": "Agent 交付报告",
    "validation": "联合仿真测试报告",
    "manual": "使用说明书",
    "technical": "技术分析报告",
}


class WordFormatOptions(BaseModel):
    body_font: str = Field("仿宋", min_length=1, max_length=80)
    body_size_pt: float = Field(14.0, ge=8, le=72)
    line_spacing: float = Field(1.5, ge=0.8, le=3)
    first_line_indent_chars: float = Field(2.0, ge=0, le=6)
    heading1_font: str = Field("黑体", min_length=1, max_length=80)
    heading1_size_pt: float = Field(22.0, ge=8, le=72)
    heading2_font: str = Field("黑体", min_length=1, max_length=80)
    heading2_size_pt: float = Field(18.0, ge=8, le=72)
    heading3_font: str = Field("黑体", min_length=1, max_length=80)
    heading3_size_pt: float = Field(16.0, ge=8, le=72)
    numbering_style: Literal["decimal", "chinese"] = "decimal"

    @field_validator("body_font", "heading1_font", "heading2_font", "heading3_font")
    @classmethod
    def normalize_font_name(cls, value: str) -> str:
        return value.strip()


class WordFormatOverrides(BaseModel):
    body_font: Optional[str] = Field(None, max_length=80)
    body_size_pt: Optional[float] = Field(None, ge=8, le=72)
    line_spacing: Optional[float] = Field(None, ge=0.8, le=3)
    first_line_indent_chars: Optional[float] = Field(None, ge=0, le=6)
    heading1_font: Optional[str] = Field(None, max_length=80)
    heading1_size_pt: Optional[float] = Field(None, ge=8, le=72)
    heading2_font: Optional[str] = Field(None, max_length=80)
    heading2_size_pt: Optional[float] = Field(None, ge=8, le=72)
    heading3_font: Optional[str] = Field(None, max_length=80)
    heading3_size_pt: Optional[float] = Field(None, ge=8, le=72)
    numbering_style: Optional[Literal["decimal", "chinese"]] = None

    @field_validator("body_font", "heading1_font", "heading2_font", "heading3_font")
    @classmethod
    def normalize_optional_font_name(cls, value: Optional[str]) -> Optional[str]:
        value = (value or "").strip()
        return value or None


class ReportCreate(BaseModel):
    title: str = Field("", max_length=300)
    report_type: Literal["delivery", "validation", "manual", "technical"] = "validation"
    instructions: str = Field("", max_length=12000)
    project_name: str = Field("", max_length=300)
    document_version: str = Field("v1.0", max_length=60)
    author: str = Field("", max_length=100)
    word_format_profile_id: str = Field("cid629-joint-simulation-v1.5", max_length=120)
    word_format: WordFormatOptions = Field(default_factory=WordFormatOptions)

    language_base_url: str = Field("", max_length=500)
    language_api_key: SecretStr = Field(default_factory=lambda: SecretStr(""), repr=False)
    language_model: str = Field("", max_length=200)
    language_temperature: Optional[float] = Field(None, ge=0, le=2)

    @field_validator("language_base_url")
    @classmethod
    def validate_model_url(cls, value: str) -> str:
        value = value.strip().rstrip("/")
        if not value:
            return value
        parsed = urlsplit(value)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise ValueError("模型地址必须是有效的 http/https URL")
        if parsed.username or parsed.password:
            raise ValueError("模型地址中不能包含用户名或密码")
        return value

    @model_validator(mode="after")
    def normalize_text(self):
        self.title = self.title.strip()
        self.instructions = self.instructions.strip()
        self.project_name = self.project_name.strip()
        self.document_version = self.document_version.strip() or "v1.0"
        self.author = self.author.strip()
        return self


class ReportBlock(BaseModel):
    type: Literal["paragraph", "bullets", "table", "image", "page_break"]
    text: str = ""
    items: List[str] = Field(default_factory=list)
    headers: List[str] = Field(default_factory=list)
    rows: List[List[str]] = Field(default_factory=list)
    image_name: str = ""
    caption: str = ""
    evidence_ids: List[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_shape(self):
        if self.type == "paragraph" and not self.text.strip():
            raise ValueError("paragraph block 缺少 text")
        if self.type == "bullets" and not self.items:
            raise ValueError("bullets block 缺少 items")
        if self.type == "table" and (not self.headers or not self.rows):
            raise ValueError("table block 缺少 headers/rows")
        if self.type == "image" and not self.image_name.strip():
            raise ValueError("image block 缺少 image_name")
        return self


class ReportSection(BaseModel):
    heading: str = Field(..., min_length=1, max_length=300)
    level: int = Field(1, ge=1, le=3)
    blocks: List[ReportBlock] = Field(default_factory=list)


class ReportSpec(BaseModel):
    title: str = Field(..., min_length=1, max_length=300)
    subtitle: str = Field("", max_length=300)
    report_type: Literal["delivery", "validation", "manual", "technical"]
    project_name: str = Field("", max_length=300)
    document_version: str = Field("v1.0", max_length=60)
    author: str = Field("", max_length=100)
    report_date: str = Field("", max_length=60)
    executive_summary: str = Field(..., min_length=1, max_length=5000)
    sections: List[ReportSection] = Field(..., min_length=1, max_length=30)
    conclusions: List[str] = Field(default_factory=list, max_length=20)
    risks: List[str] = Field(default_factory=list, max_length=20)

    @model_validator(mode="after")
    def require_heading_hierarchy(self):
        levels = {section.level for section in self.sections}
        if 1 not in levels or 2 not in levels:
            raise ValueError("报告正文必须同时包含一级标题和二级标题")
        return self


class ReportTaskInfo(BaseModel):
    task_id: str
    status: str
    stage: str = ""
    progress: int = Field(0, ge=0, le=100)
    message: str = ""
    error: Optional[str] = None
    document_url: Optional[str] = None
    cancel_requested: bool = False
    created_at: float = 0.0
    updated_at: float = 0.0
