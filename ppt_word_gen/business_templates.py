# coding=utf-8
"""业务模板注册中心：内置规范 + 持久化自定义模板。"""
from __future__ import annotations

import json
import re
import shutil
import threading
import time
import uuid
from pathlib import Path
from typing import Dict, Literal, Optional, Tuple

from pydantic import BaseModel

from .config import BUSINESS_TEMPLATE_DIR
from .report_documents import safe_filename, validate_docx_package


DocumentType = Literal["presentation", "word_report"]
_ID_RE = re.compile(r"^(builtin_[a-z0-9_]+|tpl_[0-9a-f]{12})$")


class BusinessTemplateNotFound(ValueError):
    pass


class BusinessTemplateConflict(ValueError):
    pass


class BusinessTemplate(BaseModel):
    template_id: str
    name: str
    document_type: DocumentType
    description: str = ""
    instructions: str = ""
    report_type: str = ""
    has_file: bool = False
    filename: str = ""
    builtin: bool = False
    created_at: float = 0.0


_BUILTINS: Dict[str, BusinessTemplate] = {
    "builtin_agent_delivery": BusinessTemplate(
        template_id="builtin_agent_delivery",
        name="AI 仿真 Agent 交付报告",
        document_type="word_report",
        description="面向公司项目交付，突出范围、架构、部署、验证、风险和运维。",
        instructions="采用正式交付报告结构，明确交付范围、Agent 架构、工具链、部署说明、测试证据、已知限制、风险与验收结论。",
        report_type="delivery",
        builtin=True,
    ),
    "builtin_simulation_validation": BusinessTemplate(
        template_id="builtin_simulation_validation",
        name="联合仿真测试报告",
        document_type="word_report",
        description="适用于联合仿真场景验证和测试证据归档。",
        instructions="按测试目标、环境、模型与接口、测试场景、步骤、结果证据、偏差分析和结论组织内容，禁止虚构测试数据。",
        report_type="validation",
        builtin=True,
    ),
    "builtin_user_manual": BusinessTemplate(
        template_id="builtin_user_manual",
        name="Agent 使用说明书",
        document_type="word_report",
        description="面向最终用户的安装、配置、操作和故障处理说明。",
        instructions="突出前置条件、安装配置、标准操作流程、输入输出示例、常见错误、限制和维护方式。",
        report_type="manual",
        builtin=True,
    ),
    "builtin_technical_brief": BusinessTemplate(
        template_id="builtin_technical_brief",
        name="公司技术方案汇报",
        document_type="presentation",
        description="适用于 AI 仿真和 Agent 技术方案内部汇报。",
        instructions="科技商务风格；一页一个核心结论；重点展示业务问题、总体架构、Agent 流程、仿真工具链、接口、验证结果、风险和落地计划。",
        builtin=True,
    ),
}


class BusinessTemplateStore:
    def __init__(self, root: Path = BUSINESS_TEMPLATE_DIR):
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self._lock = threading.RLock()

    def _directory(self, template_id: str) -> Path:
        value = (template_id or "").strip().lower()
        if not _ID_RE.fullmatch(value) or value.startswith("builtin_"):
            raise BusinessTemplateNotFound("业务模板不存在")
        directory = (self.root / value).resolve()
        if directory.parent != self.root:
            raise BusinessTemplateNotFound("业务模板不存在")
        return directory

    def list(self, document_type: str = "all") -> list[BusinessTemplate]:
        templates = list(_BUILTINS.values())
        with self._lock:
            for metadata_path in self.root.glob("tpl_*/metadata.json"):
                try:
                    templates.append(
                        BusinessTemplate.model_validate_json(metadata_path.read_text(encoding="utf-8"))
                    )
                except (OSError, ValueError):
                    continue
        if document_type != "all":
            templates = [item for item in templates if item.document_type == document_type]
        return sorted(templates, key=lambda item: (not item.builtin, item.name, item.template_id))

    def get(self, template_id: str) -> BusinessTemplate:
        value = (template_id or "").strip().lower()
        if value in _BUILTINS:
            return _BUILTINS[value]
        metadata_path = self._directory(value) / "metadata.json"
        try:
            return BusinessTemplate.model_validate_json(metadata_path.read_text(encoding="utf-8"))
        except (FileNotFoundError, ValueError) as exc:
            raise BusinessTemplateNotFound("业务模板不存在") from exc

    def register(
        self,
        *,
        name: str,
        document_type: DocumentType,
        description: str = "",
        instructions: str = "",
        report_type: str = "",
        file_data: Optional[bytes] = None,
        filename: str = "",
    ) -> BusinessTemplate:
        clean_name = name.strip()
        if not clean_name:
            raise ValueError("模板名称不能为空")
        if len(clean_name) > 120:
            raise ValueError("模板名称不能超过 120 个字符")
        if not instructions.strip() and file_data is None:
            raise ValueError("instructions 与模板文件至少提供一个")
        if document_type == "presentation" and file_data is not None:
            raise ValueError("PPT 业务模板当前注册风格规范，不接收 PPTX 母版文件")
        clean_filename = ""
        if file_data is not None:
            clean_filename = safe_filename(filename, "business-template.docx")
            if Path(clean_filename).suffix.lower() not in {".docx", ".dotx"}:
                raise ValueError("Word 业务模板必须是 .docx 或 .dotx")

        template_id = f"tpl_{uuid.uuid4().hex[:12]}"
        directory = self._directory(template_id)
        directory.mkdir(parents=False, exist_ok=False)
        try:
            if file_data is not None:
                asset = directory / clean_filename
                asset.write_bytes(file_data)
                validate_docx_package(asset)
            model = BusinessTemplate(
                template_id=template_id,
                name=clean_name,
                document_type=document_type,
                description=description.strip()[:1000],
                instructions=instructions.strip()[:12000],
                report_type=report_type.strip(),
                has_file=file_data is not None,
                filename=clean_filename,
                created_at=time.time(),
            )
            (directory / "metadata.json").write_text(
                model.model_dump_json(indent=2), encoding="utf-8"
            )
            return model
        except Exception:
            shutil.rmtree(directory, ignore_errors=True)
            raise

    def file(self, template_id: str) -> Optional[Tuple[bytes, str]]:
        template = self.get(template_id)
        if not template.has_file:
            return None
        path = (self._directory(template.template_id) / template.filename).resolve()
        if path.parent != self._directory(template.template_id) or not path.is_file():
            raise BusinessTemplateNotFound("业务模板文件不存在")
        return path.read_bytes(), template.filename

    def delete(self, template_id: str) -> bool:
        value = (template_id or "").strip().lower()
        if value in _BUILTINS:
            raise BusinessTemplateConflict("内置业务模板不能删除")
        directory = self._directory(value)
        if not directory.is_dir():
            raise BusinessTemplateNotFound("业务模板不存在")
        shutil.rmtree(directory)
        return True


business_template_store = BusinessTemplateStore()
