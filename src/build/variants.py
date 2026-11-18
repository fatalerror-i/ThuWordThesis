# -*- coding: utf-8 -*-
"""双变体注册表：把 --variant graduate/undergraduate 解析成一组路径与开关。"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[2]
REGISTRY = ROOT / "configs" / "variants.yaml"


class VariantError(RuntimeError):
    """变体名缺失或注册表有误。"""


@dataclass(frozen=True)
class Variant:
    """一个变体的全部构建路径与开关；路径均为绝对 Path。"""
    name: str
    label: str
    rules: Path
    implementation_map: Path
    degree_options: Path
    title: str
    docx: Path
    dotm: Path
    pdf: Path
    content_package: str
    ribbon: Path
    this_document: Path
    decorate_header: bool


def _abs(value: str | Path) -> Path:
    """把相对路径解析为仓库根下的绝对路径。"""
    p = Path(value)
    return p if p.is_absolute() else ROOT / p


def choices() -> tuple[str, ...]:
    """返回注册表中按声明顺序排列的变体名，供 argparse choices 使用。"""
    data = yaml.safe_load(REGISTRY.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not data:
        raise VariantError(f"变体注册表为空或格式错误：{REGISTRY}")
    return tuple(str(k) for k in data)


def load_registry() -> dict[str, Variant]:
    """加载并校验全部变体，返回 name -> Variant。"""
    data = yaml.safe_load(REGISTRY.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not data:
        raise VariantError(f"变体注册表为空或格式错误：{REGISTRY}")
    out: dict[str, Variant] = {}
    required = (
        "label", "rules", "implementationMap", "degreeOptions", "title",
        "docx", "dotm", "pdf", "contentPackage", "ribbon", "thisDocument",
        "decorateHeader",
    )
    for name, raw in data.items():
        name = str(name)
        if not isinstance(raw, dict):
            raise VariantError(f"变体 {name} 的配置不是映射：{REGISTRY}")
        missing = [k for k in required if k not in raw]
        if missing:
            raise VariantError(f"变体 {name} 缺少字段：{', '.join(missing)}")
        out[name] = Variant(
            name=name,
            label=str(raw["label"]),
            rules=_abs(raw["rules"]),
            implementation_map=_abs(raw["implementationMap"]),
            degree_options=_abs(raw["degreeOptions"]),
            title=str(raw["title"]),
            docx=_abs(raw["docx"]),
            dotm=_abs(raw["dotm"]),
            pdf=_abs(raw["pdf"]),
            content_package=str(raw["contentPackage"]),
            ribbon=_abs(raw["ribbon"]),
            this_document=_abs(raw["thisDocument"]),
            decorate_header=bool(raw["decorateHeader"]),
        )
    return out


def resolve(name: str) -> Variant:
    """按名字取变体配置。"""
    registry = load_registry()
    if name not in registry:
        raise VariantError(
            f"未知变体 {name!r}；可选：{', '.join(registry)}"
        )
    return registry[name]
