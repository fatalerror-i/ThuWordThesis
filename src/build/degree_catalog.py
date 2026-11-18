# -*- coding: utf-8 -*-
"""从 configs/degree-options.yaml 生成封面学位选项与 VBA 数据模块。"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

ROOT = Path(__file__).resolve().parents[2]
DEGREE_YAML = ROOT / "configs" / "degree-options.yaml"


def load_degree_config(path: Path = DEGREE_YAML) -> dict[str, Any]:
    """读取学位配置并校验学术/专业学位分组。"""
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("degree-options.yaml 必须是映射")
    for key in ("academic", "professional"):
        if not isinstance(data.get(key), list):
            raise ValueError(f"degree-options.yaml 缺少 {key} 列表")
    for group in ("academic", "professional"):
        for item in data[group]:
            if not any(item.get(k) for k in ("master_en", "doctor_en")):
                raise ValueError(f"{item['name_cn']} 未配置可用学位层次")
    return data


def degree_options(config: dict[str, Any]) -> list[str]:
    """只展开配置中列明的学位层次。"""
    options: list[str] = []
    for group, degree_type in (("academic", "学术学位"),
                               ("professional", "专业学位")):
        for item in config[group]:
            name = item["name_cn"]
            for level, key in (("硕士", "master_en"), ("博士", "doctor_en")):
                if not item.get(key):
                    continue
                options.append(
                    f"{name}{level}|{item[key]}|{level}|{degree_type}|{name}"
                )
    return options


def _category_item(config: dict[str, Any], category: str) -> dict[str, Any] | None:
    """按中文名称查找学位类别配置项。"""
    for group in ("academic", "professional"):
        for item in config[group]:
            if item["name_cn"] == category:
                return item
    return None


def degree_display_options(config: dict[str, Any]) -> list[str]:
    """封面下拉显示串：学术“工学硕士学位”，专业“金融硕士专业学位”。"""
    options: list[str] = []
    for group, suffix in (("academic", "学位"), ("professional", "专业学位")):
        for item in config[group]:
            for level, key in (("硕士", "master_en"), ("博士", "doctor_en")):
                if not item.get(key):
                    continue
                options.append(f"{item['name_cn']}{level}{suffix}")
    return options


def parse_degree_display(display: str) -> tuple[str, str, str]:
    """把封面下拉显示串拆成 (level, degree_type, category)。"""
    text = display.strip()
    if text.endswith("专业学位"):
        degree_type = "专业学位"
        text = text[: -len("专业学位")]
    elif text.endswith("学位"):
        degree_type = ""
        text = text[: -len("学位")]
    else:
        raise ValueError(f"无法识别的学位显示串：{display}")
    if len(text) < 3:
        raise ValueError(f"无法识别的学位显示串：{display}")
    level = text[-2:]
    if level not in ("硕士", "博士"):
        raise ValueError(f"学位层次应为硕士或博士：{display}")
    category = text[:-2]
    if not degree_type:
        degree_type = "学术学位"
    return level, degree_type, category


def discipline_options(config: dict[str, Any], category: str, level: str = "硕士") -> list[str]:
    """某学位类别/门类对应的学科（专业领域）候选，顺序照 YAML。"""
    item = _category_item(config, category)
    if item is None:
        raise ValueError(f"未配置的学位类别：{category}")
    # 专业领域未在指南中列明时允许手填，不把专业学位类别当作领域。
    disciplines = item.get("disciplines", [])
    names = [str(name) for name in disciplines]
    if item in config["academic"]:
        for name, categories in config.get("discipline_awards", {}).items():
            if category in categories and name not in names:
                names.append(name)
    return names


def degree_english(config: dict[str, Any], category: str, level: str = "硕士") -> str:
    """某学位类别/门类在指定层次的英文学位名。"""
    item = _category_item(config, category)
    if item is None:
        raise ValueError(f"未配置的学位类别：{category}")
    key = "doctor_en" if level == "博士" else "master_en"
    if not item.get(key):
        raise ValueError(f"未配置的学位层次：{category}{level}")
    return str(item[key])


def _vba_array(values: list[str]) -> list[str]:
    """把字符串列表渲染成 VBA 数组赋值语句。"""
    lines = []
    for index, value in enumerate(values):
        escaped = value.replace('"', '""')
        lines.append(f'    result({index}) = "{escaped}"')
    return lines


def vba_module(config: dict[str, Any]) -> str:
    """生成 modDegreeOptions.bas；唯一数据源是 degree-options.yaml。"""
    options = degree_options(config)
    lines = [
        "Attribute VB_Name = \"modDegreeOptions\"",
        "' 本模块由 src/build/degree_catalog.py 从 configs/degree-options.yaml 生成，请勿手改。",
        "Option Explicit",
        "",
        "Public Function CoverDegreeOptions() As Variant",
        f"    Dim result({len(options) - 1}) As String",
    ]
    lines.extend(_vba_array(options))
    lines.extend([
        "    CoverDegreeOptions = result",
        "End Function",
        "",
        "Public Function CoverDisciplineOptions(ByVal category As String, ByVal dtype As String) As Variant",
        "    Select Case dtype & \"|\" & category",
    ])
    for group, degree_type in (("academic", "学术学位"),
                               ("professional", "专业学位")):
        for item in config[group]:
            name = item["name_cn"]
            discs = discipline_options(config, name)
            lines.append(f'        Case "{degree_type}|{name}"')
            lines.append("            CoverDisciplineOptions = Array(" +
                         ", ".join(f'"{d}"' for d in discs) + ")")
    lines.extend([
        "        Case Else",
        "            CoverDisciplineOptions = Array(category)",
        "    End Select",
        "End Function",
        "",
        "Public Function CoverIsProfessionalCategory(ByVal category As String) As Boolean",
        "    Select Case category",
    ])
    professional_names = [item["name_cn"] for item in config["professional"]]
    for start in range(0, len(professional_names), 12):
        chunk = professional_names[start:start + 12]
        lines.append("        Case " + ", ".join(f'"{n}"' for n in chunk))
    lines.extend([
        "            CoverIsProfessionalCategory = True",
        "        Case Else",
        "            CoverIsProfessionalCategory = False",
        "    End Select",
        "End Function",
        "",
    ])
    lines.extend([
        "Public Function CoverHasProfessionalField(ByVal category As String) As Boolean",
        "    Select Case category",
    ])
    for item in config["professional"]:
        if item.get("has_field"):
            lines.extend([f'        Case "{item["name_cn"]}"',
                          "            CoverHasProfessionalField = True"])
    lines.extend(["    End Select", "End Function", ""])
    english = [
        "Public Function CoverDegreeEnglish(ByVal category As String, ByVal level As String) As String",
        "    Select Case level & \"|\" & category",
    ]
    for group in ("academic", "professional"):
        for item in config[group]:
            name = item["name_cn"]
            for level in ("硕士", "博士"):
                key = "doctor_en" if level == "博士" else "master_en"
                if not item.get(key):
                    continue
                english.append(f'        Case "{level}|{name}"')
                english.append(f'            CoverDegreeEnglish = "{item[key]}"')
    english.extend([
        "        Case Else",
        "            CoverDegreeEnglish = \"\"",
        "    End Select",
        "End Function",
        "",
    ])
    lines.extend(english)
    return "\n".join(lines)


def write_vba_module(path: Path, config: dict[str, Any] | None = None) -> Path:
    """把学位数据写成 VBA 标准模块文件。"""
    path.write_text(vba_module(config or load_degree_config()), encoding="utf-8")
    return path


def cover_defaults(config: dict[str, Any]) -> dict[str, str]:
    """提取封面的默认学位与学科取值。"""
    defaults = config.get("cover_defaults") or {}
    return {
        "degree_cn": str(defaults.get("degree_cn", "")),
        "discipline": str(defaults.get("discipline", "")),
        "discipline_en": str(defaults.get("discipline_en", "")),
    }
