# -*- coding: utf-8 -*-
"""校验论文格式规则 YAML：结构、注释、索引、单位与引用。

用法：
  python src/validate_rules.py                 # 校验研究生规则（默认）
  python src/validate_rules.py --variant undergraduate
  python src/validate_rules.py --all          # 校验全部注册变体
"""
import argparse
import os
import re
import sys

import yaml

from build.style_presentation import presentation_by_key

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOP_ORDER = ["metadata", "conventions", "pageSetup", "styles", "numbering",
             "captions", "tables", "footnotes", "headersAndFooters", "cover",
             "documentParts"]

REQUIRED_STYLES = [
    "body", "chapterTitle", "chapterTitleUnnumbered", "sectionTitle1",
    "sectionTitle2", "sectionTitle3", "appendixTitle", "caption",
    "figureCaption", "tableCaption", "formula", "referenceEntry",
    "header", "footer", "coverTitleChinese", "coverTitleEnglish",
    "authorizationTitle",
]

UNITS = {"pt", "cm", "mm", "character", "line", "multiple"}
XML_KEYS = {"pPr", "rPr", "jc", "sz", "ind", "numPr", "lvlText", "numFmt",
            "suff", "pgMar", "pgSz", "pgNumType", "tblBorders", "tblStylePr",
            "pBdr", "firstLineChars", "hangingChars", "lineRule", "outlineLvl",
            "inherits", "basedOn", "parent", "extends"}


class UniqueKeyLoader(yaml.SafeLoader):
    """拒绝重复键的 YAML 加载器。"""
    pass


def _construct_mapping(loader, node, deep=False):
    """构造映射，遇到重复键时报错。"""
    mapping = {}
    for key_node, value_node in node.value:
        key = loader.construct_object(key_node, deep=deep)
        if key in mapping:
            raise ValueError("duplicate key: %r at line %d" % (key, key_node.start_mark.line + 1))
        mapping[key] = loader.construct_object(value_node, deep=deep)
    return mapping


UniqueKeyLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG,
                                _construct_mapping)


def validate_one(rules_path, map_path, label):
    """校验一份规则文件，返回（错误、警告）两组列表。"""
    errors = []
    warnings = []

    def err(msg):
        """记录一条校验错误。"""
        errors.append(msg)

    def warn(msg):
        """记录一条校验警告。"""
        warnings.append(msg)

    raw = open(rules_path, encoding="utf-8").read()
    lines = raw.split("\n")

    try:
        data = yaml.load(raw, Loader=UniqueKeyLoader)
    except Exception as exc:
        errors.append("YAML parse error: %s" % exc)
        return errors, warnings

    # 1. 顶层键与顺序
    keys = list(data.keys())
    if keys != TOP_ORDER:
        err("top-level keys/order mismatch:\n  expected %s\n  actual   %s" % (TOP_ORDER, keys))

    for k in TOP_ORDER:
        if k not in data:
            err("missing top-level key: %s" % k)

    if (data.get("pageSetup") or {}).get("documentGrid") != "none":
        err("pageSetup.documentGrid must be none")
    hyphenation = (data.get("pageSetup") or {}).get("hyphenation") or {}
    for key in ("automatic", "hyphenateCaps"):
        if type(hyphenation.get(key)) is not bool:
            err("pageSetup.hyphenation.%s must be boolean" % key)
    limit = hyphenation.get("consecutiveLineLimit")
    if type(limit) is not int or limit < 0:
        err("pageSetup.hyphenation.consecutiveLineLimit must be a nonnegative integer")

    # 2. 小驼峰键名、XML 缩写
    def walk_keys(obj, path=""):
        """遍历规则树检查键名与结构。"""
        if isinstance(obj, dict):
            for k, v in obj.items():
                if not isinstance(k, str):
                    if isinstance(k, int) and path.endswith(".levels"):
                        continue
                    err("non-string key at %s: %r" % (path, k))
                    continue
                if not re.fullmatch(r"[a-z][A-Za-z0-9]*", k):
                    err("key not lowerCamelCase at %s: %r" % (path, k))
                if k in XML_KEYS:
                    err("XML abbreviation key at %s: %r" % (path, k))
                walk_keys(v, path + "." + str(k) if path else str(k))
        elif isinstance(obj, list):
            for i, v in enumerate(obj):
                walk_keys(v, "%s[%d]" % (path, i))

    walk_keys(data)
    for k in data.get("styles", {}):
        if k in XML_KEYS:
            err("XML abbreviation style key: %s" % k)

    # 3. 样式索引一致性
    index_block = []
    in_index = False
    for line in lines:
        if line.startswith("# 样式索引"):
            in_index = True
            continue
        if in_index:
            if line.startswith("styles:"):
                break
            index_block.append(line)
    index_keys = []
    for line in index_block:
        m = re.match(r"^#\s{3}([A-Za-z][A-Za-z0-9]*)\s+\S", line)
        if m:
            index_keys.append(m.group(1))
    actual_styles = list(data.get("styles", {}).keys())
    if index_keys != actual_styles:
        err("style index mismatch:\n  index  %s\n  actual %s" % (index_keys, actual_styles))

    # 4. 每个样式前有含中文名 / Word 样式 / 指南的注释块
    style_line_re = re.compile(r"^  ([a-zA-Z][A-Za-z0-9]*):\s*$")
    style_lines = {}
    in_styles = False
    for i, line in enumerate(lines):
        if line.rstrip() == "styles:":
            in_styles = True
            continue
        if in_styles and re.match(r"^[a-zA-Z]", line):
            in_styles = False
        if not in_styles:
            continue
        m = style_line_re.match(line)
        if m:
            style_lines[m.group(1)] = i
    for key, idx in style_lines.items():
        block = lines[max(0, idx - 4):idx]
        text = "\n".join(block)
        if not re.search(r"[\u4e00-\u9fff]", text):
            err("style %s: no Chinese comment" % key)
        if "Word 样式" not in text:
            err("style %s: comment missing 'Word 样式'" % key)
        if "指南" not in text:
            err("style %s: comment missing '指南'" % key)

    for key in REQUIRED_STYLES:
        if key not in data.get("styles", {}):
            err("missing required style: %s" % key)

    # 5. 注释行长度
    for i, line in enumerate(lines, 1):
        if line.startswith("#") and len(line) > 100:
            warn("long comment line %d (%d chars): %s" % (i, len(line), line[:60]))

    # 6. 单位 / 度量 / 行距
    def check_units(obj, path=""):
        """检查带单位的值是否使用合法单位。"""
        if isinstance(obj, dict):
            if "unit" in obj:
                unit = obj["unit"]
                if unit not in UNITS:
                    err("invalid unit at %s: %r" % (path, unit))
                if "value" not in obj:
                    err("measurement missing value at %s" % path)
                elif not isinstance(obj["value"], (int, float)):
                    err("measurement value not numeric at %s: %r" % (path, obj["value"]))
            if "rule" in obj and path.endswith("lineSpacing"):
                rule = obj["rule"]
                if rule in ("single", "oneAndHalf", "double"):
                    if "amount" in obj:
                        err("lineSpacing %s must not have amount" % path)
                elif rule == "multiple":
                    amt = obj.get("amount")
                    if not isinstance(amt, dict) or amt.get("unit") != "multiple":
                        err("lineSpacing %s: multiple needs amount.unit=multiple" % path)
                elif rule in ("minimum", "fixed"):
                    amt = obj.get("amount")
                    if not isinstance(amt, dict) or amt.get("unit") != "pt":
                        err("lineSpacing %s: %s needs amount.unit=pt" % (path, rule))
                else:
                    err("invalid lineSpacing rule at %s: %r" % (path, rule))
            for k, v in obj.items():
                check_units(v, path + "." + str(k) if path else str(k))
        elif isinstance(obj, list):
            for i, v in enumerate(obj):
                check_units(v, "%s[%d]" % (path, i))

    check_units(data)

    # 7. 编号引用
    numbering_lists = data.get("numbering", {}).get("lists", {})
    for list_name, spec in numbering_lists.items():
        levels = spec.get("levels", {})
        if not levels:
            err("numbering list %s has no levels" % list_name)
        for k in levels:
            if not isinstance(k, int):
                err("numbering list %s level key not int: %r" % (list_name, k))
    for key, style in data.get("styles", {}).items():
        num = ((style.get("paragraphFormat") or {}).get("numbering"))
        if isinstance(num, dict):
            lst = num.get("list")
            lvl = num.get("level")
            if lst not in numbering_lists:
                err("style %s references unknown numbering list %r" % (key, lst))
            elif lvl not in numbering_lists[lst]["levels"]:
                err("style %s references unknown level %r of list %s" % (key, lvl, lst))

    # 8. documentParts 样式引用
    def collect_style_refs(obj, path=""):
        """收集规则里引用的样式名。"""
        refs = []
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k == "style" and isinstance(v, str):
                    refs.append((v, path))
                refs.extend(collect_style_refs(v, path + "." + str(k) if path else str(k)))
        elif isinstance(obj, list):
            for i, v in enumerate(obj):
                refs.extend(collect_style_refs(v, "%s[%d]" % (path, i)))
        return refs

    for ref, path in collect_style_refs(data.get("documentParts", {})):
        if ref not in data.get("styles", {}):
            err("documentParts references unknown style %r at %s" % (ref, path))

    # 8b. 表格样式引用
    styles = data.get("styles", {})
    tables = data.get("tables", {})
    default_kind = tables.get("default")
    if default_kind is not None and default_kind not in tables:
        err("tables.default references unknown table kind %r" % default_kind)
    for key in ("captionStyle", "noteStyle"):
        ref = tables.get(key)
        if isinstance(ref, str) and ref not in styles:
            err("tables.%s references unknown style %r" % (key, ref))
    for kind, spec in tables.items():
        if isinstance(spec, dict) and isinstance(spec.get("style"), str):
            if spec["style"] not in styles:
                err("tables.%s.style references unknown style %r" % (kind, spec["style"]))

    # 9. 中文取值策略（宽松：允许编号/题注/封面/元数据）
    def has_cjk(s):
        """判断字符串是否含中日韩文字。"""
        return bool(re.search(r"[\u4e00-\u9fff]", s))

    def check_cjk_values(obj, path=""):
        """检查规则取值不使用中文。"""
        if isinstance(obj, dict):
            for k, v in obj.items():
                p = path + "." + str(k) if path else str(k)
                if isinstance(v, str) and has_cjk(v):
                    root = p.split(".")[0]
                    if root not in ("numbering", "captions", "cover", "metadata"):
                        warn("Chinese value outside allowed area at %s: %r" % (p, v))
                else:
                    check_cjk_values(v, p)
        elif isinstance(obj, list):
            for i, v in enumerate(obj):
                check_cjk_values(v, "%s[%d]" % (path, i))

    check_cjk_values(data)

    # 10. 横幅顺序
    banner_order = ["# 3. pageSetup", "# 4. styles", "# 5. numbering", "# 6. captions",
                    "# 7. tables", "# 8. footnotes", "# 9. headersAndFooters",
                    "# 10. cover", "# 11. documentParts", "# 12. appendix"]
    pos = -1
    for banner in banner_order:
        found = raw.find(banner)
        if found < 0:
            err("missing banner: %s" % banner)
        elif found < pos:
            err("banner out of order: %s" % banner)
        else:
            pos = found

    # 11. 实现映射基本引用
    try:
        imap = yaml.safe_load(open(map_path, encoding="utf-8"))
        for key in actual_styles:
            entry = (imap.get("styleMap") or {}).get(key)
            if entry is None:
                err("implementation-map missing styleMap entry: %s" % key)
                continue
            # styleId 不再要求：Word 保存时会重排它，写死没有意义（见文件头注释）。
            if not (entry or {}).get("name"):
                err("implementation-map styleMap entry missing name: %s" % key)
        for key, parent in (imap.get("fallback") or {}).items():
            if key not in actual_styles:
                err("implementation-map fallback unknown key: %s" % key)
            if parent not in actual_styles:
                err("implementation-map fallback unknown parent: %s -> %s" % (key, parent))
        presentation_by_key(imap)
    except Exception as exc:
        err("implementation-map parse error: %s" % exc)

    return errors, warnings


def main():
    """解析参数并校验一个或多个变体规则文件。"""
    from build.variants import load_registry

    ap = argparse.ArgumentParser()
    ap.add_argument("--variant", default=None,
                    help="只校验指定变体（默认 graduate）；与 --all 互斥")
    ap.add_argument("--all", action="store_true", help="校验全部注册变体")
    args = ap.parse_args()

    registry = load_registry()
    if args.all:
        targets = [(name, v.rules, v.implementation_map) for name, v in registry.items()]
    else:
        name = args.variant or "graduate"
        if name not in registry:
            print("未知变体 %r；可选：%s" % (name, ", ".join(registry)))
            return 1
        v = registry[name]
        targets = [(name, v.rules, v.implementation_map)]

    total_errors = 0
    for name, rules_path, map_path in targets:
        errors, warnings = validate_one(str(rules_path), str(map_path), name)
        if errors:
            print("FAILED [%s] (%d errors)" % (name, len(errors)))
            for e in errors:
                print(" -", e)
        else:
            print("OK: rules structure valid [%s]" % name)
        if warnings:
            print("\nWARNINGS [%s] (%d):" % (name, len(warnings)))
            for w in warnings:
                print(" -", w)
        total_errors += len(errors)
    return 1 if total_errors else 0


if __name__ == "__main__":
    sys.exit(main())
