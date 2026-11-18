"""设置 Word 样式窗格的推荐顺序和快速样式库。"""

from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZipFile

from docx.oxml.ns import qn
from lxml import etree
import yaml


# Word 界面从 1 开始，styles.xml 从 0 开始。超过 99 的 XML 值
# 会被 Word 读成界面优先级 1。
MAX_XML_PRIORITY = 99
DISPLAY_TAGS = ("uiPriority", "semiHidden", "unhideWhenUsed", "qFormat")
FOLLOWING_STYLE_TAGS = set(map(qn, (
    "w:locked", "w:personal", "w:personalCompose", "w:personalReply", "w:rsid",
    "w:pPr", "w:rPr", "w:tblPr", "w:trPr", "w:tcPr", "w:tblStylePr",
)))


def presentation_by_key(implementation):
    """解析分组，同时校验每个映射样式恰好分配一次。"""
    groups = implementation.get("stylePresentation")
    if not isinstance(groups, list) or not groups:
        raise ValueError("stylePresentation must be a nonempty list")
    expected = set(implementation["styleMap"])
    result = {}
    for group in groups:
        priority = group.get("priority")
        gallery = group.get("quickStyle", False)
        styles = group.get("styles")
        if type(priority) is not int or not 1 <= priority <= MAX_XML_PRIORITY + 1:
            raise ValueError(f"invalid stylePresentation priority: {priority!r}")
        if type(gallery) is not bool or not isinstance(styles, list) or not styles:
            raise ValueError(f"invalid stylePresentation group: {group!r}")
        for key in styles:
            if not isinstance(key, str) or key not in expected:
                raise ValueError(f"unknown stylePresentation style: {key!r}")
            if key in result:
                raise ValueError(f"duplicate stylePresentation style: {key}")
            result[key] = (priority - 1, gallery)
    if missing := expected - result.keys():
        raise ValueError(f"missing stylePresentation styles: {', '.join(sorted(missing))}")
    return result


def _set_display(style, priority, visible, gallery):
    if visible and (hidden := style.find(qn("w:hidden"))) is not None:
        style.remove(hidden)
    for tag in DISPLAY_TAGS:
        for child in style.findall(qn(f"w:{tag}")):
            style.remove(child)

    index = next((i for i, child in enumerate(style) if child.tag in FOLLOWING_STYLE_TAGS), len(style))
    for offset, (tag, value) in enumerate(zip(DISPLAY_TAGS, (priority, int(not visible), 0, int(gallery)))):
        element = etree.Element(qn(f"w:{tag}"))
        element.set(qn("w:val"), str(value))
        style.insert(index + offset, element)


def apply_style_presentation(root, rule_styles, implementation):
    """更新显式样式和潜在内置样式，保留规则外样式的原有优先级关系。"""
    by_key = presentation_by_key(implementation)
    by_name = {
        implementation["styleMap"][key]["name"].casefold(): by_key[key]
        for key in rule_styles
    }
    explicit = {
        style.find(qn("w:name")).get(qn("w:val")).casefold(): style
        for style in root.findall(qn("w:style"))
    }
    if missing := by_name.keys() - explicit.keys():
        raise ValueError(f"missing Word styles: {', '.join(sorted(missing))}")

    # 内置样式通过别名显示明确的中文用途，保留原名及内置身份供 Word 域使用。
    for key in rule_styles:
        spec = implementation["styleMap"][key]
        if display_name := spec.get("displayName"):
            style = explicit[spec["name"].casefold()]
            for old in style.findall(qn("w:aliases")):
                style.remove(old)
            aliases = etree.Element(qn("w:aliases"))
            aliases.set(qn("w:val"), display_name)
            style.insert(style.index(style.find(qn("w:name"))) + 1, aliases)

    latent = root.find(qn("w:latentStyles"))
    if latent is None:
        latent = etree.Element(qn("w:latentStyles"))
        root.insert(1 if root.find(qn("w:docDefaults")) is not None else 0, latent)
    default = int(latent.get(qn("w:defUIPriority"), str(MAX_XML_PRIORITY)))
    latent_styles = {
        entry.get(qn("w:name")).casefold(): entry
        for entry in latent.findall(qn("w:lsdException"))
    }
    latent_priorities = {
        name: int(entry.get(qn("w:uiPriority"), str(default)))
        for name, entry in latent_styles.items()
    }
    explicit_priorities = {
        name: int(priority.get(qn("w:val"))) if (priority := style.find(qn("w:uiPriority"))) is not None else default
        for name, style in explicit.items()
    }
    outside_priorities = {default}
    for original in (explicit_priorities, latent_priorities):
        outside_priorities.update(priority for name, priority in original.items() if name not in by_name)

    # 规则外各档占据最后的连续位次。论文样式从后往前腾出空间，
    # 尽量保留首选值，同时保持分组顺序。
    first_outside = MAX_XML_PRIORITY + 1 - len(outside_priorities)
    outside_order = {old: first_outside + i for i, old in enumerate(sorted(outside_priorities))}
    rule_order = {}
    ceiling = first_outside
    for priority in sorted({priority for priority, _ in by_name.values()}, reverse=True):
        ceiling = min(priority, ceiling - 1)
        if ceiling < 0:
            raise ValueError("Word priority range cannot represent all style groups")
        rule_order[priority] = ceiling

    for name, style in explicit.items():
        if name in by_name:
            priority, gallery = by_name[name]
            _set_display(style, rule_order[priority], True, gallery)
        else:
            _set_display(style, outside_order[explicit_priorities[name]], False, False)

    latent.set(qn("w:defUIPriority"), str(outside_order[default]))
    latent.set(qn("w:defSemiHidden"), "1")
    latent.set(qn("w:defUnhideWhenUsed"), "0")
    latent.set(qn("w:defQFormat"), "0")
    for name, entry in latent_styles.items():
        if name in by_name:
            priority, gallery = by_name[name]
            values = (rule_order[priority], 0, 0, int(gallery))
        else:
            values = (outside_order[latent_priorities[name]], 1, 0, 0)
        for attr, value in zip(DISPLAY_TAGS, values):
            entry.set(qn(f"w:{attr}"), str(value))


def finalize_style_presentation(path, variant):
    """在 Word 完成保存后写回样式窗格和样式库设置。"""
    from docx.oxml import parse_xml
    from build.hyphenation import apply_all_style_word_wrap, apply_all_paragraph_word_wrap
    path = Path(path)
    rules = yaml.safe_load(variant.rules.read_text(encoding="utf-8"))
    implementation = yaml.safe_load(variant.implementation_map.read_text(encoding="utf-8"))
    with NamedTemporaryFile(dir=path.parent, suffix=".tmp", delete=False) as stream:
        temp = Path(stream.name)
    try:
        with ZipFile(path) as source, ZipFile(temp, "w") as staged:
            for info in source.infolist():
                data = source.read(info)
                if info.filename in ("word/styles.xml", "word/stylesWithEffects.xml"):
                    root = parse_xml(data)
                    apply_style_presentation(root, rules["styles"], implementation)
                    # Word 另存为 DOTM 会省略重复的样式属性，最后再显式恢复。
                    apply_all_style_word_wrap(root, rules["pageSetup"]["documentDefaults"]["paragraphFormat"]["wordWrap"])
                    data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                elif info.filename.startswith("word/") and info.filename.endswith(".xml"):
                    root = parse_xml(data)
                    if apply_all_paragraph_word_wrap(root, rules["pageSetup"]["documentDefaults"]["paragraphFormat"]["wordWrap"]):
                        data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                staged.writestr(info, data)
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)
