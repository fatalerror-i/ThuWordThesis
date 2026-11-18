# -*- coding: utf-8 -*-
"""按变体规则文件生成清华学位论文 Word 模板（.docx）。

用法：python src/build_docx.py --variant graduate 或 --variant undergraduate。
"""
import argparse
import os
import sys
import importlib
from pathlib import Path
import yaml
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_LINE_SPACING,
                            WD_TAB_ALIGNMENT, WD_TAB_LEADER)
from docx.enum.section import WD_SECTION
from docx.enum.table import (WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL,
                             WD_ROW_HEIGHT_RULE)
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from build.style_presentation import apply_style_presentation
from build.hyphenation import apply_document_text_flow, apply_paragraph_text_flow
from build.variants import choices, resolve
from office_metadata import sanitize_office_metadata

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
parser = argparse.ArgumentParser(description="按变体生成未经 Word 更新域的 DOCX。")
parser.add_argument("--variant", required=True, choices=choices())
parser.add_argument("--docx", type=Path, help="DOCX 输出路径，默认取变体配置")
parser.add_argument("--defer-style-presentation", action="store_true", help=argparse.SUPPRESS)
args = parser.parse_args()
VARIANT = resolve(args.variant)
RULES = yaml.safe_load(open(VARIANT.rules, encoding="utf-8"))
IMAP = yaml.safe_load(open(VARIANT.implementation_map, encoding="utf-8"))
OUT = args.docx or VARIANT.docx
OUT.parent.mkdir(parents=True, exist_ok=True)

# 语义键 -> 新模板中的 Word 样式名（中文）
STYLE_NAMES = {
    "body": "Normal",
    "chapterTitle": "Heading 1",
    "chapterTitleUnnumbered": "章标题（无编号）",
    "chapterTitleNoOutline": "章标题（无编号、不进目录）",
    "sectionTitle1": "Heading 2",
    "sectionTitle2": "Heading 3",
    "sectionTitle3": "Heading 4",
    "sectionTitleUnnumbered": "章内无编号小节",
    "committeeSectionTitle": "名单页小节标题",
    "appendixTitle": "附录标题",
    "appendixSectionTitle1": "附录一级节标题",
    "appendixSectionTitle2": "附录二级节标题",
    "appendixSectionTitle3": "附录三级节标题",
    "tocLevel1": "toc 1",
    "tocLevel2": "toc 2",
    "tocLevel3": "toc 3",
    "listOfFiguresEntry": "toc 4",
    "tableOfFigures": "table of figures",
    "symbolEntry": "符号说明条目",
    "caption": "Caption",
    "figureCaption": "图注",
    "subfigureCaption": "子图注",
    "appendixFigureCaption": "附录图注",
    "appendixTableCaption": "附录表注",
    "tableCaption": "表注",
    "figureTableNote": "图表附注",
    "tableText": "表格文字",
    "threeLineTable": "三线表",
    "image": "图片",
    "formula": "公式",
    "inlineFormula": "行内公式",
    "referenceEntry": "参考文献条目",
    "referenceEntryAuthorYear": "参考文献条目（著者出版年）",
    "citation": "文献引用",
    "subtleEmphasis": "Subtle Emphasis",
    "emphasis": "Emphasis",
    "strong": "Strong",
    "header": "Header",
    "footer": "Footer",
    "footnoteText": "Footnote Text",
    "footnoteReference": "footnote reference",
    "hyperlink": "Hyperlink",
    "coverTitleChinese": "封面论文题目",
    "coverDegreeInfo": "封面学位类别",
    "coverAuthorInfo": "封面作者信息",
    "coverSecurity": "封面密级",
    "coverDate": "封面日期",
    "coverTitleEnglish": "英文封面题目",
    "coverInfoEnglish": "英文封面信息",
    "coverSupervisorEnglish": "英文封面导师信息",
    "authorizationTitle": "授权说明标题",
    "authorizationText": "授权说明文字",
    "listNumber1": "List Number",
    "listNumber2": "List Number 2",
    "listNumber3": "List Number 3",
    "listBullet1": "List Bullet",
    "listBullet2": "List Bullet 2",
    "listBullet3": "List Bullet 3",
    "inlineCode": "行内代码",
    "blockCode": "行间代码",
}
FALLBACK = IMAP.get("fallback", {})
CHARACTER_STYLE_KEYS = {
    key for key, spec in (RULES.get("styles") or {}).items()
    if (spec or {}).get("styleType") == "character"
}
# 表格样式由 add_table_style() 单独建（w:type="table"），不能走段落/字符样式循环。
TABLE_STYLE_KEYS = {
    key for key, spec in (RULES.get("styles") or {}).items()
    if (spec or {}).get("styleType") == "table"
}

ALIGN = {
    "justified": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}
RULE = {
    "single": WD_LINE_SPACING.SINGLE,
    "oneAndHalf": WD_LINE_SPACING.ONE_POINT_FIVE,
    "double": WD_LINE_SPACING.DOUBLE,
    "multiple": WD_LINE_SPACING.MULTIPLE,
    "minimum": WD_LINE_SPACING.AT_LEAST,
    "fixed": WD_LINE_SPACING.EXACTLY,
}


# ---------------------------------------------------------------------------
# XML 辅助函数
# ---------------------------------------------------------------------------
def _el(tag, **attrs):
    """创建带 w: 命名空间属性的 XML 元素。"""
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def _child(parent, tag, **attrs):
    """创建子元素并追加到父元素。"""
    e = _el(tag, **attrs)
    parent.append(e)
    return e


def _clear_ppr_tags(pPr, *tags):
    """从段落属性中移除指定标签。"""
    for tag in tags:
        for e in pPr.findall(qn("w:" + tag)):
            pPr.remove(e)


def set_rfonts(rPr, chinese=None, western=None):
    """把中西文字体声明写入 run 属性。"""
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = _el("w:rFonts")
        rPr.insert(0, rFonts)
    if western:
        rFonts.set(qn("w:ascii"), western)
        rFonts.set(qn("w:hAnsi"), western)
        rFonts.set(qn("w:cs"), western)
    if chinese:
        rFonts.set(qn("w:eastAsia"), chinese)
        # 中西文共用的字符（全角引号、破折号等）按中文字体渲染。
        rFonts.set(qn("w:hint"), "eastAsia")


def apply_character(style, cf):
    """把规则里的字符格式套用到样式。"""
    if not cf:
        return
    fam = cf.get("fontFamily") or {}
    if fam:
        if fam.get("western"):
            style.font.name = fam["western"]
    if "fontSize" in cf:
        style.font.size = Pt(cf["fontSize"]["value"])
    if cf.get("bold") is not None:
        style.font.bold = bool(cf["bold"])
    if cf.get("italic") is not None:
        style.font.italic = bool(cf["italic"])
    if cf.get("superscript") is not None:
        style.font.superscript = bool(cf["superscript"])
    rPr = style.element.get_or_add_rPr()
    if cf.get("underline") is not None:
        # 「无下划线」必须显式写 <w:u w:val="none"/>：Word 内置的超链接样式
        # 自带 single 下划线，只靠不写属性是清不掉的。
        _clear_ppr_tags(rPr, "u")
        u = _el("w:u", val="none" if cf["underline"] in (False, "none") else "single")
        rPr.insert_element_before(u, "w:effect", "w:bdr", "w:shd", "w:fitText",
                                  "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
                                  "w:eastAsianLayout", "w:specVanish", "w:oMath")
    if cf.get("color"):
        # 只写 w:val，不写 themeColor：带主题色时 Word 点击后会换成
        # FollowedHyperlink 的主题色（紫色）。
        _clear_ppr_tags(rPr, "color")
        color = _el("w:color", val=str(cf["color"]).upper())
        rPr.insert_element_before(color, "w:spacing", "w:w", "w:kern", "w:position",
                                  "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
                                  "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl",
                                  "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
                                  "w:specVanish", "w:oMath")
    set_rfonts(rPr, (fam or {}).get("chinese"), (fam or {}).get("western"))
    if cf.get("characterSpacing"):
        _clear_ppr_tags(rPr, "spacing")
        sp = _el("w:spacing", val=int(round(cf["characterSpacing"]["value"] * 20)))
        rPr.append(sp)


def apply_paragraph(style, pf, base_font_pt=12):
    """把规则里的段落格式套用到样式。"""
    if not pf:
        return
    p = style.paragraph_format
    if "widowControl" in pf:
        p.widow_control = pf["widowControl"]
    if "alignment" in pf:
        p.alignment = ALIGN[pf["alignment"]]
    ls = pf.get("lineSpacing")
    if ls:
        rule = RULE[ls["rule"]]
        if ls["rule"] in ("single", "oneAndHalf", "double"):
            p.line_spacing = {"single": 1.0, "oneAndHalf": 1.5, "double": 2.0}[ls["rule"]]
        elif ls["rule"] == "multiple":
            p.line_spacing = ls["amount"]["value"]
        else:  # 最小值 / 固定值
            p.line_spacing = Pt(ls["amount"]["value"])
        p.line_spacing_rule = rule
    if "spaceBefore" in pf:
        p.space_before = Pt(pf["spaceBefore"]["value"])
    if "spaceAfter" in pf:
        p.space_after = Pt(pf["spaceAfter"]["value"])
    pPr = style.element.get_or_add_pPr()
    apply_paragraph_text_flow(pPr, pf)
    # 以字符为单位的缩进
    ind = None
    if any(k in pf for k in ("firstLineIndent", "leftIndent", "hangingIndent")):
        ind = pPr.get_or_add_ind()
    if "firstLineIndent" in pf:
        chars = pf["firstLineIndent"]["value"]
        ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))
        ind.set(qn("w:firstLine"), str(int(round(chars * base_font_pt * 20))))
    if "leftIndent" in pf:
        left = pf["leftIndent"]
        unit = left.get("unit", "character")
        value = float(left["value"])
        if unit == "character":
            ind.set(qn("w:leftChars"), str(int(round(value * 100))))
            ind.set(qn("w:left"), str(int(round(value * base_font_pt * 20))))
        else:
            twips_per_unit = {"cm": 566.929, "mm": 56.6929, "pt": 20}[unit]
            if ind.get(qn("w:leftChars")) is not None:
                del ind.attrib[qn("w:leftChars")]
            ind.set(qn("w:left"), str(int(round(value * twips_per_unit))))
    if "hangingIndent" in pf:
        hang = pf["hangingIndent"]
        unit = hang.get("unit", "character")
        value = float(hang["value"])
        if unit == "character":
            ind.set(qn("w:hangingChars"), str(int(round(value * 100))))
            ind.set(qn("w:hanging"), str(int(round(value * base_font_pt * 20))))
        else:
            # 绝对单位：同时写左缩进与悬挂缩进，并清除字符单位属性，
            # 避免 Word 用 hangingChars/leftChars 覆盖厘米值。
            twips_per_unit = {"cm": 566.929, "mm": 56.6929, "pt": 20}[unit]
            twips = str(int(round(value * twips_per_unit)))
            for _attr in ("w:hangingChars", "w:leftChars"):
                if ind.get(qn(_attr)) is not None:
                    del ind.attrib[qn(_attr)]
            ind.set(qn("w:left"), twips)
            ind.set(qn("w:hanging"), twips)
    if pf.get("outlineLevel") is not None:
        _clear_ppr_tags(pPr, "outlineLvl")
        lvl = pf["outlineLevel"]
        val = 9 if lvl == "bodyText" else int(lvl) - 1
        _child(pPr, "w:outlineLvl", val=val)
    for flag, tag, after in (("autoSpaceDE", "autoSpaceDE", "w:autoSpaceDN"),
                             ("autoSpaceDN", "autoSpaceDN", "w:bidi")):
        if flag in pf:
            _clear_ppr_tags(pPr, tag)
            tag = "w:" + tag
            pPr.insert_element_before(
                _el(tag, val="1" if pf[flag] else "0"),
                after, "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing",
                "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
                "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
                "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
                "w:pPrChange")
    for flag, attr in (("keepLinesTogether", "keep_together"),
                       ("keepWithNext", "keep_with_next"),
                       ("pageBreakBefore", "page_break_before")):
        if flag in pf:
            setattr(p, attr, bool(pf[flag]))
    border = pf.get("borderBottom")
    if border:
        _clear_ppr_tags(pPr, "pBdr")
        pBdr = _el("w:pBdr")
        bottom = _el("w:bottom", val="single",
                     sz=int(round(border["width"]["value"] * 8)),
                     space=int(border.get("spacing", {}).get("value", 1)),
                     color="auto")
        pBdr.append(bottom)
        pPr.insert_element_before(
            pBdr, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
            "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
            "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
            "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
            "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
            "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
            "w:rPr", "w:sectPr", "w:pPrChange")
    if "tabs" in pf:
        # tabs: [] 表示「本样式不带任何制表位」——必须先把继承来的制表位清掉，
        # 否则 Word 内置样式自带的制表位（如页眉的居中/右对齐制表位）会留下来。
        _clear_ppr_tags(pPr, "tabs")
    for tab in pf.get("tabs", []) or []:
        align = {"left": WD_TAB_ALIGNMENT.LEFT, "right": WD_TAB_ALIGNMENT.RIGHT,
                 "center": WD_TAB_ALIGNMENT.CENTER}[tab["alignment"]]
        leader = {"dots": WD_TAB_LEADER.DOTS, "none": WD_TAB_LEADER.SPACES}.get(
            tab.get("leader", "none"), WD_TAB_LEADER.SPACES)
        p.tab_stops.add_tab_stop(Cm(tab["position"]["value"]), align, leader)


def apply_paragraph_to_ppr(pPr, pf, size_pt=12):
    """把规则里的段落格式写到裸 w:pPr 上（docDefaults 的 pPrDefault 用）。"""
    attrs = {}
    ls = pf.get("lineSpacing")
    if ls:
        rule = ls["rule"]
        if rule in ("single", "oneAndHalf", "double"):
            attrs["line"] = {"single": 240, "oneAndHalf": 360, "double": 480}[rule]
            attrs["lineRule"] = "auto"
        elif rule == "multiple":
            attrs["line"] = int(round(ls["amount"]["value"] * 240))
            attrs["lineRule"] = "auto"
        else:
            attrs["line"] = int(round(ls["amount"]["value"] * 20))
            attrs["lineRule"] = "atLeast" if rule == "minimum" else "exact"
    if "spaceBefore" in pf:
        attrs["before"] = int(round(pf["spaceBefore"]["value"] * 20))
    if "spaceAfter" in pf:
        attrs["after"] = int(round(pf["spaceAfter"]["value"] * 20))
    if not attrs:
        return
    _clear_ppr_tags(pPr, "spacing")
    sp = _el("w:spacing", **attrs)
    pPr.insert_element_before(sp, "w:ind", "w:contextualSpacing", "w:mirrorIndents",
                              "w:suppressOverlap", "w:jc", "w:textDirection",
                              "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                              "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")


def _style_sort_key(key):
    """按继承深度返回样式创建顺序的排序键。"""
    depth = 0
    cur = key
    while cur in FALLBACK:
        depth += 1
        cur = FALLBACK[cur]
        if depth > 50:
            break
    return depth


# ---------------------------------------------------------------------------
# 编号
# ---------------------------------------------------------------------------
def build_numbering(doc):
    """按规则构建多级编号定义。"""
    numbering = doc.part.numbering_part.element
    abs_ids = [int(e.get(qn("w:abstractNumId"))) for e in numbering.findall(qn("w:abstractNum"))]
    num_ids = [int(e.get(qn("w:numId"))) for e in numbering.findall(qn("w:num"))]
    base_abs = max(abs_ids + [0]) + 100
    base_num = max(num_ids + [0]) + 100

    # OOXML 的 w:suff 只接受 tab / space / nothing；规则文件用更直观的
    # "none" 表示"无分隔符"，需映射为 "nothing"。
    _SUFFIX_MAP = {"none": "nothing", "nothing": "nothing",
                   "tab": "tab", "space": "space"}

    def make_abs(aid, levels, fullwidth=False, hanging_chars=0):
        """构建一份抽象多级编号定义。"""
        abs_el = _el("w:abstractNum", abstractNumId=aid)
        _child(abs_el, "w:multiLevelType", val="multilevel")
        for ilvl, spec in sorted(levels.items()):
            lvl = _el("w:lvl", ilvl=ilvl - 1)
            _child(lvl, "w:start", val=spec.get("start", 1))
            _child(lvl, "w:numFmt", val=spec["format"])
            text = spec["text"] + ("\u3000" if fullwidth else "")
            _child(lvl, "w:lvlText", val=text)
            _raw_suffix = spec.get("suffix", "tab")
            _child(lvl, "w:suff",
                   val=_SUFFIX_MAP.get(_raw_suffix, _raw_suffix))
            if spec.get("font"):
                rPr = _el("w:rPr")
                _child(rPr, "w:rFonts", ascii=spec["font"], hAnsi=spec["font"],
                       eastAsia=spec["font"], cs=spec["font"])
                lvl.append(rPr)
            _child(lvl, "w:lvlJc", val="left")
            pPr = _el("w:pPr")
            if hanging_chars:
                # 只写悬挂缩进：编号/项目符号首行顶格，续行缩进。
                # 不能再写 w:left，否则它与样式层的悬挂缩进叠加，产生额外左缩进。
                _child(pPr, "w:ind",
                       hanging=hanging_chars * 12 * 20,
                       hangingChars=hanging_chars * 100)
            else:
                _child(pPr, "w:ind", left=0, firstLine=0)
            lvl.append(pPr)
            abs_el.append(lvl)
        return abs_el

    specs = [("outline", RULES["numbering"]["lists"]["outline"]["levels"], False),
             ("appendix", RULES["numbering"]["lists"]["appendix"]["levels"], False),
             ("references", RULES["numbering"]["lists"]["references"]["levels"], False),
             ("codeLines", RULES["numbering"]["lists"]["codeLines"]["levels"], False),
             ("listNumber1", RULES["numbering"]["lists"]["listNumber1"]["levels"], False),
             ("listNumber2", RULES["numbering"]["lists"]["listNumber2"]["levels"], False),
             ("listNumber3", RULES["numbering"]["lists"]["listNumber3"]["levels"], False),
             ("listBullet1", RULES["numbering"]["lists"]["listBullet1"]["levels"], False),
             ("listBullet2", RULES["numbering"]["lists"]["listBullet2"]["levels"], False),
             ("listBullet3", RULES["numbering"]["lists"]["listBullet3"]["levels"], False),
             ("achievements", RULES["numbering"]["lists"]["achievements"]["levels"], False),
             ("appendixReferences", RULES["numbering"]["lists"]["appendixReferences"]["levels"], False)]
    first_num = numbering.find(qn("w:num"))
    for i, (name, levels, fullwidth) in enumerate(specs):
        _hanging_chars = 2 if name.startswith("list") else 0
        el = make_abs(base_abs + i, levels, fullwidth, _hanging_chars)
        if first_num is not None:
            first_num.addprevious(el)
        else:
            numbering.append(el)
    ids = {}
    for i, (name, levels, fullwidth) in enumerate(specs):
        nid = base_num + i
        num_el = _el("w:num", numId=nid)
        num_el.append(_el("w:abstractNumId", val=base_abs + i))
        if name == "achievements":
            # Word 保存时可能会合并格式相同的列表，显式重启个人成果编号。
            override = _child(num_el, "w:lvlOverride", ilvl=0)
            _child(override, "w:startOverride", val=levels[1]["start"])
        numbering.append(num_el)
        ids[name] = (base_abs + i, nid)
    return ids


PPR_AFTER_NUMPR = (
    "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
    "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
    "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing",
    "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap",
    "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)


def set_style_numbering(style, num_id, level):
    """把样式绑定到编号 id 与级别。"""
    pPr = style.element.get_or_add_pPr()
    for e in pPr.findall(qn("w:numPr")):
        pPr.remove(e)
    numPr = _el("w:numPr")
    numPr.append(_el("w:ilvl", val=level - 1))
    numPr.append(_el("w:numId", val=num_id))
    pPr.insert_element_before(numPr, *PPR_AFTER_NUMPR)


def disable_style_numbering(style):
    """移除样式的编号绑定。"""
    pPr = style.element.get_or_add_pPr()
    for e in pPr.findall(qn("w:numPr")):
        pPr.remove(e)
    numPr = _el("w:numPr")
    numPr.append(_el("w:numId", val=0))
    pPr.insert_element_before(numPr, *PPR_AFTER_NUMPR)


def set_paragraph_numbering(p, num_id, level=1):
    """覆盖段落编号（用于简历/附录列表）。"""
    pPr = p._p.get_or_add_pPr()
    for e in pPr.findall(qn("w:numPr")):
        pPr.remove(e)
    numPr = _el("w:numPr")
    numPr.append(_el("w:ilvl", val=level - 1))
    numPr.append(_el("w:numId", val=num_id))
    pPr.insert_element_before(numPr, *PPR_AFTER_NUMPR)


# ---------------------------------------------------------------------------
# 域 / 分节
# ---------------------------------------------------------------------------
def add_field(paragraph, instr, result="", *, locked=False, char_style=None):
    """在段落里插入 Word 域。"""
    r = paragraph.add_run()
    begin = _el("w:fldChar", fldCharType="begin")
    if locked:
        begin.set(qn("w:fldLock"), "true")
    r._r.append(begin)
    r2 = paragraph.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = " " + instr + " "
    r2._r.append(it)
    r3 = paragraph.add_run()
    r3._r.append(_el("w:fldChar", fldCharType="separate"))
    if result:
        paragraph.add_run(result, style=char_style)
    r4 = paragraph.add_run()
    r4._r.append(_el("w:fldChar", fldCharType="end"))


# 题注书签沿用 VBA 宏预期的名称（THU_FIG_0001、THU_TAB_0001），
# 功能区命令据此找到已有书签，而不是在同一编号上再加一个。
CAPTION_BOOKMARK_PREFIX = {
    "图": "THU_FIG_", "附录图": "THU_FIG_",
    "表": "THU_TAB_", "附录表": "THU_TAB_",
}

_BOOKMARK_ID = [9000]


def add_bookmark_start(paragraph, name):
    """插入书签起点并分配唯一 id。"""
    _BOOKMARK_ID[0] += 1
    bid = _BOOKMARK_ID[0]
    paragraph._p.append(_el("w:bookmarkStart", id=bid, name=name))
    return bid


def add_bookmark_end(paragraph, bid):
    """插入与给定 id 匹配的书签终点。"""
    paragraph._p.append(_el("w:bookmarkEnd", id=bid))


_OMATH_TAG = ("{http://schemas.openxmlformats.org/officeDocument/2006/math}"
              "oMath")


def _has_omath(paragraph):
    """段落里是否有 Word 公式（行间公式包在 m:oMathPara 里）。"""
    return any(el.tag == _OMATH_TAG for el in paragraph._p.iter())


def _caption_scan():
    """扫描已写入的章节、题注与公式，供缓存域结果使用。

    SEQ 域在每个一级标题后重新编号，所以缓存值等于章号加上该章内的序号；
    书签序号在全文档范围内连续。公式段落由 finalize_docx.py 按文档顺序编号
    并加书签，这里用同样的顺序数出来，附录内另外单独计数。
    """
    chapter = 0
    appendix = 0
    counts = {}
    ords = {"THU_FIG_": 0, "THU_TAB_": 0}
    eq_total = 0
    eq_app = 0
    chapter_style = STYLE_NAMES["chapterTitle"]
    appendix_style = STYLE_NAMES["appendixTitle"]
    formula_style = STYLE_NAMES["formula"]
    for p in doc.paragraphs:
        name = p.style.name
        if name == chapter_style:
            chapter += 1
            counts = {}
        elif name == appendix_style:
            appendix += 1
            counts = {}
            eq_app = 0
        if name == formula_style and _has_omath(p):
            eq_total += 1
            if appendix:
                eq_app += 1
        for it in p._p.iter(qn("w:instrText")):
            text = (it.text or "").strip()
            if not text.startswith("SEQ "):
                continue
            parts = text.split()
            if len(parts) < 2 or "\\r" in text:
                continue
            ident = parts[1]
            counts[ident] = counts.get(ident, 0) + 1
            prefix = CAPTION_BOOKMARK_PREFIX.get(ident)
            if prefix:
                ords[prefix] += 1
    return {"chapter": chapter, "appendix": appendix, "counts": counts,
            "ords": ords, "eq_total": eq_total, "eq_app": eq_app}


def peek_caption(seq_id, offset=0):
    """返回稍后写入的题注的书签名与编号。

    Word 在更新 REF 域时才解析它，因此引用可以出现在题注之前。
    两侧都从同一次扫描里取得书签名与缓存编号，
    从而保持一致。
    """
    scan = _caption_scan()
    prefix = CAPTION_BOOKMARK_PREFIX[seq_id]
    bookmark = "%s%04d" % (prefix, scan["ords"][prefix] + offset + 1)
    index = scan["counts"].get(seq_id, 0) + offset + 1
    if seq_id in ("附录图", "附录表"):
        appendix = max(scan["appendix"], 1)
        number = "%s.%d" % (chr(ord("A") + appendix - 1), index)
    else:
        number = "%d.%d" % (scan["chapter"], index)
    return bookmark, number


def ref_caption(paragraph, seq_id, prefix_text="", offset=0):
    """为稍后写入的题注写入 REF 域。

    offset 从已写入的题注往前计数，因此位于表格之前的段落
    也能正确引用它。
    """
    bookmark, number = peek_caption(seq_id, offset)
    add_ref_field(paragraph, bookmark, number, prefix_text)
    return number


def peek_formula(offset=-1):
    """返回稍后写入的公式编号的书签名与编号。

    offset 沿用 peek_caption：-1 表示刚写入的公式，0 表示下一个。
    finalize_docx.py 按文档顺序给「公式」段落编号，并从 THU_EQ_0001
    起加书签，所以书签序号按全文计数。附录通过 add_seq_reset 重启
    「式」，那里的缓存编号是附录字母加上附录内的序号。
    """
    scan = _caption_scan()
    index = scan["eq_total"] + offset + 1
    bookmark = "THU_EQ_%04d" % index
    if scan["appendix"]:
        app_index = scan["eq_app"] + offset + 1
        if app_index >= 1:
            letter = chr(ord("A") + scan["appendix"] - 1)
            return bookmark, "%s.%d" % (letter, app_index)
    return bookmark, "%d.%d" % (scan["chapter"], index)


def ref_formula(paragraph, offset=-1, prefix_text=""):
    """为 finalize_docx.py 加书签的公式编号写入 REF 域。"""
    bookmark, number = peek_formula(offset)
    add_ref_field(paragraph, bookmark, number, prefix_text)


def add_ref_field(paragraph, bookmark, number, prefix_text=""):
    """写入指向题注书签的 REF 域。

    number 是缓存的域结果。首次更新域之前 Word 就显示它，而
    finalize_docx.py 不更新 REF 域，所以它一开始就得带对文字。
    """
    if prefix_text:
        paragraph.add_run(prefix_text)
    add_field(paragraph, "REF %s \\h" % bookmark, number)


def clear_header_footer(part):
    """清空页眉或页脚，只留一个空段落。"""
    for p in list(part.paragraphs)[1:]:
        p._element.getparent().remove(p._element)
    p = part.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    p.text = ""
    return p


def _clear_para_first_line(p):
    """把段落首行缩进清零。"""
    p.paragraph_format.first_line_indent = Pt(0)
    pPr = p._p.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")


def set_header(sec, kind=None):
    """写入页眉文字与章标题域。"""
    hdr = sec.header
    hdr.is_linked_to_previous = False
    p = clear_header_footer(hdr)
    _clear_para_first_line(p)
    # 规则是唯一事实来源：headersAndFooters.header.content = none 表示整份
    # 模板不设页眉（本科综合论文训练规范：页眉：无），此时忽略调用方传入的
    # 页眉类型，所有节一律留空页眉。
    _header_spec = (RULES.get("headersAndFooters") or {}).get("header") or {}
    if _header_spec.get("content") == "none":
        kind = None
    if not kind:
        # 封面不能继承"页眉"样式（默认模板里它带下框线），
        # 留一个空白正文段落即可。
        p.style = doc.styles[STYLE_NAMES["body"]]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        clear_indent(p)
        return
    p.style = doc.styles[STYLE_NAMES["header"]]
    if kind == "numbered":
        add_field(p, 'STYLEREF 1 \\n', "第1章　")
        add_field(p, 'STYLEREF 1', "章标题")
    elif kind == "unchaptered":
        add_field(p, 'STYLEREF "章标题（无编号）"', "摘要")
    elif kind == "appendix":
        add_field(p, 'STYLEREF "附录标题" \\n', "附录A　")
        add_field(p, 'STYLEREF "附录标题"', "附录标题")


def set_footer(sec, page=True):
    """写入页脚页码或留空。"""
    ftr = sec.footer
    ftr.is_linked_to_previous = False
    p = clear_header_footer(ftr)
    _clear_para_first_line(p)
    if not page:
        p.style = doc.styles[STYLE_NAMES["body"]]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        clear_indent(p)
        return
    p.style = doc.styles[STYLE_NAMES["footer"]]
    add_field(p, "PAGE", "1")


def set_pgnum(sec, fmt=None, start=None):
    """设置分节的页码格式与起始值。"""
    sectPr = sec._sectPr
    for e in sectPr.findall(qn("w:pgNumType")):
        sectPr.remove(e)
    if fmt or start:
        el = _el("w:pgNumType")
        if fmt:
            el.set(qn("w:fmt"), fmt)
        if start is not None:
            el.set(qn("w:start"), str(start))
        sectPr.append(el)


def set_vertical(sec):
    """把分节文字方向设为竖排。"""
    sectPr = sec._sectPr
    for e in sectPr.findall(qn("w:textDirection")):
        sectPr.remove(e)
    sectPr.append(_el("w:textDirection", val="tbRl"))


def clear_vertical(sec):
    """移除分节的竖排文字方向。"""
    sectPr = sec._sectPr
    for e in sectPr.findall(qn("w:textDirection")):
        sectPr.remove(e)


def setup_section(sec, top, bottom, left, right, header_dist, footer_dist,
                  header_kind=None, page_footer=True, pgfmt=None, start=None,
                  vertical=False, gutter=0.0):
    """按规则设置页面尺寸、页边距、页眉页脚与分节属性。"""
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(top), Cm(bottom)
    sec.left_margin, sec.right_margin = Cm(left), Cm(right)
    sec.header_distance = Cm(header_dist)
    sec.footer_distance = Cm(footer_dist)
    # 装订线按本节规则显式写入：add_section() 会逐字克隆上一节 sectPr，
    # 不显式赋值就会让封面节的装订线漏到后面所有节（本科 0.2 cm 即如此）。
    sec.gutter = Cm(gutter)
    clear_vertical(sec)
    if vertical:
        set_vertical(sec)
    set_header(sec, header_kind)
    set_footer(sec, page_footer)
    set_pgnum(sec, pgfmt, start)
    grid_type = {"none": "default"}[RULES["pageSetup"]["documentGrid"]]
    sec._sectPr.remove_all("w:docGrid")
    sec._sectPr.insert_element_before(
        _el("w:docGrid", type=grid_type), "w:printerSettings", "w:sectPrChange"
    )


def margin_cm(margins, key):
    """读取规则中的页边距值；页边距统一用厘米，避免字符单位被 Word 叠加。"""
    item = margins[key]
    if item.get("unit", "cm") != "cm":
        raise ValueError(f"页边距规则只支持厘米：{key} = {item!r}")
    return float(item["value"])


def gutter_cm(margins):
    """装订线：规则未写时按 0 处理（页边距里的 gutter 可选）。"""
    item = margins.get("gutter")
    if not item:
        return 0.0
    if item.get("unit", "cm") != "cm":
        raise ValueError(f"装订线规则只支持厘米：gutter = {item!r}")
    return float(item["value"])


def section_margins(name):
    """取某节自定义的 pageMargins；该节未单列时继承默认页边距。"""
    page_setup = RULES["pageSetup"]
    spec = (page_setup.get("sections") or {}).get(name) or {}
    return spec.get("pageMargins") or page_setup["defaultPageMargins"]


def chapter_section(header_kind=None, page_footer=True, pgfmt=None, start=None):
    """以"下一页"分节符开始一个章级部分。

    各章级部分都用这个辅助函数，保证页面设置与页眉页脚继承一致；
    除非显式设置起始值，页码继续累计。页边距与页眉页脚距离取自
    pageSetup 规则，让 YAML 保持唯一事实来源。
    """
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    margins = section_margins("body")
    setup_section(sec, margin_cm(margins, "top"), margin_cm(margins, "bottom"),
                  margin_cm(margins, "left"), margin_cm(margins, "right"),
                  margin_cm(margins, "header"), margin_cm(margins, "footer"),
                  header_kind, page_footer, pgfmt, start,
                  gutter=gutter_cm(margins))
    return sec


# ---------------------------------------------------------------------------
# 内容辅助函数
# ---------------------------------------------------------------------------
doc = Document()
doc.core_properties.title = VARIANT.title
doc.core_properties.subject = "Word 模板：格式与使用说明"

# 打开时更新域
settings = doc.settings.element
footnote_pr = settings.find(qn("w:footnotePr"))
if footnote_pr is None:
    footnote_pr = OxmlElement("w:footnotePr")
    settings.insert_element_before(footnote_pr, "w:endnotePr", "w:compat",
                                   "w:docVars", "w:rsids", "m:mathPr")
# CT_FtnDocProps 的子元素有固定顺序：numFmt/numRestart 必须在 footnote 之前。
footnote_pr.insert(0, _el("w:numFmt", val="decimalEnclosedCircle"))
footnote_pr.insert(1, _el("w:numRestart", val="eachPage"))
upd = OxmlElement("w:updateFields")
upd.set(qn("w:val"), "true")
settings.append(upd)

# 数学字体由规则指定（当前两变体均为 XITS Math）。python-docx 默认模板的
# m:mathFont 是 Cambria Math；若不改，Word 保存时会给每个数学 run 注入
# Cambria Math 直接字体，覆盖「行内公式」字符样式里的 XITS Math。
_math_font = (RULES["styles"]["inlineFormula"]["characterFormat"]
              ["fontFamily"]["western"])
_math_pr = settings.find(qn("m:mathPr"))
if _math_pr is None:
    _math_pr = OxmlElement("m:mathPr")
    settings.insert_element_before(_math_pr, "w:updateFields")
_math_font_el = _math_pr.find(qn("m:mathFont"))
if _math_font_el is None:
    _math_font_el = OxmlElement("m:mathFont")
    _math_pr.insert(0, _math_font_el)
_math_font_el.set(qn("m:val"), _math_font)


def set_doc_defaults(cjk, western, size_pt, pf=None):
    """默认段落字体 / 文档默认值：基础字体的唯一来源。"""
    styles_el = doc.styles.element
    dd = styles_el.find(qn("w:docDefaults"))
    rPrD = dd.find(qn("w:rPrDefault"))
    rPr = rPrD.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        rPrD.append(rPr)
    for tag in ("w:rFonts", "w:sz", "w:szCs"):
        for e in rPr.findall(qn(tag)):
            rPr.remove(e)
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), western)
    rFonts.set(qn("w:hAnsi"), western)
    rFonts.set(qn("w:cs"), western)
    rFonts.set(qn("w:eastAsia"), cjk)
    # 全角引号等“东亚宽度”字符使用中文字体，而不是西文字体。
    rFonts.set(qn("w:hint"), "eastAsia")
    rPr.insert(0, rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(round(size_pt * 2))))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(round(size_pt * 2))))
    rPr.append(szCs)
    if pf:
        _set_ppr_default(dd, pf)


def _set_ppr_default(dd, pf):
    """w:pPrDefault：文档级段落默认值（段前后 0、行距最小值 20 磅）。

    python-docx 的最小模板里残留着 spacing after=200 line=276 auto，
    不覆盖掉的话，任何没写行距的段落都会带着 10 磅段后、1.15 倍行距。
    """
    pPrD = dd.find(qn("w:pPrDefault"))
    if pPrD is None:
        pPrD = OxmlElement("w:pPrDefault")
        # CT_DocDefaults 的子元素顺序固定：rPrDefault → pPrDefault。
        rPrD = dd.find(qn("w:rPrDefault"))
        if rPrD is None:
            dd.insert(0, pPrD)
        else:
            rPrD.addnext(pPrD)
    pPr = pPrD.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        pPrD.append(pPr)
    apply_paragraph_to_ppr(pPr, pf, size_pt=12)


_doc_defaults = (RULES.get("pageSetup") or {}).get("documentDefaults") or {}
_body_cf = RULES["styles"]["body"]["characterFormat"]
_dd_cf = _doc_defaults.get("characterFormat") or {}
_dd_fam = _dd_cf.get("fontFamily") or {}
set_doc_defaults(_dd_fam.get("chinese") or _body_cf["fontFamily"]["chinese"],
                 _dd_fam.get("western") or _body_cf["fontFamily"]["western"],
                 (_dd_cf.get("fontSize") or _body_cf["fontSize"])["value"],
                 _doc_defaults.get("paragraphFormat"))

NUM = build_numbering(doc)

# --- 样式 -----------------------------------------------------------------
# 规则与 Word 内置样式一致时直接复用，使域生成的内容
# （页眉/页脚/题注/目录）也能套上对应格式。
BUILTIN = {}


def _ensure_builtin_paragraph(name, style_id=None):
    """取一个内置段落样式，若仅为潜在状态则先创建。"""
    try:
        st = doc.styles[name]
    except KeyError:
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=True)
    if style_id and st.element.get(qn("w:styleId")) != style_id:
        st.element.set(qn("w:styleId"), style_id)
    return st


# python-docx 的最小默认模板里没有 TOC 1-4，必须在规则循环之前建好，
# 否则旧代码会另建名为"目录 1"之类的自定义样式，而 Word 的目录域
# 根本不会用到它们。
for _key, _name, _sid in (
        ("tocLevel1", "toc 1", "TOC1"),
        ("tocLevel2", "toc 2", "TOC2"),
        ("tocLevel3", "toc 3", "TOC3"),
        ("listOfFiguresEntry", "toc 4", "TOC4"),
        # \c "图" / \c "表" 域生成的清单条目由 Word 用内置样式 “table of figures”
        # 实例化；python-docx 的最小模板里没有它，必须按内置样式自建。
        ("tableOfFigures", "table of figures", "TableofFigures"),
):
    _st = _ensure_builtin_paragraph(_name, _sid)
    BUILTIN[_key] = _st
    STYLE_NAMES[_key] = _st.name

# python-docx 的最小模板里同样没有"脚注文本"样式。
_st = _ensure_builtin_paragraph("Footnote Text", "FootnoteText")
BUILTIN["footnoteText"] = _st
STYLE_NAMES["footnoteText"] = _st.name

# 模板只定义 1–3 级编号与项目符号；保留这些内置样式的身份。
# 它们都是内置样式，不是自定义新增。
for _key, _name in (
        ("listNumber1", "List Number"), ("listNumber2", "List Number 2"),
        ("listNumber3", "List Number 3"),
        ("listBullet1", "List Bullet"), ("listBullet2", "List Bullet 2"),
        ("listBullet3", "List Bullet 3"),
):
    _st = _ensure_builtin_paragraph(_name)
    BUILTIN[_key] = _st
    STYLE_NAMES[_key] = _st.name

BUILTIN_IDS = {
    "body": "Normal",
    "header": "Header",
    "footer": "Footer",
    "caption": "Caption",
    "tocLevel1": "TOC1",
    "tocLevel2": "TOC2",
    "tocLevel3": "TOC3",
    "footnoteText": "FootnoteText",
    "chapterTitle": "Heading1",
    "sectionTitle1": "Heading2",
    "sectionTitle2": "Heading3",
    "sectionTitle3": "Heading4",
}
_builtin_name_map = {"body": "normal", "header": "header", "footer": "footer",
                     "caption": "caption", "footnoteText": "footnote text",
                     "chapterTitle": "heading 1", "sectionTitle1": "heading 2",
                     "sectionTitle2": "heading 3", "sectionTitle3": "heading 4"}
_by_id = {s.style_id: s for s in doc.styles}
_by_name = {s.name.casefold(): s for s in doc.styles}
for _key, _sid in BUILTIN_IDS.items():
    st = _by_id.get(_sid)
    if st is None and _key in _builtin_name_map:
        st = _by_name.get(_builtin_name_map[_key])
    if st is not None:
        BUILTIN[_key] = st
        STYLE_NAMES[_key] = st.name


def _set_first_line_zero(style):
    """把样式的首行缩进清零。"""
    pPr = style.element.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")


order = [key for key in sorted(STYLE_NAMES, key=_style_sort_key)
         if key not in CHARACTER_STYLE_KEYS and key not in TABLE_STYLE_KEYS]
for key in order:
    if key in BUILTIN:
        st = BUILTIN[key]
    else:
        name = STYLE_NAMES[key]
        if name in [s.name for s in doc.styles]:
            continue
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    BASE_OVERRIDE = {
        "committeeSectionTitle": "sectionTitle1",
        "authorizationTitle": "chapterTitle",
    }
    parent = BASE_OVERRIDE.get(key, FALLBACK.get(key))
    if parent == "body":
        parent = None
    if parent and parent != key and parent in STYLE_NAMES:
        st.base_style = doc.styles[STYLE_NAMES[parent]]
    else:
        st.base_style = None
    spec = RULES["styles"].get(key, {})
    apply_character(st, spec.get("characterFormat"))
    apply_paragraph(st, spec.get("paragraphFormat"))

# 字符样式不是段落：在段落循环之外创建，避免它们带上 pPr
# （首行缩进、分页等）。
for key in sorted(CHARACTER_STYLE_KEYS):
    name = STYLE_NAMES[key]
    try:
        st = doc.styles[name]
    except KeyError:
        # Hyperlink / footnote reference 在 python-docx 的最小模板里是
        # latent 样式，必须按内置样式建，Word 才会把域生成的链接套上去。
        st = doc.styles.add_style(name, WD_STYLE_TYPE.CHARACTER, builtin=True)
    st.base_style = None
    # styleId 不在这里指定：Word 打开并保存后会按内部规则重排它，
    # 写入任何固定值都存不下来。内置样式按 name + builtin=True 建，
    # Word 会用正确的内部名识别。
    spec = RULES["styles"].get(key, {})
    apply_character(st, spec.get("characterFormat"))
    if key in ("subtleEmphasis", "emphasis", "strong"):
        st.font.color.rgb = RGBColor(0, 0, 0)


HEADING_KEYS = ("chapterTitle", "chapterTitleUnnumbered", "chapterTitleNoOutline",
                "sectionTitle1", "sectionTitle2", "sectionTitle3",
                "sectionTitleUnnumbered",
                "committeeSectionTitle", "appendixTitle",
                "appendixSectionTitle1", "appendixSectionTitle2",
                "appendixSectionTitle3")


def _force_style_font(style, cjk, western, bold=None):
    """强制写入样式的中西文字体与字重。"""
    style.font.name = western
    if bold is not None:
        style.font.bold = bold
    style.font.italic = False
    style.font.color.rgb = RGBColor(0, 0, 0)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        if rFonts.get(qn("w:" + attr)) is not None:
            del rFonts.attrib[qn("w:" + attr)]
    rFonts.set(qn("w:ascii"), western)
    rFonts.set(qn("w:hAnsi"), western)
    rFonts.set(qn("w:cs"), western)
    rFonts.set(qn("w:eastAsia"), cjk)
    for color in rPr.findall(qn("w:color")):
        for attr in ("themeColor", "themeShade", "themeTint"):
            if color.get(qn("w:" + attr)) is not None:
                del color.attrib[qn("w:" + attr)]


# 内置强调样式自带主题字体/斜体，显式固定在规则要求的字体和字形上。
for _key, _cjk, _west, _bold in (
    ("subtleEmphasis", "SimHei", "Arial", False),
    ("emphasis", "SimSun", "Times New Roman", True),
    ("strong", "SimHei", "Arial", True),
):
    _force_style_font(doc.styles[STYLE_NAMES[_key]], _cjk, _west, _bold)


# 基础字体来自文档默认值；标题 1 用黑体/Arial 覆盖，其余下级标题
# 继承标题 1，而不各自声明字体。
_force_style_font(doc.styles[STYLE_NAMES["chapterTitle"]], "SimHei", "Arial", bold=False)
for _key in HEADING_KEYS:
    if _key == "chapterTitle":
        continue
    _rPr = doc.styles[STYLE_NAMES[_key]].element.get_or_add_rPr()
    for _tag in ("w:rFonts", "w:color", "w:b", "w:bCs", "w:i", "w:iCs"):
        for _e in _rPr.findall(qn(_tag)):
            _rPr.remove(_e)

# 正文（Normal）只保留段落格式；字体来自 docDefaults。
_normal_rPr = doc.styles["Normal"].element.find(qn("w:rPr"))
if _normal_rPr is not None:
    for _tag in ("w:rFonts", "w:sz", "w:szCs"):
        for _e in _normal_rPr.findall(qn(_tag)):
            _normal_rPr.remove(_e)

# 委员会子标题与授权书标题不纳入目录
for _key in ("committeeSectionTitle", "authorizationTitle"):
    _pPr = doc.styles[STYLE_NAMES[_key]].element.get_or_add_pPr()
    for _e in _pPr.findall(qn("w:outlineLvl")):
        _pPr.remove(_e)
    _pPr.append(_el("w:outlineLvl", val=9))

# 只有正文保留 2 字符首行缩进，其余样式一律清零，
# 否则居中的题注/页眉/目录会被继承的缩进推向右侧。
for _key, _name in STYLE_NAMES.items():
    if _key in CHARACTER_STYLE_KEYS or _key in TABLE_STYLE_KEYS:
        continue
    _pf = (RULES["styles"].get(_key, {}) or {}).get("paragraphFormat") or {}
    if _key != "body" and "firstLineIndent" not in _pf and "hangingIndent" not in _pf:
        _set_first_line_zero(doc.styles[_name])

# 目录行距
for _key in ("tocLevel1", "tocLevel2", "tocLevel3"):
    _st = doc.styles[STYLE_NAMES[_key]]
    _pf = _st.paragraph_format
    _pf.line_spacing = Pt(20)
    _pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    _pf.space_after = Pt(0)

# Word 用 TOC 4 排目录域的第 4 级条目。图表清单要求宋体小四，
# 与主目录里黑体的章标题不同，因此由 \t ...,4 生成并映射到该样式。
try:
    _toc4 = doc.styles["toc 4"]
except KeyError:
    _toc4 = doc.styles.add_style("toc 4", WD_STYLE_TYPE.PARAGRAPH, builtin=True)
    _toc4.element.set(qn("w:styleId"), "TOC4")
_toc4.base_style = None
_toc4.font.name = "Times New Roman"
_toc4.font.size = Pt(12)
_toc4.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
_toc4.paragraph_format.line_spacing = Pt(20)
_toc4.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
_toc4.paragraph_format.space_before = Pt(0)
_toc4.paragraph_format.space_after = Pt(0)
_toc4.paragraph_format.left_indent = Pt(0)
_rpr = _toc4.element.get_or_add_rPr()
_rfonts = _rpr.find(qn("w:rFonts"))
if _rfonts is None:
    _rfonts = OxmlElement("w:rFonts")
    _rpr.insert(0, _rfonts)
_rfonts.set(qn("w:ascii"), "Times New Roman")
_rfonts.set(qn("w:hAnsi"), "Times New Roman")
_rfonts.set(qn("w:eastAsia"), "SimSun")
# 制表位不在这里硬写：listOfFiguresEntry 的 rules.tabs 已由主循环套到 toc 4，
# 这里再补一次会出现两条同位置同类型的 w:tab。

# 绑定编号
num = RULES["styles"]
for key, attr in (("chapterTitle", ("outline", 1)),
                  ("sectionTitle1", ("outline", 2)),
                  ("sectionTitle2", ("outline", 3)),
                  ("sectionTitle3", ("outline", 4)),
                  ("appendixTitle", ("appendix", 1)),
                  ("appendixSectionTitle1", ("appendix", 2)),
                  ("appendixSectionTitle2", ("appendix", 3)),
                  ("appendixSectionTitle3", ("appendix", 4))):
    set_style_numbering(doc.styles[STYLE_NAMES[key]], NUM[attr[0]][1], attr[1])
for key in ("chapterTitleUnnumbered", "chapterTitleNoOutline",
             "sectionTitleUnnumbered"):
    disable_style_numbering(doc.styles[STYLE_NAMES[key]])
set_style_numbering(doc.styles[STYLE_NAMES["referenceEntry"]],
                    NUM["references"][1], 1)
_ref_ay_style = doc.styles[STYLE_NAMES["referenceEntryAuthorYear"]]
apply_character(_ref_ay_style,
                RULES["styles"]["referenceEntryAuthorYear"].get("characterFormat"))
apply_paragraph(_ref_ay_style,
                RULES["styles"]["referenceEntryAuthorYear"].get("paragraphFormat"))
disable_style_numbering(_ref_ay_style)

# 参考文献条目在内置母版中已存在，主循环会跳过；这里强制套用规则中的
# 段前 3 磅 / 段后 0 磅，避免规则改了但生成物不更新。
_ref_style = doc.styles[STYLE_NAMES["referenceEntry"]]
apply_character(_ref_style, RULES["styles"]["referenceEntry"].get("characterFormat"))
apply_paragraph(_ref_style, RULES["styles"]["referenceEntry"].get("paragraphFormat"))
set_style_numbering(doc.styles[STYLE_NAMES["blockCode"]],
                    NUM["codeLines"][1], 1)
for _key in ("listNumber1", "listNumber2", "listNumber3",
             "listBullet1", "listBullet2", "listBullet3"):
    set_style_numbering(doc.styles[STYLE_NAMES[_key]], NUM[_key][1], 1)
    # 内置列表样式在 docx 模板中已存在，主循环的“已存在则 continue”会跳过它们；
    # 这里强制套用与正文一致的字体/字号/行距/缩进。
    _st = doc.styles[STYLE_NAMES[_key]]
    _spec = RULES["styles"].get(_key, {})
    apply_character(_st, _spec.get("characterFormat"))
    apply_paragraph(_st, _spec.get("paragraphFormat"))

# 委员会子标题不能继承章标题的分页符
_cs = doc.styles[STYLE_NAMES["committeeSectionTitle"]]
_cs.paragraph_format.page_break_before = False
_cs.paragraph_format.keep_with_next = True

# 章级标题不再使用“段前分页”；每个章级部分由显式分节符（下一页）起始。
for _key, _name in STYLE_NAMES.items():
    if _key in CHARACTER_STYLE_KEYS or _key in TABLE_STYLE_KEYS:
        continue
    doc.styles[_name].paragraph_format.page_break_before = False


def get_style(name):
    """把语义样式键解析为 Word 样式对象。"""
    return doc.styles[STYLE_NAMES.get(name, name)]


def P(text="", style="body"):
    """按给定样式追加一个正文段落。"""
    p = doc.add_paragraph(style=get_style(style))
    if text:
        p.add_run(text)
    return p


def ref_entry(text, entry_id):
    """追加一条参考文献条目。"""
    p = doc.add_paragraph(style=get_style("referenceEntry"))
    bid = add_bookmark_start(p, f"THU_REF_ENTRY_{entry_id:04}")
    p.add_run(text)
    add_bookmark_end(p, bid)
    return p


def add_citation(paragraph, *entry_ids):
    """写入与文献引用宏相同的域，保留普通 F9 下的编号结果。"""
    ids = ";".join(f"{entry_id:04}" for entry_id in entry_ids)
    result = "[" + ", ".join(map(str, entry_ids)) + "]"
    add_field(paragraph, f'QUOTE "THU_CITE:{ids}"', result,
              locked=True, char_style=get_style("citation"))

def H(text, style):
    """按给定标题样式追加一个标题。"""
    p = doc.add_paragraph(style=get_style(style))
    p.add_run(text)
    return p


def clear_indent(p):
    """清除段落的首行缩进。"""
    pPr = p._p.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:leftChars"), "0")
    ind.set(qn("w:left"), "0")


def cell_text(cell, text, size=11, bold=False):
    """按给定字号与字重把文字写入表格单元格。"""
    p = cell.paragraphs[0]
    p.style = get_style("tableText")
    r = p.add_run(text)
    if bold:
        set_char_style(r, "strong")


def repeat_header(table):
    """把表格首行标记为重复标题行。"""
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(_el("w:tblHeader", val="true"))


def _border_el(tag, spec, default_sz=8):
    """按规则写一条框线：{style: single|none, width: {value, unit}}。"""
    e = OxmlElement("w:" + tag)
    style = spec.get("style", "none")
    e.set(qn("w:val"), "nil" if style == "none" else style)
    width = spec.get("width")
    e.set(qn("w:sz"), str(int(round(width["value"] * 8)) if width else default_sz))
    e.set(qn("w:space"), "0")
    e.set(qn("w:color"), "auto")
    return e


_TBL_EDGE_TAGS = ("top", "left", "bottom", "right", "insideH", "insideV")


def add_table_style(key="threeLineTable"):
    """建表格样式（Word: w:type="table"），框线/内边距全部来自规则。"""
    spec = RULES["styles"][key]
    fmt = spec.get("tableFormat") or {}
    st = doc.styles.add_style(STYLE_NAMES[key], WD_STYLE_TYPE.TABLE)
    # styleId 由 Word 在保存时决定，这里不写死（理由同字符样式处）。
    el = st.element
    tblPr = OxmlElement("w:tblPr")
    align = fmt.get("alignment")
    if align:
        # w:tblPr 的子元素顺序：… tblW → jc → tblBorders → tblCellMar …
        tblPr.append(_el("w:jc", val={"center": "center", "left": "left",
                                      "right": "right"}[align]))
    borders_spec = fmt.get("tableBorders") or {}
    borders = OxmlElement("w:tblBorders")
    for edge in _TBL_EDGE_TAGS:
        borders.append(_border_el(edge, borders_spec.get(edge) or {"style": "none"},
                                  default_sz=12))
    tblPr.append(borders)
    margins = fmt.get("cellMargins") or {}
    if margins:
        mar = OxmlElement("w:tblCellMar")
        for edge in ("top", "left", "bottom", "right"):
            item = margins.get(edge)
            if not item:
                continue
            e = OxmlElement("w:" + edge)
            e.set(qn("w:w"), str(int(round(item["value"] * 20))))
            e.set(qn("w:type"), "dxa")
            mar.append(e)
        tblPr.append(mar)
    el.append(tblPr)
    first = OxmlElement("w:tblStylePr")
    first.set(qn("w:type"), "firstRow")
    tcPr = OxmlElement("w:tcPr")
    tcB = OxmlElement("w:tcBorders")
    for edge, edge_spec in (fmt.get("firstRowBorders") or {}).items():
        tcB.append(_border_el(edge, edge_spec))
    if len(tcB):
        tcPr.append(tcB)
        first.append(tcPr)
        el.append(first)
    return st


TABLE_STYLE = add_table_style()


def apply_table_text(table):
    """给单元格套用"表格文字"段落样式（无段前段后间距）。"""
    style = get_style("tableText")
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.style = style


def three_line_table(rows, cols):
    """用三线表样式建表。"""
    t = doc.add_table(rows=rows, cols=cols)
    t.style = TABLE_STYLE
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    repeat_header(t)
    apply_table_text(t)
    # 表题已与表头同页；表头还须与首条数据同页，避免只留下空表头。
    for cell in t.rows[0].cells:
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.keep_with_next = True
    # 行内不拆分；不设 keepNext，让表格紧接前文并在页间断开。
    for row in t.rows:
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
    return t


def set_table_widths(table, widths_cm):
    """按厘米固定各列宽度，避免默认等分把长文字挤成窄条。"""
    table.autofit = False
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for i, width in enumerate(widths_cm):
        table.columns[i].width = Cm(width)
        for cell in table.columns[i].cells:
            cell.width = Cm(width)


def add_seq_reset(ids=("表", "图", "式", "附录表", "附录图"), paragraph=None):
    """给段落追加空的 SEQ 重置域（无隐藏文字、不占额外行）。

    章级题注已用 ``SEQ ... \s 1``；附录标题需要显式重置，
    因为 ``\s 1`` 只在标题 1 处重启。
    """
    p = paragraph if paragraph is not None else doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    for ident in ids:
        add_field(p, r"SEQ %s \r 0 \h" % ident, "")
    return p


def _math_rpr():
    """OMML 数学 run 共用的 run 属性：行内公式字符样式。"""
    style_id = _xml_escape(doc.styles[STYLE_NAMES["inlineFormula"]].style_id)
    return '<w:rPr><w:rStyle w:val="%s"/></w:rPr>' % style_id


def add_formula_omml(p, block=True, number=None):
    """追加一段各 run 都带行内公式字符样式的 OMML 公式。"""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    rpr = _math_rpr()
    numbered = (('<m:r>' + rpr + '<m:t>#(%s)</m:t></m:r>') % number
                if number else '')
    math = (
        '<m:oMath>'
        + '<m:r>' + rpr + '<m:t>E = m</m:t></m:r>'
        + '<m:sSup><m:e><m:r>' + rpr + '<m:t>c</m:t></m:r></m:e>'
        + '<m:sup><m:r>' + rpr + '<m:t>2</m:t></m:r></m:sup></m:sSup>'
        + numbered
        + '</m:oMath>'
    )
    if block:
        xml = ('<m:oMathPara %s><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
               + math + '</m:oMathPara>') % nsdecls("m", "w")
    else:
        xml = ('<m:oMath %s>' % nsdecls("m", "w")) + math[len('<m:oMath>'):]
    p._p.append(parse_xml(xml))


def add_inline_formula(p, text):
    """插入一个带"行内公式"字符样式的行内 OMath 对象。"""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    xml = ('<m:oMath %s><m:r>' % nsdecls("m", "w")) + _math_rpr() \
          + '<m:t>' + _xml_escape(text) + '</m:t></m:r></m:oMath>'
    p._p.append(parse_xml(xml))
    return p


def add_caption(style, label, chapter_style, seq_id, title):
    """写入题注并给编号加书签，供 REF 域引用。"""
    p = doc.add_paragraph(style=get_style(style))
    gaps = (RULES.get("captions") or {}).get("defaults") or {}
    bookmark, number = peek_caption(seq_id)
    p.add_run(label + gaps.get("labelGap", ""))
    ref = chapter_style if chapter_style.isdigit() else '"%s"' % chapter_style
    bid = add_bookmark_start(p, bookmark)
    add_field(p, 'STYLEREF %s \\s' % ref, "1")
    p.add_run(".")
    add_field(p, "SEQ %s \\* ARABIC \\s 1" % seq_id, "1")
    add_bookmark_end(p, bid)
    p.add_run(gaps.get("titleGap", "　") + title)
    if style in ("tableCaption", "appendixTableCaption"):
        p.paragraph_format.keep_with_next = True
    p.caption_bookmark = bookmark
    p.caption_number = number
    return p


# ---------------------------------------------------------------------------
# 封面辅助函数：内容控件、浮动表格、页面边框
# ---------------------------------------------------------------------------
from docx.enum.table import WD_ALIGN_VERTICAL
from xml.sax.saxutils import escape as _xml_escape

_SDT_ID = [7000]
_FIT_ID = [-350900000]


def _set_fit_text(run, width_twips=1280):
    """把短标签压缩/拉伸到固定字符宽度。"""
    fit = OxmlElement("w:fitText")
    fit.set(qn("w:val"), str(width_twips))
    fit.set(qn("w:id"), str(_FIT_ID[0]))
    _FIT_ID[0] -= 1
    run._r.get_or_add_rPr().append(fit)
    return run


def set_char_style(run, key):
    """套用某个可复用字符样式，不加直接格式。"""
    run.style = doc.styles[STYLE_NAMES[key]]
    return run


def add_svg_picture(paragraph, png_path, svg_path, width, height):
    """插入矢量形态为 SVG 徽标、并带 PNG 回退的图片。

    python-docx 不原生支持 SVG，因此照常插入位图，再在
    ``<a:blip>`` 上挂 ``<asvg:svgBlip>`` 扩展。Word 2016+ 渲染
    矢量 SVG，较旧的阅读器回退到 PNG。SVG 作为独立的包部件存储，
    内容类型为 ``image/svg+xml``，通过标准图片关系关联。
    """
    from docx.oxml import parse_xml
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    run = paragraph.add_run()
    shape = run.add_picture(png_path, width=width, height=height)
    blip = shape._inline.find(".//" + qn("a:blip"))
    svg_bytes = open(svg_path, "rb").read()
    partname = PackURI("/word/media/" + os.path.basename(svg_path))
    svg_part = Part(partname, "image/svg+xml", svg_bytes, doc.part.package)
    rId = doc.part.relate_to(svg_part, RT.IMAGE)
    extLst = OxmlElement("a:extLst")
    ext = OxmlElement("a:ext")
    ext.set("uri", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
    svgBlip = parse_xml(
        '<asvg:svgBlip '
        'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'r:embed="%s"/>' % rId)
    ext.append(svgBlip)
    extLst.append(ext)
    blip.append(extLst)
    return shape


_DRAWING_ID = [9000]


def add_placeholder_shape(paragraph, width, height, label="示例插图"):
    """插入原生 DrawingML 插图占位符（带标签的虚线框）。

    形状完全存在于文档 XML 里：没有图片部件、没有外部资源，也不依赖
    图像处理库，因此干净检出后示例章节仍可构建。它在 Word 里保持可编辑：
    读者可删除它再插入真实图片。宽高为 EMU 值。
    """
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    cx = int(width)
    cy = int(height)
    _DRAWING_ID[0] += 1
    xml = (
        '<w:r __NS__>'
        '<w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%(cx)d" cy="%(cy)d"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%(id)d" name="%(name)s"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic>'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:cNvSpPr/>'
        '<wps:spPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="%(cx)d" cy="%(cy)d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill>'
        '<a:prstDash val="dash"/></a:ln>'
        '</wps:spPr>'
        '<wps:txbx><w:txbxContent>'
        '<w:p><w:pPr><w:pStyle w:val="%(style)s"/>'
        '<w:spacing w:before="0" w:after="0"/>'
        '<w:ind w:left="0" w:right="0" w:firstLine="0" w:firstLineChars="0"/>'
        '<w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:color w:val="808080"/></w:rPr>'
        '<w:t xml:space="preserve">%(label)s</w:t></w:r></w:p>'
        '</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0"'
        ' bIns="0" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        '</wps:wsp>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:inline>'
        '</w:drawing>'
        '</w:r>'
    ) % {
        "cx": cx,
        "cy": cy,
        "id": _DRAWING_ID[0],
        "name": _xml_escape(label),
        "label": _xml_escape(label),
        "style": get_style("caption").style_id,
    }
    paragraph._p.append(parse_xml(xml.replace("__NS__", nsdecls("w", "a", "wp"))))
    return paragraph


def _cc_run_xml(text, rpr):
    """内容控件里的一个 run；保留开头制表符与行内换行。"""
    tabs = len(text) - len(text.lstrip("\t"))
    body = '</w:t><w:br/><w:t xml:space="preserve">'.join(
        _xml_escape(line) for line in text[tabs:].split("\n"))
    return '<w:r>%s%s<w:t xml:space="preserve">%s</w:t></w:r>' % (
        rpr, "<w:tab/>" * tabs, body)


def _cc_list_items_xml(items):
    """为下拉或组合框生成 w:listItem 子元素。"""
    pieces = []
    for item in items:
        if isinstance(item, (tuple, list)) and len(item) == 2:
            display, value = item
        else:
            display = value = item
        pieces.append('<w:listItem w:displayText="%s" w:value="%s"/>'
                      % (_xml_escape(str(display)), _xml_escape(str(value))))
    return "".join(pieces)


def add_content_control(paragraph, placeholder, alias, kind="text", bold=False,
                        font=None, size=None, parts=None,
                        showing_placeholder=True, char_style=None, items=None):
    """给段落追加 Word 内容控件（纯文本、富文本、下拉、组合框）。

    控件保存易读的占位文字，并标记 w:showingPlcHdr，使 Word 在读者
    输入前显示提示。``parts`` 在一个控件里渲染多个 run；某个 part 开头的
    制表符会转成真正的 <w:tab/>，供英文导师行使用。``kind`` 可取
    "text"、"richText"、"dropdownList" 或 "comboBox"；两种列表通过
    ``items``（字符串或 (display, value) 对）接收选项，未给占位文字时
    显示第一项。
    """
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    _SDT_ID[0] += 1
    if kind == "dropdownList":
        typ = "<w:dropDownList>%s</w:dropDownList>" % _cc_list_items_xml(items or [])
    elif kind == "comboBox":
        typ = "<w:comboBox>%s</w:comboBox>" % _cc_list_items_xml(items or [])
    elif kind == "richText":
        typ = "<w:richText/>"
    else:
        typ = "<w:text/>"
    rpr_parts = []
    if char_style:
        rpr_parts.append('<w:rStyle w:val="%s"/>'
                         % _xml_escape(doc.styles[STYLE_NAMES[char_style]].style_id))
    if font:
        rpr_parts.append(
            '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="SimSun"/>'
            % (_xml_escape(font), _xml_escape(font), _xml_escape(font))
        )
    if bold:
        rpr_parts.append("<w:b/>")
    if size:
        rpr_parts.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>'
                         % (int(round(size * 2)), int(round(size * 2))))
    rpr = "<w:rPr>%s</w:rPr>" % "".join(rpr_parts) if rpr_parts else ""
    if parts is not None:
        pieces = parts
    elif placeholder:
        pieces = [placeholder]
    else:
        pieces = []
    content = "".join(_cc_run_xml(piece, rpr) for piece in pieces)
    xml = (
        '<w:sdt %s><w:sdtPr>'
        '<w:alias w:val="%s"/><w:tag w:val="%s"/><w:id w:val="%d"/>%s%s'
        '</w:sdtPr>'
        '<w:sdtContent>%s</w:sdtContent>'
        '</w:sdt>'
        % (nsdecls("w"), _xml_escape(alias), _xml_escape(alias), _SDT_ID[0],
           typ, '<w:showingPlcHdr/>' if showing_placeholder else '', content)
    )
    paragraph._p.append(parse_xml(xml))
    return paragraph


def _exact_line(p, points):
    """把段落行距设为固定磅值。"""
    pf = p.paragraph_format
    pf.line_spacing = Pt(points)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY


# --- section 1: 中文封面（无框线表格固定各部分高度）----------------------
def _cover_table_properties(table, width_pt):
    """设置单列封面表格：固定宽度、无边框、无单元格边距。"""
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.insert(0, tblW)
    tblW.set(qn("w:w"), str(int(round(width_pt * 20))))
    tblW.set(qn("w:type"), "dxa")

    for tag in ("w:tblBorders", "w:tblLayout", "w:tblCellMar"):
        for old in tblPr.findall(qn(tag)):
            tblPr.remove(old)

    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:%s" % edge)
        el.set(qn("w:val"), "nil")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)

    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")

    margins = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:%s" % edge)
        el.set(qn("w:w"), "0")
        el.set(qn("w:type"), "dxa")
        margins.append(el)

    for el in (borders, layout, margins):
        if tblPr.find(qn("w:tblLook")) is not None:
            tblPr.insert_element_before(el, "w:tblLook")
        else:
            tblPr.append(el)


def add_cover_layout_table(row_heights, width_pt):
    """创建固定高度、无边框的封面版式表格。"""
    table = doc.add_table(rows=len(row_heights), cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Pt(width_pt)
    _cover_table_properties(table, width_pt)
    for row, height in zip(table.rows, row_heights):
        row.height = Pt(height)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        cell = row.cells[0]
        cell.width = Pt(width_pt)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    return table


def merge_spacer_rows(table, spacer_row_indexes):
    """把空白占位行合并进上方单元格（不留空行单元格）。"""
    for idx in sorted(spacer_row_indexes, reverse=True):
        table.cell(idx - 1, 0).merge(table.cell(idx, 0))


def format_cover_cell_paragraph(p, style, align=WD_ALIGN_PARAGRAPH.LEFT,
                                line_pt=None, before=0, after=0):
    """设置封面单元格段落的样式、对齐与间距。"""
    p.style = get_style(style)
    p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    clear_indent(p)
    if line_pt is not None:
        _exact_line(p, line_pt)
    return p


def cover_cell_paragraph(cell, style, align=WD_ALIGN_PARAGRAPH.LEFT,
                         line_pt=None, before=0, after=0):
    """设置封面表格单元格的首段格式。"""
    return format_cover_cell_paragraph(
        cell.paragraphs[0], style, align=align, line_pt=line_pt,
        before=before, after=after)


def add_cover_security_header(sec):
    """密级只占中文封面的独立页眉，公开论文默认留空。"""
    sec.header.is_linked_to_previous = False
    p = clear_header_footer(sec.header)
    format_cover_cell_paragraph(p, "coverSecurity")
    add_content_control(p, "", "密级", showing_placeholder=False)


def add_cover_info_paragraph(p, label, placeholder, alias, kind="text", items=None,
                             showing_placeholder=True):
    """写入带标签与内容控件的封面信息行。"""
    p.style = get_style("coverAuthorInfo")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    clear_indent(p)
    _exact_line(p, 26)
    p.paragraph_format.tab_stops.add_tab_stop(
        Pt(159), WD_TAB_ALIGNMENT.RIGHT)
    p.paragraph_format.tab_stops.add_tab_stop(
        Pt(164), WD_TAB_ALIGNMENT.LEFT)
    p.add_run("	")
    _set_fit_text(p.add_run(label))
    p.add_run("：")
    p.add_run("	")
    add_content_control(p, placeholder, alias, kind=kind, items=items,
                        showing_placeholder=showing_placeholder)


import build.content.examples as examples

_ctx = sys.modules[__name__]
_pkg = VARIANT.content_package
cover = importlib.import_module(_pkg + ".cover")
front_matter = importlib.import_module(_pkg + ".front_matter")
back_matter = importlib.import_module(_pkg + ".back_matter")

cover.build(_ctx)
front_matter.build(_ctx)
examples.build(_ctx)
back_matter.build(_ctx)

apply_document_text_flow(doc, RULES, IMAP)
if not args.defer_style_presentation:
    apply_style_presentation(doc.styles.element, RULES["styles"], IMAP)
doc.save(OUT)
sanitize_office_metadata(OUT)
print("saved:", OUT)
