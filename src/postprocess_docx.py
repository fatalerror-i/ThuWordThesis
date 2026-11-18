# -*- coding: utf-8 -*-
"""在 Word 更新域之后后处理生成的模板。

Word 重建目录条目时会带上内置目录样式的直接格式，
并给内置"题注"样式套上默认的加粗蓝色外观。
本脚本把这些样式/段落规范到论文格式。
"""
import argparse
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
import yaml

from build.variants import choices, resolve
from build.hyphenation import apply_document_text_flow

parser = argparse.ArgumentParser(description="规范 Word 更新域后的 DOCX 格式。")
parser.add_argument("--variant", required=True, choices=choices())
parser.add_argument("--docx", type=Path, help="待处理的 DOCX 路径，默认取变体配置")
args = parser.parse_args()
VARIANT = resolve(args.variant)
DOCX = args.docx or VARIANT.docx
# 规则文件是唯一事实来源：这里要用到页眉下框线、清单条目制表位等取值。
with open(VARIANT.rules, encoding="utf-8") as _f:
    RULES = yaml.safe_load(_f)


def _style_pf(key):
    """取某个样式键的规则段落格式（拿不到就返回空字典）。"""
    return (((RULES.get("styles") or {}).get(key) or {}).get("paragraphFormat") or {})


def set_fonts(style, chinese, western):
    """写入样式的中西文字体与东亚宽度提示。"""
    style.font.name = western
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), western)
    rFonts.set(qn("w:hAnsi"), western)
    rFonts.set(qn("w:cs"), western)
    rFonts.set(qn("w:eastAsia"), chinese)
    # 全角引号、破折号等东亚宽度字符按中文字体渲染。
    rFonts.set(qn("w:hint"), "eastAsia")


def set_indent(style, left_chars, first_line_chars=0):
    """按字符数写入样式的左缩进与首行缩进。"""
    pPr = style.element.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    ind.set(qn("w:leftChars"), str(left_chars))
    ind.set(qn("w:left"), str(int(left_chars * 2.4)))  # 12 磅 -> 每字符 240 缇
    ind.set(qn("w:firstLineChars"), str(first_line_chars))
    ind.set(qn("w:firstLine"), "0")


def clear_direct_run_format(p):
    """清除 run 与段落标记上的直接 rPr，让样式生效。"""
    pPr = p._p.find(qn("w:pPr"))
    if pPr is not None:
        for r in pPr.findall(qn("w:rPr")):
            pPr.remove(r)
    for run in p.runs:
        rPr = run._r.find(qn("w:rPr"))
        if rPr is not None:
            run._r.remove(rPr)


d = Document(DOCX)

# TOC levels: chapter entries 小四黑体, others 小四宋体, fixed 20 pt, 0/1/2 chars
toc_specs = (("toc 1", "SimHei", 0), ("toc 2", "SimSun", 100), ("toc 3", "SimSun", 200), ("toc 4", "SimSun", 0))
for name, cjk, left in toc_specs:
    st = d.styles[name]
    set_fonts(st, cjk, "Arial" if cjk == "SimHei" else "Times New Roman")
    st.font.size = Pt(12)
    pf = st.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing = Pt(20)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    set_indent(st, left)

for p in d.paragraphs:
    if p.style.name.lower() in ("toc 1", "toc 2", "toc 3", "toc 4"):
        clear_direct_run_format(p)
        pPr = p._p.get_or_add_pPr()
        for ind in pPr.findall(qn("w:ind")):
            pPr.remove(ind)

# 两种变体都用 \c "图" / \c "表" 图表目录域。Word 会用它内置的
# "图表目录"样式，带 2 字符左缩进、10 磅段后 / 1.15 倍行距，
# 并给每条生成的条目盖上直接的 960/480 缇缩进。指南 2.3.7 要求
# 条目为小四宋体、固定 20 磅行距、段前段后 0 磅、无缩进（本科
# 官方样例通过基于 TOC 1 实现）。这里同时规范样式与条目段落，
# 且必须在每次更新域之后运行（Word 会重新加缩进），并按 styleId
# 匹配，避免内置样式名的翻译让清理漏掉条目段落。
_tof = None
for _st in d.styles:
    if _st.type == 1 and (_st.name or "").lower() == "table of figures":
        _tof = _st
        break
if _tof is not None:
    _tof_id = _tof.style_id
    _tof.base_style = d.styles["toc 1"]
    # 指南 2.3.7：宋体小四、固定 20 磅、段前段后 0 磅。内置样式默认是段后
    # 10 磅 + 1.15 倍行距，必须显式覆盖（Word 在 SaveAs2 会剥掉 basedOn）。
    set_fonts(_tof, "SimSun", "Times New Roman")
    _tof.font.size = Pt(12)
    _tof_pf = _tof.paragraph_format
    _tof_pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _tof_pf.line_spacing = Pt(20)
    _tof_pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    _tof_pf.space_before = Pt(0)
    _tof_pf.space_after = Pt(0)
    _tof_pPr = _tof.element.get_or_add_pPr()
    for _ind in _tof_pPr.findall(qn("w:ind")):
        _tof_pPr.remove(_ind)
    # 意见 8：清单条目的页码右对齐到版心右缘（15 cm、前导点），与 toc 1 一致。
    _tof_tabs = _tof_pPr.findall(qn("w:tabs"))
    for _t in _tof_tabs:
        _tof_pPr.remove(_t)
    for _tab in _style_pf("tableOfFigures").get("tabs") or []:
        _tabs_el = _tof_pPr.find(qn("w:tabs"))
        if _tabs_el is None:
            _tabs_el = OxmlElement("w:tabs")
            # w:pPr 子元素顺序：… pBdr → shd → tabs → … → spacing → ind → jc …
            _anchor = None
            for _tag in ("w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
                         "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
                         "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
                         "w:snapToGrid", "w:spacing", "w:ind", "w:jc",
                         "w:outlineLvl", "w:rPr"):
                _anchor = _tof_pPr.find(qn(_tag))
                if _anchor is not None:
                    break
            if _anchor is not None:
                _anchor.addprevious(_tabs_el)
            else:
                _tof_pPr.append(_tabs_el)
        _tab_el = OxmlElement("w:tab")
        _tab_el.set(qn("w:val"), {"right": "right", "left": "left",
                                  "center": "center"}[_tab["alignment"]])
        _tab_el.set(qn("w:leader"), {"dots": "dot", "none": "none"}.get(
            _tab.get("leader", "none"), "none"))
        _tab_el.set(qn("w:pos"), str(int(round(_tab["position"]["value"] * 566.929))))
        _tabs_el.append(_tab_el)
    for p in d.paragraphs:
        _pPr = p._p.find(qn("w:pPr"))
        if _pPr is None:
            continue
        _pStyle = _pPr.find(qn("w:pStyle"))
        if _pStyle is None or _pStyle.get(qn("w:val")) != _tof_id:
            continue
        clear_direct_run_format(p)
        for ind in _pPr.findall(qn("w:ind")):
            _pPr.remove(ind)
        # 域更新可能给结果段落盖直接 w:spacing；清掉，让样式层生效。
        for sp in _pPr.findall(qn("w:spacing")):
            _pPr.remove(sp)

# 题注：黑色、11 磅、居中、无继承缩进
cap = d.styles["Caption"]
cap.font.bold = False
cap.font.color.rgb = RGBColor(0, 0, 0)
cap.font.size = Pt(11)
cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap.paragraph_format.first_line_indent = Pt(0)
pPr = cap.element.get_or_add_pPr()
ind = pPr.get_or_add_ind()
ind.set(qn("w:firstLineChars"), "0")
ind.set(qn("w:firstLine"), "0")

# 页眉 / 页脚：五号宋体 + TNR、居中、无首行缩进
for name in ("Header", "Footer"):
    st = d.styles[name]
    set_fonts(st, "SimSun", "Times New Roman")
    st.font.size = Pt(10.5)
    st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    st.paragraph_format.first_line_indent = Pt(0)


def _clear_tabs(style):
    """清掉样式自带的制表位（意见 2：页眉与页脚一样不设制表位）。"""
    pPr = style.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:tabs")):
        pPr.remove(old)


# 页眉页脚都不设边框和制表位；只有研究生版页眉保留规则里的下边框。
for name in ("Header", "Footer"):
    _clear_tabs(d.styles[name])

# 正文保留 2 字符首行缩进；其余样式一律清零
body = d.styles["Normal"]
body.paragraph_format.first_line_indent = Pt(24)
pPr = body.element.get_or_add_pPr()
ind = pPr.get_or_add_ind()
ind.set(qn("w:firstLineChars"), "200")
ind.set(qn("w:firstLine"), "480")

for st in d.styles:
    if st.type != 1 or st.name == "Normal":
        continue
    if st.name.lower().startswith("toc "):
        continue
    if st.name.startswith(("List Number", "List Bullet")):
        continue
    if st.name == "授权说明文字":
        # 授权说明正文保留规则里的 2 字符首行缩进。
        continue
    pPr = st.element.get_or_add_pPr()
    ind = pPr.get_or_add_ind()
    if ind.get(qn("w:hangingChars")) or ind.get(qn("w:hanging")):
        for attr in ("w:firstLineChars", "w:firstLine"):
            if ind.get(qn(attr)):
                del ind.attrib[qn(attr)]
        continue
    ind.set(qn("w:firstLineChars"), "0")
    ind.set(qn("w:firstLine"), "0")


def _clear_pbdr(pPr):
    """清除段落属性中的段落边框定义。"""
    for old in pPr.findall(qn("w:pBdr")):
        pPr.remove(old)


def _nil_pbdr(pPr):
    """写入四边为空的段落边框定义。"""
    _clear_pbdr(pPr)
    pBdr = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "nil")
        e.set(qn("w:sz"), "0")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        pBdr.append(e)
    pPr.insert(0, pBdr)


def _header_pbdr(pPr):
    """按规则文件写入页眉下框线（单线 0.75 磅、间隔 1 磅）。"""
    border = (_style_pf("header").get("borderBottom")
              or {"width": {"value": 0.75}, "spacing": {"value": 1}})
    _clear_pbdr(pPr)
    pBdr = OxmlElement("w:pBdr")
    for edge in ("top", "left", "right"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "nil")
        e.set(qn("w:sz"), "0")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        pBdr.append(e)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    # sz 以 1/8 磅计：0.75pt → 6。
    bottom.set(qn("w:sz"), str(int(round(border["width"]["value"] * 8))))
    bottom.set(qn("w:space"), str(int(round(border.get("spacing", {}).get("value", 1)))))
    bottom.set(qn("w:color"), "auto")
    pBdr.append(bottom)
    pPr.insert(0, pBdr)


# 意见 2：页眉页脚一律不带框线；只有研究生版页眉按规则加一条下框线。
# 本科版页眉是“无”（headersAndFooters.header.content = none），同样只清不画。
_has_header_border = bool(_style_pf("header").get("borderBottom"))
for st in d.styles:
    name = (st.name or "").lower()
    if name in ("header", "页眉"):
        if VARIANT.decorate_header and _has_header_border:
            _header_pbdr(st.element.get_or_add_pPr())
        else:
            _nil_pbdr(st.element.get_or_add_pPr())
    elif name in ("footer", "页脚"):
        _nil_pbdr(st.element.get_or_add_pPr())
for sec in d.sections:
    for part in (sec.header, sec.even_page_header, sec.first_page_header):
        try:
            paras = part.paragraphs
        except Exception:
            continue
        for p in paras:
            # 只有真正的页眉段落带规则文件里的框线；封面
            # 密级行和空白/普通页眉保持无框线。
            if (VARIANT.decorate_header and _has_header_border
                    and p.style.name.lower() in ("header", "页眉")):
                _header_pbdr(p._p.get_or_add_pPr())
            else:
                _nil_pbdr(p._p.get_or_add_pPr())
    for part in (sec.footer, sec.even_page_footer, sec.first_page_footer):
        try:
            paras = part.paragraphs
        except Exception:
            continue
        for p in paras:
            _nil_pbdr(p._p.get_or_add_pPr())

# 参考文献：1 厘米悬挂缩进（首行顶格，续行 +1 厘米）
ref = d.styles["参考文献条目"]
pPr = ref.element.get_or_add_pPr()
ind = pPr.get_or_add_ind()
for attr in list(ind.attrib):
    del ind.attrib[attr]
# 悬挂缩进 1 厘米（567 twips）：Word 的标准写法是“左缩进 1 厘米 + 悬挂 1 厘米”。
# 不写 hangingChars/leftChars，否则 Word 会优先按字符单位显示，覆盖厘米值。
ind.set(qn("w:left"), "567")
ind.set(qn("w:hanging"), "567")
# Word 会把与 docDefaults 相同的 w:after 从样式中删掉；在 Word 更新域之后
# 再写回一次，保证打开“修改样式”看到的是规则里的 3/0 磅。
sp = pPr.find(qn("w:spacing"))
if sp is None:
    sp = OxmlElement("w:spacing")
    pPr.insert(0, sp)
sp.set(qn("w:before"), "60")
sp.set(qn("w:after"), "0")
for p in d.paragraphs:
    if p.style.name == "参考文献条目":
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(0)
        pPr = p._p.get_or_add_pPr()
        for e in pPr.findall(qn("w:ind")):
            pPr.remove(e)
        ind = pPr.get_or_add_ind()
        ind.set(qn("w:left"), "567")
        ind.set(qn("w:hanging"), "567")

for _name in ("符号说明条目",):
    try:
        _pf = d.styles[_name].paragraph_format
        _pf.line_spacing = Pt(20)
        _pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        _pf.space_before = Pt(0)
        _pf.space_after = Pt(0)
    except KeyError:
        pass

# 每个样式的字体都继承自文档默认段落字体。
# 只有标题/题注/目录这几条链保留父样式；没有样式继承正文（Normal），
# 因此正文的 2 字符首行缩进不会漏进其他样式。
BASE_MAP = {
    "Heading 2": "Heading 1",
    "Heading 3": "Heading 2",
    "Heading 4": "Heading 3",
    "章内无编号小节": "Heading 2",
    "章标题（无编号）": "Heading 1",
    "章标题（无编号、不进目录）": "章标题（无编号）",
    "附录标题": "Heading 1",
    "附录一级节标题": "Heading 2",
    # 附录二/三级节标题继承正文二/三级节标题（13pt / 12pt），与
    # configs/implementation-map.yaml 的 fallback 保持一致；曾经错误地串成
    # 「二级→一级、三级→二级」，导致两者都继承到 14pt。
    "附录二级节标题": "Heading 3",
    "附录三级节标题": "Heading 4",
    "名单页小节标题": "Heading 2",
    "授权说明标题": "章标题（无编号）",
    "图注": "Caption",
    "子图注": "图注",
    "附录图注": "图注",
    "表注": "Caption",
    "附录表注": "表注",
    "toc 2": "toc 1",
    "toc 3": "toc 2",
    "toc 4": "toc 1",
    # 英文封面题目跟随中文封面题目（居中、1.25 倍行距）；
    # Word 会把这两项当作冗余属性先删掉，本脚本随后清理 basedOn，
    # 所以必须显式恢复这条继承链。
    "英文封面题目": "封面论文题目",
    "参考文献条目（著者出版年）": "参考文献条目",
}
for st in d.styles:
    if st.type != 1:
        continue
    if st.name in BASE_MAP and BASE_MAP[st.name] in [s.name for s in d.styles]:
        st.base_style = d.styles[BASE_MAP[st.name]]
    else:
        st.base_style = None

# 无编号的章级标题不能继承标题 1/2 的章编号。
# 局部 numId=0 可抑制继承的编号，又不改动段落的大纲级别或外观。
def ensure_no_numbering(style_name):
    """清除指定样式的编号绑定。"""
    try:
        st = d.styles[style_name]
    except KeyError:
        return
    pPr = st.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:numPr")):
        pPr.remove(old)
    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    numId = OxmlElement("w:numId")
    numId.set(qn("w:val"), "0")
    numPr.append(ilvl)
    numPr.append(numId)
    anchor = None
    for tag in ("w:spacing", "w:ind", "w:jc", "w:textDirection",
                "w:textAlignment", "w:outlineLvl", "w:rPr"):
        anchor = pPr.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(numPr)
    else:
        pPr.append(numPr)

for _name in ("章标题（无编号）", "章标题（无编号、不进目录）",
              "章内无编号小节", "授权说明标题", "名单页小节标题",
              "参考文献条目（著者出版年）"):
    ensure_no_numbering(_name)


def _set_keep_next(style_name, value=True):
    """把「与下段同页」（w:keepNext）写到样式上。

    规则里 sectionTitleUnnumbered.keepWithNext 已经是 true，构建阶段也写了。
    这里补写一份是给「祖先样式不带 keepNext」的情况兜底：这两个样式
    basedOn heading 2 -> heading 1，只有 heading 1 有 keepNext 时才由继承
    提供；祖先哪天去掉这条，样式层的显式值就直接顶上。
    """
    try:
        st = d.styles[style_name]
    except KeyError:
        return
    pPr = st.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:keepNext")):
        pPr.remove(old)
    if not value:
        return
    el = OxmlElement("w:keepNext")
    # w:pPr 子元素顺序：keepNext 在 keepLines 之前、pBdr/shd/tabs 等之先。
    anchor = None
    for tag in ("w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl",
                "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
                "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
                "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
                "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
                "w:spacing", "w:ind", "w:jc", "w:outlineLvl", "w:rPr"):
        anchor = pPr.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(el)
    else:
        pPr.append(el)


# 意见 6：章内无编号小节要与下段同页。规则为唯一事实来源，这里按规则取值。
# 注意 Word 的正常化行为：样式上写了与继承值相同的 keepNext（这两个样式
# basedOn heading 2 -> heading 1，heading 1 自带 keepNext），保存时会当作
# 冗余剥掉，但有效值仍是「与下段同页」，所以不需要、也不应该在段落上另写
# 直接格式。这里保留样式层补写，是为了在祖先样式不带 keepNext 时仍生效。
for _key, _name in (("sectionTitleUnnumbered", "章内无编号小节"),
                    ("committeeSectionTitle", "名单页小节标题")):
    _spec = ((RULES.get("styles") or {}).get(_key) or {}).get("paragraphFormat") or {}
    if "keepWithNext" in _spec:
        _set_keep_next(_name, bool(_spec["keepWithNext"]))

_image_spec = _style_pf("image")
if "widowControl" in _image_spec:
    d.styles["图片"].paragraph_format.widow_control = _image_spec["widowControl"]

# Word 在更新域后保存时，会丢掉未用附录 3/4 样式上的直接编号。
# 这里在最后一轮 Word 处理之后恢复与附录列表的绑定，
# 让这些样式保留 B.1.1 / B.1.1.1 编号。
def _appendix_num_id():
    """查找附录标题编号定义对应的编号 ID。"""
    try:
        numbering = d.part.numbering_part.element
    except Exception:
        return None
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for abstract in numbering.findall(qn("w:abstractNum")):
        first = abstract.find(qn("w:lvl"))
        if first is None:
            continue
        text = first.find(qn("w:lvlText"))
        if text is not None and text.get(qn("w:val")) == "附录%1":
            aid = abstract.get(qn("w:abstractNumId"))
            for num in numbering.findall(qn("w:num")):
                ref = num.find(qn("w:abstractNumId"))
                if ref is not None and ref.get(qn("w:val")) == aid:
                    return num.get(qn("w:numId"))
    return None

def _num_id_for_lvl_text(want):
    """按一级编号文本查找连续列表，跳过要求重新开始的列表。"""
    try:
        numbering = d.part.numbering_part.element
    except Exception:
        return None
    for abstract in numbering.findall(qn("w:abstractNum")):
        first = abstract.find(qn("w:lvl"))
        if first is None:
            continue
        text = first.find(qn("w:lvlText"))
        if text is not None and text.get(qn("w:val")) == want:
            aid = abstract.get(qn("w:abstractNumId"))
            for num in numbering.findall(qn("w:num")):
                if num.find(".//" + qn("w:startOverride")) is not None:
                    continue
                ref = num.find(qn("w:abstractNumId"))
                if ref is not None and ref.get(qn("w:val")) == aid:
                    return num.get(qn("w:numId"))
    return None


def _ensure_style_numbering(style_name, level, num_id):
    """把指定样式绑定到给定编号级别。"""
    try:
        st = d.styles[style_name]
    except KeyError:
        return
    if not num_id:
        return
    pPr = st.element.get_or_add_pPr()
    for old in pPr.findall(qn("w:numPr")):
        pPr.remove(old)
    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level - 1))
    numId = OxmlElement("w:numId")
    numId.set(qn("w:val"), str(num_id))
    numPr.append(ilvl)
    numPr.append(numId)
    anchor = None
    for tag in ("w:spacing", "w:ind", "w:jc", "w:textDirection",
                "w:textAlignment", "w:outlineLvl", "w:rPr"):
        anchor = pPr.find(qn(tag))
        if anchor is not None:
            break
    if anchor is not None:
        anchor.addprevious(numPr)
    else:
        pPr.append(numPr)

_apx_num = _appendix_num_id()
_ensure_style_numbering("附录二级节标题", 3, _apx_num)
_ensure_style_numbering("附录三级节标题", 4, _apx_num)
# Word 会重新排列编号定义。优先保留已有绑定，避免用同样显示为 [%1]
# 的个人成果列表替换参考文献列表。
_ref_binding = ref.element.find(qn("w:pPr") + "/" + qn("w:numPr") + "/" + qn("w:numId"))
if _ref_binding is None or _ref_binding.get(qn("w:val")) == "0":
    _ensure_style_numbering("参考文献条目", 1, _num_id_for_lvl_text("[%1]"))

# 内置列表样式：规则文件给它们正文格式（小四宋体 /
# Times New Roman、固定 20 磅、段前段后 0、两端对齐），并按层级写入左缩进
# 与 2 字符悬挂缩进。Word 保存时可能把样式的直接格式规范化掉，
# 因此在其他修正之后按规则恢复。
_LIST_INDENT_RULES = (
    ("List Number", "listNumber1"),
    ("List Number 2", "listNumber2"),
    ("List Number 3", "listNumber3"),
    ("List Bullet", "listBullet1"),
    ("List Bullet 2", "listBullet2"),
    ("List Bullet 3", "listBullet3"),
)
for _name, _key in _LIST_INDENT_RULES:
    try:
        _st = d.styles[_name]
    except KeyError:
        continue
    set_fonts(_st, "SimSun", "Times New Roman")
    _st.font.size = Pt(12)
    _pf = _st.paragraph_format
    _pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _pf.line_spacing = Pt(20)
    _pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    _pf.space_before = Pt(0)
    _pf.space_after = Pt(0)
    _rule_pf = _style_pf(_key)
    _left = _rule_pf.get("leftIndent") or {}
    _hang = _rule_pf.get("hangingIndent") or {}
    _pPr = _st.element.get_or_add_pPr()
    _ind = _pPr.get_or_add_ind()
    # 先清直接格式，再按规则写左缩进与悬挂缩进；首行缩进一律不写。
    for _attr in ("w:leftChars", "w:left", "w:firstLineChars", "w:firstLine",
                  "w:hangingChars", "w:hanging"):
        _ind.attrib.pop(qn(_attr), None)
    if _left.get("unit") == "character":
        _chars = int(round(float(_left["value"]) * 100))
        _ind.set(qn("w:leftChars"), str(_chars))
        _ind.set(qn("w:left"), str(int(round(_chars * 2.4))))
    if _hang.get("unit") == "character":
        _chars = int(round(float(_hang["value"]) * 100))
        _ind.set(qn("w:hangingChars"), str(_chars))
        _ind.set(qn("w:hanging"), str(int(round(_chars * 2.4))))

def normalize_east_asian_hint(doc):
    """全角引号等东亚宽度字符按中文字体渲染。

    Word 保存时会把中西文混排的 run 拆开，引号可能单独成 run；这类 run
    只带西文字体声明，Word 会显示成半角英文引号。加上 w:hint="eastAsia"
    后，这类字符改用 eastAsia 字体（宋体、黑体等），显示为全角。
    """
    for run in doc.element.body.iter(qn("w:r")):
        rPr = run.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            run.insert(0, rPr)
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:hint"), "eastAsia")


def normalize_footnote_text_style(doc):
    """脚注文本只保留 1.5 字符的悬挂缩进，不要叠加固定 twips 左缩进。"""
    for style in doc.styles:
        if (style.name or "").lower() != "footnote text":
            continue
        pPr = style.element.find(qn("w:pPr"))
        if pPr is None:
            continue
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            continue
        for attr in ("w:left", "w:leftChars", "w:firstLine", "w:firstLineChars"):
            ind.attrib.pop(qn(attr), None)
        # 左缩进 0 字符；只写悬挂 1.5 字符。
        ind.set(qn("w:left"), "0")
        ind.set(qn("w:leftChars"), "0")
        ind.set(qn("w:hanging"), "270")
        ind.set(qn("w:hangingChars"), "150")


def normalize_footnote_paragraphs(doc):
    """让脚注段落完全继承“footnote text”样式，清除 Word 带入的直接缩进/段距。"""
    for part in doc.part.package.iter_parts():
        if str(part.partname) != "/word/footnotes.xml":
            continue
        root = etree.fromstring(part.blob)
        for paragraph in root.iter(qn("w:p")):
            pPr = paragraph.find(qn("w:pPr"))
            if pPr is None:
                continue
            for tag in ("w:spacing",):
                for element in pPr.findall(qn(tag)):
                    pPr.remove(element)
            # 删除 Word 带入的直接缩进，完全继承“footnote text”样式
            # （左缩进 0 字符、悬挂 1.5 字符）。
            for element in pPr.findall(qn("w:ind")):
                pPr.remove(element)
        part._blob = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )


def normalize_footnote_reference_style(doc):
    """脚注文本中的编号用普通文本，正文引用标记由直接格式保持上标。"""
    for style in doc.styles:
        if (style.name or "").lower() != "footnote reference":
            continue
        rPr = style.element.find(qn("w:rPr"))
        if rPr is None:
            continue
        for vert in rPr.findall(qn("w:vertAlign")):
            rPr.remove(vert)

    # Word 保存时可能把正文脚注标记的上标折叠进样式；样式改为普通文本后，
    # 需要给正文中的标记补回直接上标。
    for ref in doc.element.body.iter(qn("w:footnoteReference")):
        run = ref.getparent()
        rPr = run.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            run.insert(0, rPr)
        if rPr.find(qn("w:vertAlign")) is None:
            vert = OxmlElement("w:vertAlign")
            vert.set(qn("w:val"), "superscript")
            rPr.append(vert)


normalize_east_asian_hint(d)
normalize_footnote_text_style(d)
normalize_footnote_paragraphs(d)
normalize_footnote_reference_style(d)

implementation = yaml.safe_load(VARIANT.implementation_map.read_text(encoding="utf-8"))
apply_document_text_flow(d, RULES, implementation)

d.save(DOCX)
print("postprocessed:", DOCX)
