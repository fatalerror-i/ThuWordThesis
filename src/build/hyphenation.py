"""按规则设置文档和样式的断字；不检查文档内容。"""
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def apply_paragraph_text_flow(ppr, spec):
    """按 Word 的文件语义写入段落属性。"""
    following = ("w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
                 "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                 "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                 "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
                 "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")
    for key, tag in (("hyphenation", "suppressAutoHyphens"), ("wordWrap", "wordWrap")):
        if key not in spec:
            continue
        for old in ppr.findall(qn("w:" + tag)):
            ppr.remove(old)
        node = OxmlElement("w:" + tag)
        # 两个规则值均按功能表达。Word 的 wordWrap=0 会在字符间拆词，
        # wordWrap=1 才让不适合断字的单词整体移到下一行。
        enabled = not spec[key]
        node.set(qn("w:val"), "1" if enabled else "0")
        before = ("w:kinsoku", "w:wordWrap") + following if key == "hyphenation" else following
        ppr.insert_element_before(node, *before)


def apply_all_style_word_wrap(styles, enabled):
    """给全部样式及已有表格条件段落格式显式写入开关。"""
    for style in styles.findall(qn("w:style")):
        ppr = style.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            following = {qn("w:" + tag) for tag in ("rPr", "tblPr", "trPr", "tcPr", "tblStylePr")}
            index = next((i for i, node in enumerate(style) if node.tag in following), len(style))
            style.insert(index, ppr)
        apply_paragraph_text_flow(ppr, {"wordWrap": enabled})
        for conditional in style.findall(qn("w:tblStylePr")):
            conditional_ppr = conditional.find(qn("w:pPr"))
            if conditional_ppr is not None:
                apply_paragraph_text_flow(conditional_ppr, {"wordWrap": enabled})


def apply_all_paragraph_word_wrap(root, enabled):
    """统一设置全部 Word 段落的直接格式，不读取文字或按内容判断。"""
    count = 0
    for paragraph in root.iter(qn("w:p")):
        apply_paragraph_text_flow(paragraph.get_or_add_pPr(), {"wordWrap": enabled})
        count += 1
    return count


def apply_document_text_flow(doc, rules, implementation):
    """恢复文档默认值和规则明确规定的样式设置，供生成和后处理共用。"""
    settings = doc.settings.element
    hyphenation = rules["pageSetup"]["hyphenation"]
    for tag, value in (("autoHyphenation", int(hyphenation["automatic"])),
                       ("consecutiveHyphenLimit", hyphenation["consecutiveLineLimit"]),
                       ("doNotHyphenateCaps", int(not hyphenation["hyphenateCaps"]))):
        for old in settings.findall(qn("w:" + tag)):
            settings.remove(old)
        node = OxmlElement("w:" + tag)
        node.set(qn("w:val"), str(value))
        settings.insert_element_before(node, "w:showEnvelope", "w:summaryLength",
                                       "w:clickAndTypeStyle", "w:defaultTableStyle",
                                       "w:evenAndOddHeaders", "w:compat", "w:docVars",
                                       "w:rsids", "m:mathPr", "w:updateFields")
    defaults = doc.styles.element.find(qn("w:docDefaults"))
    pdefault = defaults.find(qn("w:pPrDefault"))
    ppr = pdefault.find(qn("w:pPr"))
    apply_paragraph_text_flow(ppr, rules["pageSetup"]["documentDefaults"]["paragraphFormat"])
    # 对每个已定义样式显式写入全局换行设置，不只依赖 docDefaults/basedOn。
    # 包括未使用的内置样式，以及表格样式中单独覆盖的条件段落格式。
    default_flow = rules["pageSetup"]["documentDefaults"]["paragraphFormat"]
    if "wordWrap" in default_flow:
        apply_all_style_word_wrap(doc.styles.element, default_flow["wordWrap"])
    for key, style in rules["styles"].items():
        spec = style.get("paragraphFormat", {})
        if "hyphenation" in spec or "wordWrap" in spec:
            name = implementation["styleMap"][key]["name"]
            apply_paragraph_text_flow(doc.styles[name].element.get_or_add_pPr(), spec)
