"""不启动 Word 的静态检查：规则、母版 docx、dotm 包、Ribbon 回调。"""

from __future__ import annotations

import json
import re
import subprocess
import sys
import unittest
import xml.etree.ElementTree as ET
from zipfile import ZipFile

import yaml
from lxml import etree
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from pypdf import PdfReader
from pyopenvba.cfb import CFB
from pyopenvba.vba import parse_vba_project
from src.build.style_presentation import apply_style_presentation
from src.office_metadata import LIBID, PACKAGE_DATE

from tests.support import (
    VARIANTS,
    ProductCase,
    ROOT,
    product_paths,
    read_part,
    spacing_attr,
    style_block,
    style_names,
)


def style_para(style, attr):
    """沿 basedOn 链取实际的段落格式值。"""
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        value = getattr(style.paragraph_format, attr)
        if value is not None:
            return value
        style = style.base_style
    return None


def load_degree_config():
    """读取 configs/degree-options.yaml（学位选项的唯一数据源）。"""
    path = ROOT / "configs" / "degree-options.yaml"
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def find_sdt(container, tag):
    """在段落/单元格等 XML 容器中按标签查找内容控件。"""
    for sdt in container.iter(qn("w:sdt")):
        tag_el = sdt.find(".//" + qn("w:tag"))
        if tag_el is not None and tag_el.get(qn("w:val")) == tag:
            return sdt
    return None


class ProjectStaticTest(unittest.TestCase):
    """不随变体变化的项目级检查：VBA 源码、规则校验、交付物结构、清单统一性。"""

    def test_paragraph_word_wrap_overrides_direct_format_everywhere(self):
        from docx.oxml import parse_xml
        from src.build.hyphenation import apply_all_paragraph_word_wrap
        root = parse_xml('''<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body><w:p><w:pPr><w:suppressAutoHyphens/><w:wordWrap w:val="0"/></w:pPr><w:r><w:t>保留文字</w:t></w:r></w:p>
          <w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
          <w:p><w:r><w:txbxContent><w:p/></w:txbxContent></w:r></w:p></w:body></w:document>''')
        self.assertEqual(apply_all_paragraph_word_wrap(root, False), 4)
        for paragraph in root.iter(qn("w:p")):
            self.assertEqual(paragraph.find("w:pPr/w:wordWrap", {"w": root.nsmap["w"]}).get(qn("w:val")), "1")
        self.assertEqual(len(list(root.iter(qn("w:suppressAutoHyphens")))), 1)
        self.assertEqual(next(root.iter(qn("w:t"))).text, "保留文字")

    def test_all_style_word_wrap_overrides_unused_and_conditional_styles(self):
        from docx.oxml import parse_xml
        from src.build.hyphenation import apply_all_style_word_wrap
        root = parse_xml('''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="Unused"/></w:style>
            <w:style w:type="table" w:styleId="Table"><w:name w:val="Table"/>
              <w:pPr><w:keepNext/><w:wordWrap w:val="0"/></w:pPr>
              <w:tblStylePr w:type="firstRow"><w:pPr><w:wordWrap w:val="0"/></w:pPr></w:tblStylePr>
            </w:style></w:styles>''')
        for _ in range(2):
            apply_all_style_word_wrap(root, False)
        wraps = list(root.iter(qn("w:wordWrap")))
        self.assertEqual(len(wraps), 3)
        self.assertTrue(all(node.get(qn("w:val")) == "1" for node in wraps))
        self.assertEqual(len(list(root.iter(qn("w:keepNext")))), 1)

    def test_style_priority_keeps_original_relationships(self) -> None:
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        styles = etree.fromstring(f"""
            <w:styles xmlns:w="{ns}">
              <w:latentStyles w:defUIPriority="99">
                <w:lsdException w:name="First latent" w:uiPriority="4" w:unhideWhenUsed="1"/>
                <w:lsdException w:name="Second latent" w:uiPriority="25"/>
              </w:latentStyles>
              <w:style><w:name w:val="Normal"/></w:style>
              <w:style><w:name w:val="Cover"/></w:style>
              <w:style><w:name w:val="First"/><w:uiPriority w:val="4"/><w:qFormat/></w:style>
              <w:style><w:name w:val="Second"/><w:uiPriority w:val="25"/></w:style>
              <w:style><w:name w:val="Last"/></w:style>
            </w:styles>
        """)
        implementation = {
            "styleMap": {"body": {"name": "Normal"}, "cover": {"name": "Cover"}},
            "stylePresentation": [
                {"priority": 1, "quickStyle": True, "styles": ["body"]},
                {"priority": 100, "styles": ["cover"]},
            ],
        }
        apply_style_presentation(styles, implementation["styleMap"], implementation)
        explicit = styles.findall(qn("w:style"))
        self.assertEqual([int(s.find(qn("w:uiPriority")).get(qn("w:val"))) for s in explicit],
                         [0, 96, 97, 98, 99])
        latent = styles.find(qn("w:latentStyles"))
        self.assertEqual([e.get(qn("w:uiPriority")) for e in latent], ["97", "98"])
        self.assertEqual([e.get(qn("w:unhideWhenUsed")) for e in latent], ["0", "0"])
        once = etree.tostring(styles)
        apply_style_presentation(styles, implementation["styleMap"], implementation)
        self.assertEqual(etree.tostring(styles), once)

    def test_vba_has_no_uncommented_chinese(self) -> None:
        """VBA 模块级/语句级的中文必须以单引号注释，否则整模块编译失败。"""
        for path in sorted((ROOT / "src" / "dotm" / "vba").glob("*.bas")):
            for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
                stripped = line.lstrip()
                if stripped and re.match(r"[一-鿿]", stripped):
                    self.fail(f"{path.name}:{lineno} 中文未注释：{line!r}")
    def test_chapter_reference_uses_chapter_bookmark(self) -> None:
        """检查章引用改用 REF 书签而非 STYLEREF。"""
        source = (ROOT / "src" / "dotm" / "vba" / "modMain.bas").read_text(encoding="utf-8")
        self.assertIn('BM_CH_PREFIX As String = "THU_CH_"', source)
        self.assertIn('ChapterReferenceFieldCode = "REF " & bmName & " ' + chr(92) + 'n ' + chr(92) + 'h"', source)
        self.assertIn("EnsureChapterBookmark", source)
        self.assertNotIn('ChapterReferenceCode = "STYLEREF', source)
    def test_rules_validate(self) -> None:
        """以子进程运行规则校验并要求通过。"""
        result = subprocess.run(
            [sys.executable, "src/validate_rules.py"],
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=120,
        )
        self.assertEqual(
            result.returncode,
            0,
            f"validate_rules.py 失败：\n{result.stdout}\n{result.stderr}",
        )
        self.assertIn("OK", result.stdout)
    def test_delivery_artifacts(self) -> None:
        """交付物本身可自证：宏与功能区在包内，级联入口在源码里。"""
        # ThisDocument 级联入口
        this_document = ROOT / "src" / "dotm" / "vba" / "ThisDocument-graduate.txt"
        self.assertTrue(this_document.exists(), this_document)
        self.assertIn(
            "Document_ContentControlOnExit",
            this_document.read_text(encoding="utf-8"),
        )
        # 两份 dotm 都带 VBA 与功能区 XML
        for variant in VARIANTS:
            paths = product_paths(variant)
            with ZipFile(paths["dotm"]) as z:
                names = set(z.namelist())
            self.assertIn("word/vbaProject.bin", names, variant)
            self.assertIn("customUI/customUI14.xml", names, variant)
        # 帮助入口已回退：功能区无“功能说明”，构建不生成帮助窗体/模块。
        ribbon = (ROOT / "src" / "dotm" / "customUI14-graduate.xml").read_text(encoding="utf-8")
        self.assertNotIn('id="btnHelp"', ribbon)
        self.assertNotIn('onAction="OpenHelp"', ribbon)
        forms_spec = json.loads(
            (ROOT / "src" / "dotm" / "forms" / "forms.json").read_text(encoding="utf-8")
        )
        self.assertNotIn("frmHelp", forms_spec)
        mod_main = (ROOT / "src" / "dotm" / "vba" / "modMain.bas").read_text(encoding="utf-8")
        self.assertNotIn("Public Sub OpenHelp", mod_main)
        self.assertNotIn("HelpHtml", mod_main)

    def test_built_artifacts_omit_private_metadata(self) -> None:
        for variant in VARIANTS:
            paths = product_paths(variant)
            for kind in ("docx", "dotm"):
                with self.subTest(variant=variant, kind=kind), ZipFile(paths[kind]) as z:
                    self.assertFalse(z.comment)
                    self.assertTrue(all(info.date_time == PACKAGE_DATE and
                                        not (info.extra or info.comment)
                                        for info in z.infolist()))
                    self.assertNotIn("docProps/custom.xml", z.namelist())
                    core = etree.fromstring(z.read("docProps/core.xml"))
                    self.assertEqual({etree.QName(x).localname for x in core}, {"title", "subject"})
                    app = etree.fromstring(z.read("docProps/app.xml"))
                    self.assertTrue({etree.QName(x).localname for x in app} <= {
                        "Pages", "Words", "Characters", "CharactersWithSpaces", "Lines", "Paragraphs",
                    })
                    if kind == "dotm":
                        cfb = CFB.from_bytes(z.read("word/vbaProject.bin"))
                        project = parse_vba_project(cfb)
                        self.assertTrue(project.references)
                        self.assertTrue(all(not m[1] for m in LIBID.finditer(project.dir_raw)))
                        self.assertTrue(all(m.text_offset == 0 for m in project.modules))
                        self.assertEqual(cfb.get_stream_in_storage("VBA", "_VBA_PROJECT"),
                                         b"\xcc\x61\xff\xff\x00\x01\x00")
                        self.assertFalse(any(n.startswith("__SRP_")
                                             for n in cfb.list_streams_in_storage("VBA")))
                        self.assertTrue(all(not any(e.raw[100:116]) for e in cfb._directory))
            with self.subTest(variant=variant, kind="pdf"):
                pdf = PdfReader(paths["pdf"])
                self.assertNotIn("/Info", pdf.trailer)
                self.assertNotIn("/ID", pdf.trailer)
                self.assertNotIn("/Metadata", pdf.root_object)
                self.assertIn("/StructTreeRoot", pdf.root_object)

    def test_preview_achievement_list_starts_at_one(self) -> None:
        for variant in VARIANTS:
            with self.subTest(variant=variant):
                pdf = PdfReader(product_paths(variant)["pdf"])
                lines = [line.strip() for page in pdf.pages for line in page.extract_text().splitlines()
                         if "作者. 论文题目[J]." in line]
                self.assertEqual(len(lines), 1)
                self.assertTrue(lines[0].startswith("[1] "), lines)

    def test_example_fields_and_placeholder_styles(self) -> None:
        for variant in VARIANTS:
            for kind in ("docx", "dotm"):
                with self.subTest(variant=variant, kind=kind), ZipFile(product_paths(variant)[kind]) as z:
                    root = etree.fromstring(z.read("word/document.xml"))
                    styles = etree.fromstring(z.read("word/styles.xml"))
                    ns = {"w": root.nsmap["w"],
                          "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"}
                    fields = [text.strip() for text in root.xpath("//w:instrText/text()", namespaces=ns)]
                    citations = [code for code in fields if code.startswith('QUOTE "THU_CITE:')]
                    self.assertEqual(citations, [
                        'QUOTE "THU_CITE:0001"', 'QUOTE "THU_CITE:0002"',
                        'QUOTE "THU_CITE:0001"', 'QUOTE "THU_CITE:0003"',
                        'QUOTE "THU_CITE:0001;0003"',
                    ])
                    targets = root.xpath(
                        '//w:bookmarkStart[starts-with(@w:name,"THU_REF_ENTRY_")]/@w:name',
                        namespaces=ns,
                    )
                    self.assertEqual(targets, [f"THU_REF_ENTRY_{i:04}" for i in (1, 2, 3)])
                    paragraph_codes = [
                        "".join(p.xpath("w:r/w:instrText/text()", namespaces=ns)).strip()
                        for p in root.xpath("//w:body/w:p", namespaces=ns)
                    ]
                    self.assertEqual([code for code in paragraph_codes if code.startswith("SEQ 子图")], [
                        r"SEQ 子图 \* alphabetic", r"SEQ 子图 \* alphabetic",
                    ])
                    caption = styles.xpath(
                        '//w:style[w:name/@w:val="caption" or w:name/@w:val="Caption"]',
                        namespaces=ns,
                    )
                    self.assertEqual(len(caption), 1)
                    self.assertEqual(caption[0].xpath("w:pPr/w:jc/@w:val", namespaces=ns), ["center"])
                    placeholders = root.xpath("//wps:txbx/w:txbxContent/w:p", namespaces=ns)
                    self.assertEqual(len(placeholders), 4)
                    for p in placeholders:
                        self.assertEqual(p.xpath("w:pPr/w:pStyle/@w:val", namespaces=ns),
                                         [caption[0].get(qn("w:styleId"))])
                        self.assertTrue(all(v == "0" for v in p.xpath(
                            "w:pPr/w:ind/@w:firstLine | w:pPr/w:ind/@w:firstLineChars", namespaces=ns,
                        )))
                    self.assertEqual(root.xpath("//wps:txbx/../wps:bodyPr/@anchor", namespaces=ns),
                                     ["ctr"] * len(placeholders))

    def test_numbered_examples_have_body_references(self) -> None:
        for variant in VARIANTS:
            for kind in ("docx", "dotm"):
                with self.subTest(variant=variant, kind=kind), ZipFile(product_paths(variant)[kind]) as z:
                    root = etree.fromstring(z.read("word/document.xml"))
                    ns = {"w": root.nsmap["w"]}
                    bookmark_name = re.compile(r"THU_(?:FIG|TAB|EQ|SUBFIG)_[A-Z0-9]+")
                    targets = {
                        name for name in root.xpath("//w:bookmarkStart/@w:name", namespaces=ns)
                        if bookmark_name.fullmatch(name)
                    }
                    self.assertEqual(len(targets), 14)  # 3 图、2 子图、7 表、2 公式
                    paragraphs = root.xpath("w:body/w:p", namespaces=ns)
                    refs_by_paragraph = [
                        set(re.findall(r"\bREF\s+(" + bookmark_name.pattern + r")\b",
                                       " ".join(p.xpath(".//w:instrText/text()", namespaces=ns))))
                        for p in paragraphs
                    ]
                    self.assertIn({"THU_TAB_0003", "THU_TAB_0005"}, refs_by_paragraph)
                    self.assertIn({"THU_FIG_0002", "THU_SUBFIG_A", "THU_SUBFIG_B"}, refs_by_paragraph)
                    self.assertEqual(set().union(*refs_by_paragraph), targets - {"THU_TAB_0004"})

    def test_both_variants_tof_unified(self) -> None:
        r"""两变体图表清单统一为 \c "图"/\c "表"，条目顶格且不含附录图表。

        回归 1：Word 更新 Table of Figures 域时会实例化内置 "table of figures"
        样式（leftChars/left=200）并给每条结果段落盖直接 left=960 hanging=480，
        表现为用户报告的"图表目录出现不该出现的左边距"；postprocess_docx.py
        每次字段更新后清掉样式层与段落层的 w:ind。
        回归 2：两变体清单都按题注标签收集，附录题注改用独立 SEQ 标识符
        （附录图/附录表），因此清单条目里不能出现 表A.x / 图A.x。
        回归 3：指南 2.3.7 要求条目宋体小四、固定 20 磅、段前段后 0 磅；
        Word 内置样式默认段后 10 磅 + 1.15 倍行距，postprocess_docx.py
        须在样式层显式写 after=0 / line=400(exact) 并清掉条目的直接 spacing。
        """
        variants = {
            "graduate": ROOT / "output" / "清华大学研究生学位论文模板.dotm",
            "undergraduate": ROOT / "output" / "清华大学综合论文训练论文模板.dotm",
        }
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        for variant, dotm in variants.items():
            with self.subTest(variant=variant):
                if not dotm.exists():
                    self.skipTest(f"缺少 {variant} dotm：{dotm}")
                with ZipFile(dotm) as zf:
                    styles = ET.fromstring(zf.read("word/styles.xml"))
                    document = ET.fromstring(zf.read("word/document.xml"))

                # 统一样式：两变体都必须有且仅有一个 table of figures 样式
                tof_ids = []
                for st in styles.findall("w:style", ns):
                    nm = st.find("w:name", ns)
                    if nm is not None and (nm.get(qn("w:val")) or "").lower() == "table of figures":
                        sid = st.get(qn("w:styleId"))
                        tof_ids.append(sid)
                        pPr = st.find("w:pPr", ns)
                        ind = pPr.find("w:ind", ns) if pPr is not None else None
                        self.assertIsNone(ind, f"{variant} table of figures 样式仍含 w:ind（左缩进）")
                        if ind is not None:
                            self.assertIsNone(ind.get(qn("w:leftChars")))
                            self.assertIsNone(ind.get(qn("w:left")))
                        # 指南 2.3.7：段前段后 0 磅、固定 20 磅（line=400 twips, exact）
                        self.assertIsNotNone(pPr, f"{variant} table of figures 样式缺 w:pPr")
                        sp = pPr.find("w:spacing", ns) if pPr is not None else None
                        self.assertIsNotNone(sp, f"{variant} table of figures 样式缺 w:spacing")
                        if sp is not None:
                            # Word 在 SaveAs2 时会去掉值为 0 的 w:before/w:after
                            # （从 docDefaults 继承 0），TOC 1 也是如此。
                            self.assertIn(sp.get(qn("w:before")), (None, "0"),
                                          f"{variant} 清单条目样式段前不为 0")
                            self.assertIn(sp.get(qn("w:after")), (None, "0"),
                                          f"{variant} 清单条目样式段后不为 0")
                            self.assertEqual(sp.get(qn("w:line")), "400",
                                             f"{variant} 清单条目样式行距不是固定 20 磅")
                            self.assertEqual(sp.get(qn("w:lineRule")), "exact",
                                             f"{variant} 清单条目样式行距规则不是 fixed")
                self.assertEqual(len(tof_ids), 1, f"{variant} 应有且仅有一个 table of figures 样式：{tof_ids}")

                # 统一域机制：\c "图" / \c "表" 存在，旧的 \t "图注,4" 不存在
                instr = "".join(
                    (t.text or "") for t in document.iter(qn("w:instrText"))
                )
                self.assertIn('\\c "图"', instr, f"{variant} 缺少 \\c \"图\" 清单域")
                self.assertIn('\\c "表"', instr, f"{variant} 缺少 \\c \"表\" 清单域")
                self.assertNotIn('图注,4', instr, f"{variant} 仍在用 \\t \"图注,4\" 旧机制")
                self.assertNotIn('表注,4', instr, f"{variant} 仍在用 \\t \"表注,4\" 旧机制")

                # 条目顶格，且不含附录图表
                entries = 0
                for p in document.iter(qn("w:p")):
                    pPr = p.find(qn("w:pPr"))
                    if pPr is None:
                        continue
                    ps = pPr.find(qn("w:pStyle"))
                    if ps is None or ps.get(qn("w:val")) not in tof_ids:
                        continue
                    entries += 1
                    self.assertIsNone(
                        pPr.find(qn("w:ind")),
                        f"{variant} 图表清单条目段落仍含直接 w:ind（左缩进）",
                    )
                    self.assertIsNone(
                        pPr.find(qn("w:spacing")),
                        f"{variant} 图表清单条目段落仍含直接 w:spacing（盖掉样式层段间距）",
                    )
                    text = "".join((t.text or "") for t in p.iter(qn("w:t")))
                    self.assertNotIn("表A.", text, f"{variant} 附表清单含附录表条目：{text}")
                    self.assertNotIn("图A.", text, f"{variant} 插图清单含附录图条目：{text}")
                self.assertGreaterEqual(entries, 1, f"{variant} dotm 未找到图表清单条目段落")


class _SharedStaticMixin:
    """两变体共用的静态断言。

    由 GraduateStaticTest / UndergraduateStaticTest 各继承一次：
    断言只写一份，两个变体各跑一遍，差异只体现在 counts 与产物路径。
    """

    def test_document_and_style_hyphenation(self) -> None:
        """断字来自文档默认与样式定义，两种交付格式保持一致。"""
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        for path in (self.docx_path, self.dotm_path):
            with ZipFile(path) as archive:
                settings = etree.fromstring(archive.read("word/settings.xml"))
                styles = etree.fromstring(archive.read("word/styles.xml"))
                for part in archive.namelist():
                    if part.startswith("word/") and part.endswith(".xml"):
                        content = etree.fromstring(archive.read(part))
                        for paragraph in content.findall(".//w:p", ns):
                            node = paragraph.find("w:pPr/w:wordWrap", ns)
                            self.assertIsNotNone(node, (path.name, part))
                            self.assertIn(node.get(qn("w:val")), ("1", "true", "on"))
                if "word/stylesWithEffects.xml" in archive.namelist():
                    effects = etree.fromstring(archive.read("word/stylesWithEffects.xml"))
                    for style in effects.findall("w:style", ns):
                        node = style.find("w:pPr/w:wordWrap", ns)
                        self.assertIsNotNone(node, (path.name, "stylesWithEffects", style.get(qn("w:styleId"))))
                        self.assertIn(node.get(qn("w:val")), ("1", "true", "on"))
            for tag, expected in (("autoHyphenation", "1"), ("doNotHyphenateCaps", "1"),
                                  ("consecutiveHyphenLimit", "2")):
                node = settings.find("w:" + tag, ns)
                self.assertIsNotNone(node, (path.name, tag))
                self.assertIn(node.get(qn("w:val"), "1"), (expected, "true") if expected == "1" else (expected,))
            defaults = styles.find("w:docDefaults/w:pPrDefault/w:pPr", ns)
            default_wrap = defaults.find("w:wordWrap", ns)
            # Word 保存 DOTM 时可以省略与应用默认值相同的文档默认属性。
            if path.suffix == ".docx":
                self.assertIsNotNone(default_wrap, path.name)
            if default_wrap is not None:
                self.assertIn(default_wrap.get(qn("w:val"), "1"), ("1", "true", "on"))
            by_id = {style.get(qn("w:styleId")): style for style in styles.findall("w:style", ns)}
            by_name = {style.find("w:name", ns).get(qn("w:val")).casefold(): style for style in by_id.values()}

            def value(name, tag):
                style = by_name[name.casefold()]
                while style is not None:
                    node = style.find("w:pPr/w:" + tag, ns)
                    if node is not None:
                        return node.get(qn("w:val"), "1") not in ("0", "false", "off")
                    parent = style.find("w:basedOn", ns)
                    style = by_id.get(parent.get(qn("w:val"))) if parent is not None else None
                node = defaults.find("w:" + tag, ns)
                return node is not None and node.get(qn("w:val"), "1") not in ("0", "false", "off")

            for name in ("Normal", "heading 1", "table of figures", "表格文字"):
                self.assertFalse(value(name, "suppressAutoHyphens"), name)
                self.assertTrue(value(name, "wordWrap"), name)
            for name in ("公式", "行间代码", "参考文献条目", "参考文献条目（著者出版年）"):
                self.assertTrue(value(name, "suppressAutoHyphens"), name)
                self.assertTrue(value(name, "wordWrap"), name)

            for name in by_name:
                # 每个样式显式设置按词换行。
                node = by_name[name].find("w:pPr/w:wordWrap", ns)
                self.assertIsNotNone(node, (path.name, name))
                self.assertIn(node.get(qn("w:val")), ("1", "true", "on"))
            for node in styles.findall(".//w:tblStylePr/w:pPr/w:wordWrap", ns):
                self.assertIn(node.get(qn("w:val")), ("1", "true", "on"))

    def test_heading_number_includes_ideographic_separator(self) -> None:
        """章、节及附录编号自带间隔，标题文字不额外放空格。"""
        for paragraph in self.doc.paragraphs:
            if paragraph.style.name in ("Heading 1", "Heading 2", "Heading 3", "Heading 4") or paragraph.style.name.startswith("附录"):
                self.assertFalse(paragraph.text.startswith("　"), paragraph.text)
        with ZipFile(self.docx_path) as zf:
            numbering = etree.fromstring(zf.read("word/numbering.xml"))
        texts = numbering.xpath('//*[local-name()="lvlText"]/@*[local-name()="val"]')
        for value in ("第%1章　", "%1.%2　", "%1.%2.%3　", "%1.%2.%3.%4　", "附录%1　"):
            self.assertIn(value, texts)
    def test_no_broken_field_results(self) -> None:
        """SEQ 复位域必须使用字面量 \r/\h，不能变成回车/退格。"""
        doc = self.doc
        broken = [p.text for p in doc.paragraphs if "错误!未定义书签" in p.text]
        self.assertEqual(broken, [])
        with ZipFile(self.docx_path) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        self.assertNotIn("\r", xml)
        self.assertNotIn("\x08", xml)
    def test_no_hidden_text(self) -> None:
        """全文不允许出现任何隐藏文字（包括空段里的隐藏域代码）。"""
        for path in (self.docx_path, self.dotm_path):
            if not path.exists():
                continue
            with ZipFile(path) as zf:
                for part in ("word/document.xml", "word/styles.xml"):
                    xml = zf.read(part).decode("utf-8")
                    self.assertNotIn("<w:vanish", xml, f"{path.name}:{part} 含隐藏文字")
    def test_code_styles(self) -> None:
        """检查行内与行间代码样式的类型与字体。"""
        doc = self.doc
        inline = doc.styles["行内代码"]
        block = doc.styles["行间代码"]
        self.assertEqual(inline.type, WD_STYLE_TYPE.CHARACTER)
        self.assertEqual(block.type, WD_STYLE_TYPE.PARAGRAPH)
        for style in (inline, block):
            xml = style.element.xml
            self.assertIn("Consolas", xml)
            # Word 保存往返后会把中文字体存成仿宋。
            self.assertTrue("FangSong" in xml or "仿宋" in xml, style.name)
        self.assertIsNotNone(block.element.find(".//" + qn("w:numPr")))
        self.assertIn(block.paragraph_format.space_before, (None, Pt(0)))
        # Word 在 SaveAs2 时会去掉值为 0 的 w:before/w:after（从 docDefaults
        # 继承 0），所以属性缺省与 Pt(0) 都表示段后 0 磅。
        self.assertIn(block.paragraph_format.space_after, (None, Pt(0)))
        self.assertEqual(block.paragraph_format.line_spacing, 1.0)
        self.assertIsNotNone(block.element.find(".//" + qn("w:spacing")))
        self.assertIsNotNone(block.element.find(".//" + qn("w:autoSpaceDE")))
    def test_table_text_style(self) -> None:
        """表格文字样式：段前段后 3 磅，且示例表格单元格实际套用该样式。"""
        doc = self.doc
        style = doc.styles["表格文字"]
        self.assertEqual(style.type, WD_STYLE_TYPE.PARAGRAPH)
        pf = style.paragraph_format
        self.assertAlmostEqual(style_para(style, "space_before").pt, 3.0)
        self.assertAlmostEqual(style_para(style, "space_after").pt, 3.0)
        self.assertEqual(pf.first_line_indent or 0, 0)
        self.assertIsNotNone(style.font.size)
        self.assertAlmostEqual(style.font.size.pt, 11.0)
        # 正文示例表格（三线表）单元格应显式使用“表格文字”，且实际文字为宋体/Times 11pt
        for table in [t for t in doc.tables if t.style.name == "三线表"]:
            p = table.rows[0].cells[0].paragraphs[0]
            self.assertEqual(p.style.name, "表格文字")
            sb = p.paragraph_format.space_before or style_para(p.style, "space_before")
            sa = p.paragraph_format.space_after or style_para(p.style, "space_after")
            self.assertAlmostEqual(sb.pt, 3.0)
            self.assertAlmostEqual(sa.pt, 3.0)
            run = p.runs[0]
            self.assertTrue(run.text)
            self.assertTrue(p.paragraph_format.line_spacing in (None, 1.0))
    def test_list_styles_and_ribbon(self) -> None:
        """有序/无序列表的编号格式与功能区菜单必须一致。"""
        doc = self.doc
        expected = {
            "List Number": ("decimal", "%1."),
            "List Number 2": ("decimal", "(%1)"),
            "List Number 3": ("lowerLetter", "%1."),
            "List Bullet": ("bullet", "•"),
            "List Bullet 2": ("bullet", "◦"),
            "List Bullet 3": ("bullet", "*"),
        }
        with ZipFile(self.docx_path) as zf:
            numbering = ET.fromstring(zf.read("word/numbering.xml"))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        nums = {
            n.get(qn("w:numId")): n.find("w:abstractNumId", ns).get(qn("w:val"))
            for n in numbering.findall("w:num", ns)
        }
        abstracts = {
            a.get(qn("w:abstractNumId")): a for a in numbering.findall("w:abstractNum", ns)
        }
        for style_name, (fmt, text) in expected.items():
            style = doc.styles[style_name]
            num_pr = style.element.find(".//" + qn("w:numPr"))
            self.assertIsNotNone(num_pr, style_name)
            num_id = num_pr.find("w:numId", ns).get(qn("w:val"))
            abstract = abstracts[nums[num_id]]
            level = abstract.find("w:lvl", ns)
            self.assertEqual(level.find("w:numFmt", ns).get(qn("w:val")), fmt, style_name)
            self.assertEqual(level.find("w:lvlText", ns).get(qn("w:val")), text, style_name)
            # 列表样式与正文同格式：小四、固定 20 磅、段前后 0、两端对齐、悬挂缩进 2 字符
            pf = style.paragraph_format
            self.assertAlmostEqual(style.font.size.pt, 12.0, msg=style_name)
            self.assertAlmostEqual(pf.line_spacing.pt, 20.0, msg=style_name)
            self.assertEqual(pf.line_spacing_rule, WD_LINE_SPACING.EXACTLY, style_name)
            self.assertEqual(pf.space_before or 0, 0, style_name)
            self.assertEqual(pf.space_after or 0, 0, style_name)
            self.assertEqual(pf.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY, style_name)
            self.assertAlmostEqual(pf.first_line_indent.pt, -24.0, msg=style_name)

        ribbon = ET.parse(self.ribbon_path)
        m = {"m": "http://schemas.microsoft.com/office/2009/07/customui"}
        buttons = {b.get("id"): b.get("label") for b in ribbon.iter("{%s}button" % m["m"])}
        self.assertEqual(buttons.get("btnListNumber1"), "1.")
        self.assertEqual(buttons.get("btnListNumber2"), "(1)")
        self.assertEqual(buttons.get("btnListNumber3"), "a.")
        self.assertEqual(buttons.get("btnListBullet1"), "•")
        self.assertEqual(buttons.get("btnListBullet2"), "◦")
        self.assertEqual(buttons.get("btnListBullet3"), "*")
        self.assertEqual(buttons.get("btnListContinue"), "继续编号")
        self.assertEqual(buttons.get("btnListRestart"), "重新开始于1")
        self.assertEqual(buttons.get("btnReferenceNumeric"), "顺序编码制")
        self.assertEqual(buttons.get("btnReferenceAuthorYear"), "著者出版年制")
        self.assertNotIn("btnListReference", buttons)
        self.assertEqual(buttons.get("btnOtherChapter"), "章标题")
        self.assertEqual(buttons.get("btnUnnumberedSubsection"), "章内无编号小节")
        self.assertNotIn("btnListNumber4", buttons)
        self.assertNotIn("btnListBullet5", buttons)
        # 参考文献菜单的图标：Word 内置 imageMso LabelInsert（标签）。
        ref_menu = next(
            el for el in ribbon.iter("{%s}menu" % m["m"])
            if el.get("id") == "mnuReferenceFormat"
        )
        self.assertEqual(ref_menu.get("label"), "参考文献")
        self.assertEqual(ref_menu.get("imageMso"), "LabelInsert")
    def test_authorization_text_format(self) -> None:
        """检查授权说明文字的字号与对齐。"""
        style = self.doc.styles["授权说明文字"]
        pf = style.paragraph_format
        self.assertEqual(style.font.size.pt, 14)
        self.assertEqual(pf.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
        self.assertEqual(pf.line_spacing.pt, 26)
        # w:before/w:after 缺省等价于 0。
        self.assertTrue(pf.space_before is None or pf.space_before.pt == 0)
        self.assertTrue(pf.space_after is None or pf.space_after.pt == 0)
        ind = style.element.find(".//" + qn("w:ind"))
        self.assertIsNotNone(ind)
        self.assertEqual(ind.get(qn("w:firstLineChars")), "200")
    def test_reference_entry_spacing(self) -> None:
        """参考文献条目：段前 3 磅、段后 0 磅，悬挂缩进 1 厘米。"""
        doc = self.doc
        style = doc.styles["参考文献条目"]
        self.assertEqual(style.paragraph_format.space_before, Pt(3))
        self.assertEqual(style.paragraph_format.space_after, Pt(0))
        style_ind = style.element.get_or_add_pPr().get_or_add_ind()
        self.assertEqual(style_ind.get(qn("w:left")), "567")
        self.assertEqual(style_ind.get(qn("w:hanging")), "567")
        self.assertIsNone(style_ind.get(qn("w:hangingChars")))
        self.assertIsNone(style_ind.get(qn("w:leftChars")))
        self.assertIsNone(style_ind.get(qn("w:firstLineChars")))
        ay = doc.styles["参考文献条目（著者出版年）"]
        self.assertEqual(ay.base_style.name, "参考文献条目")
        self.assertEqual(style_para(ay, "space_before"), Pt(3))
        self.assertEqual(style_para(ay, "space_after"), Pt(0))
        ay_ind = ay.element.get_or_add_pPr().get_or_add_ind()
        self.assertEqual(ay_ind.get(qn("w:left")), "567")
        self.assertEqual(ay_ind.get(qn("w:hanging")), "567")
        self.assertIsNone(ay_ind.get(qn("w:hangingChars")))
        ay_num = ay.element.find(".//" + qn("w:numPr"))
        self.assertIsNotNone(ay_num)
        ay_num_id = ay_num.find(qn("w:numId"))
        self.assertEqual(ay_num_id.get(qn("w:val")), "0")
        refs = [p for p in doc.paragraphs if p.style.name == "参考文献条目"]
        self.assertTrue(refs, "缺少参考文献条目段落")
        for p in refs:
            self.assertEqual(p.paragraph_format.space_before, Pt(3))
            self.assertEqual(p.paragraph_format.space_after, Pt(0))
            ind = p._p.get_or_add_pPr().get_or_add_ind()
            self.assertEqual(ind.get(qn("w:left")), "567")
            self.assertEqual(ind.get(qn("w:hanging")), "567")
            self.assertIsNone(ind.get(qn("w:hangingChars")))
            self.assertIsNone(ind.get(qn("w:leftChars")))
            self.assertIsNone(ind.get(qn("w:firstLineChars")))
    def test_symbol_entry_indent(self) -> None:
        """符号说明条目：符号在版心左缘，说明从 12 字符（5.08 厘米）处左对齐。"""
        doc = self.doc
        style = doc.styles["符号说明条目"]
        ind = style.element.get_or_add_pPr().get_or_add_ind()
        # 小四 12 pt × 12 字符 = 2880 twips = 144 pt = 5.08 cm
        self.assertEqual(ind.get(qn("w:left")), "2880")
        self.assertEqual(ind.get(qn("w:hanging")), "2880")
        self.assertIsNone(ind.get(qn("w:leftChars")))
        self.assertIsNone(ind.get(qn("w:hangingChars")))
        self.assertEqual(style.paragraph_format.left_indent, Pt(144))
        self.assertEqual(style.paragraph_format.first_line_indent, Pt(-144))
        tabs = style.element.get_or_add_pPr().find(qn("w:tabs"))
        self.assertIsNotNone(tabs, "符号说明条目缺少制表位")
        tab = tabs.find(qn("w:tab"))
        self.assertEqual(tab.get(qn("w:val")), "left")
        self.assertEqual(tab.get(qn("w:pos")), "2880")
        entries = [p for p in doc.paragraphs if p.style.name == "符号说明条目"]
        self.assertTrue(entries, "缺少符号说明条目段落")
        for p in entries:
            self.assertIn(chr(9), p.text)
            pPr = p._p.pPr
            p_ind = pPr.find(qn("w:ind")) if pPr is not None else None
            self.assertIsNone(p_ind, "符号说明条目段落不应带直接缩进")
    def test_cover_security_is_isolated(self) -> None:
        """本科密级在封面首行，研究生在首节页眉；默认空白。"""
        cover = self.doc.sections[0].header
        if self.variant == "undergraduate":
            container = self.doc.tables[0].cell(0, 0)
            cc = find_sdt(container._tc, "密级")
            paragraph = container.paragraphs[0]
            self.assertIsNone(find_sdt(cover._element, "密级"))
        else:
            cc = find_sdt(cover._element, "密级")
            paragraph = cover.paragraphs[0]
            self.assertIsNone(find_sdt(self.doc.element.body, "密级"))
        self.assertIsNotNone(cc)
        self.assertEqual("".join(t.text or "" for t in cc.iter(qn("w:t"))), "")
        self.assertEqual(paragraph.style.name, "封面密级")
        alignment = paragraph.alignment
        if alignment is None:
            alignment = style_para(paragraph.style, "alignment")
        if alignment is None:
            alignment = WD_ALIGN_PARAGRAPH.LEFT  # Word 省略默认左对齐属性
        self.assertEqual(alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self.assertNotEqual(cover.part.partname, self.doc.sections[1].header.part.partname)
        font_size = self.doc.styles["封面密级"].font.size
        size = font_size.pt if font_size is not None else self.rules["pageSetup"]["documentDefaults"]["characterFormat"]["fontSize"]["value"]
        self.assertEqual(size, 16 if self.variant == "graduate" else 14)

    def test_master_docx_baseline(self) -> None:
        """检查母版节数、表格数与内容控件数基线。"""
        self.assertTrue(self.docx_path.exists(), self.docx_path)
        doc = self.doc
        self.assertEqual(len(doc.sections), self.counts["sections"])
        self.assertEqual(len(doc.tables), self.counts["tables"])
        controls = len(doc.element.body.findall(".//" + qn("w:sdt")))
        controls += len(doc.sections[0].header._element.findall(".//" + qn("w:sdt")))
        self.assertEqual(controls, self.counts["content_controls"])
        self.assertGreaterEqual(len(doc.styles), 190)

        citation = doc.styles["文献引用"]
        self.assertEqual(citation.type, WD_STYLE_TYPE.CHARACTER)
        self.assertTrue(citation.font.superscript)
        inline_formula = doc.styles["行内公式"]
        self.assertEqual(inline_formula.type, WD_STYLE_TYPE.CHARACTER)
        self.assertEqual(inline_formula.font.name, "XITS Math")
        self.assertEqual(inline_formula.font.size, Pt(12))
        self.assertIsNotNone(doc.styles["公式"])
        self.assertIsNotNone(doc.styles["图注"])
        self.assertIsNotNone(doc.styles["表注"])

        # 公式示例：文档级数学字体须与规则一致（否则 Word 保存时会按默认
        # Cambria Math 给数学 run 注入直接字体，覆盖「行内公式」样式）。
        western = self.rules["styles"]["inlineFormula"]["characterFormat"]["fontFamily"]["western"]
        with ZipFile(self.docx_path) as zf:
            settings_root = ET.fromstring(zf.read("word/settings.xml"))
        math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        math_font = settings_root.find(f".//{{{math_ns}}}mathFont")
        self.assertIsNotNone(math_font, "settings.xml 缺少 m:mathFont")
        self.assertEqual(math_font.get(f"{{{math_ns}}}val"), western)

        # 第4章公式小节的 x、y、f(x) 必须是真实行内 OMath，并统一应用「行内公式」字符样式。
        body = self.doc.element.body
        omath_texts = [
            "".join(t.text or "" for t in o.iter(qn("m:t")))
            for o in body.xpath(".//m:oMath")
        ]
        for symbol in ("x", "y", "f(x)"):
            self.assertIn(symbol, omath_texts, f"{symbol} 应为行内公式对象")
        inline_id = self.doc.styles["行内公式"].style_id
        math_runs = body.xpath(".//m:r")
        self.assertTrue(math_runs)
        rstyles = body.xpath(".//m:r/w:rPr/w:rStyle")
        self.assertEqual(len(rstyles), len(math_runs),
                         "所有数学 run 须引用「行内公式」字符样式")
        for rstyle in rstyles:
            self.assertEqual(rstyle.get(qn("w:val")), inline_id)
        direct_fonts = [
            r for r in body.xpath(".//m:r/w:rPr/w:rFonts")
            if r.get(qn("w:ascii")) or r.get(qn("w:hAnsi"))
        ]
        self.assertEqual(direct_fonts, [],
                         "数学 run 不应带直接西文字体，以免覆盖样式")
    def test_page_footer_distance_matches_rules(self) -> None:
        """页脚距离按规则 1.5cm 生成：页码字符底端距页底约 1.9cm、基线约 2.0cm。"""
        rules = yaml.safe_load(
            self.rules_path.read_text(encoding="utf-8"))
        margins = rules["pageSetup"]["defaultPageMargins"]
        self.assertEqual(margins["footer"]["unit"], "cm")
        self.assertEqual(margins["footer"]["value"], 1.5)

        doc = self.doc
        # 封面无页码，页脚距离与正文一致取规则值。
        for section in doc.sections:
            self.assertAlmostEqual(section.footer_distance.cm, 1.5, delta=0.01)
        for section in doc.sections[2:]:
            self.assertAlmostEqual(section.header_distance.cm, 2.2, delta=0.01)
        with ZipFile(self.docx_path) as zf:
            document_xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn('w:footer="850"', document_xml)  # 1.5 厘米 = 850.4 缇
        self.assertNotIn('w:footer="964"', document_xml)

    def test_document_grid_disabled_in_all_sections(self) -> None:
        rules = yaml.safe_load(self.rules_path.read_text(encoding="utf-8"))
        self.assertEqual(rules["pageSetup"]["documentGrid"], "none")
        for path in (self.docx_path, self.dotm_path):
            with self.subTest(path=path.name), ZipFile(path) as zf:
                root = etree.fromstring(zf.read("word/document.xml"))
                ns = {"w": root.nsmap["w"]}
                sections = root.xpath("w:body/w:sectPr | w:body/w:p/w:pPr/w:sectPr", namespaces=ns)
                self.assertEqual(len(sections), len(self.doc.sections))
                for section in sections:
                    grid = section.find("w:docGrid", ns)
                    if grid is not None:
                        self.assertEqual(grid.get(qn("w:type")), "default")

    def test_numbered_example_chapters(self) -> None:
        """两变体的正文使用相同的四章结构。"""
        headings = [p.text.strip() for p in self.doc.paragraphs if p.style.name == "Heading 1"]
        self.assertEqual(headings, ["引言", "DOTM宏模板的使用", "DOCX模板的使用", "模板格式与使用范例"])
        mixed_space = re.compile(r"[\u3400-\u9fff] +[A-Za-z0-9]|[A-Za-z0-9] +[\u3400-\u9fff]")
        visible = ["".join(p.xpath("./w:r/w:t/text()")) for p in self.doc.element.body.xpath(".//w:p")]
        self.assertEqual([text for text in visible if mixed_space.search(text)], [])
        self.assertNotRegex("\n".join(visible), r"\b(?:STYLEREF|SEQ|MERGEFORMAT)\b")

    def test_style_index_covers_guide_styles(self) -> None:
        """第4章样式索引覆盖指南涉及、功能区未必有按钮的样式。"""
        doc = self.doc
        required = {
            "章标题（有编号）", "章标题（无编号）", "章标题（无编号、不进目录）",
            "一级节标题", "二级节标题", "三级节标题", "章内无编号小节",
            "名单页小节标题", "附录标题", "附录一级节标题", "附录二级节标题",
            "附录三级节标题", "正文", "表格文字", "图表附注", "图片", "图注",
            "子图注", "表注", "附录图注", "附录表注", "公式", "行内公式",
            "脚注文本", "页眉", "页脚", "目录一级条目", "目录二级条目", "目录三级条目", "符号说明条目",
            "参考文献条目", "参考文献条目（著者出版年）", "文献引用",
            "不明显强调", "强调", "明显强调", "行内代码", "行间代码",
            "列表编号1", "列表编号2", "列表编号3", "列表项目符号1", "列表项目符号2", "列表项目符号3", "封面论文题目", "封面学位类别",
            "封面作者信息", "封面密级", "封面日期", "英文封面题目", "英文封面信息",
            "英文封面导师信息", "授权说明标题", "授权说明文字",
        }
        indexed = set()
        for table in doc.tables:
            header = [cell.text.strip() for cell in table.rows[0].cells]
            if header[:2] != ["样式", "用途"]:
                continue
            for row in table.rows[1:]:
                indexed.add(row.cells[0].text.strip())
        self.assertEqual(sorted(required - indexed), [], "样式索引缺少样式")
        # 索引里的名称要能对上模板中实际存在的样式（标题样式按 Word 名对应）。
        docx_name = {
            "章标题（有编号）": "Heading 1",
            "一级节标题": "Heading 2",
            "二级节标题": "Heading 3",
            "三级节标题": "Heading 4",
            "正文": "Normal",
            "页眉": "Header",
            "页脚": "Footer",
            "脚注文本": "footnote text",
            "不明显强调": "Subtle Emphasis",
            "强调": "Emphasis",
            "明显强调": "Strong",
        }
        docx_name.update({spec["displayName"]: spec["name"] for spec in self.style_map.values()
                          if "displayName" in spec})
        available = {style.name for style in doc.styles}
        for name in required:
            self.assertIn(docx_name.get(name, name), available)
    def test_chapter_level_parts_start_with_next_page_sections(self) -> None:
        """所有章级标题前必须是“分节符（下一页）”，不能靠分页符或段前分页。"""
        doc = self.doc
        chapter_styles = {
            "Heading 1",
            "章标题（无编号）",
            "章标题（无编号、不进目录）",
            "附录标题",
        }
        for style_name in chapter_styles:
            style = doc.styles[style_name]
            pPr = style.element.pPr
            pb = pPr.find(qn("w:pageBreakBefore")) if pPr is not None else None
            self.assertIsNone(pb, f"{style_name} 不应设置段前分页")

        body = doc.element.body
        children = []
        for child in body.iterchildren():
            if child.tag == qn("w:p"):
                children.append(Paragraph(child, doc))
            else:
                children.append(child)

        headings = 0
        for i, obj in enumerate(children):
            if not isinstance(obj, Paragraph) or obj.style.name not in chapter_styles:
                continue
            headings += 1
            prev = children[i - 1] if i else None
            self.assertIsInstance(prev, Paragraph, f"{obj.text!r} 前缺少段落")
            pPr = prev._p.pPr
            self.assertIsNotNone(pPr, f"{obj.text!r} 前一段没有段属性")
            self.assertIsNotNone(
                pPr.find(qn("w:sectPr")),
                f"{obj.text!r} 前不是分节符（下一页）",
            )

        self.assertGreaterEqual(headings, self.counts["chapter_headings"])
        page_breaks = [
            br for br in body.iter(qn("w:br")) if br.get(qn("w:type")) == "page"
        ]
        self.assertEqual(page_breaks, [], "正文中不应再有普通分页符")
    def test_ribbon_callbacks_exist_in_vba(self) -> None:
        """检查功能区回调在 VBA 工程中均有对应入口。"""
        self.assertTrue(self.dotm_path.exists(), "缺少 dotm；先运行 src/build_dotm.py")
        with ZipFile(self.dotm_path) as zf:
            names = zf.namelist()
            self.assertIn("word/vbaProject.bin", names)
            custom_ui = next(
                (n for n in names if n.endswith("customUI14.xml")), None
            )
            self.assertIsNotNone(custom_ui, names)
            self.assertIn("[Content_Types].xml", names)
            self.assertIn("_rels/.rels", names)
            ET.fromstring(zf.read(custom_ui))

        vba_source = "\n".join(
            p.read_text(encoding="utf-8") for p in (ROOT / "src" / "dotm" / "vba").glob("*.bas")
        )
        tree = ET.parse(self.ribbon_path)
        callbacks = set()
        for element in tree.iter():
            for key in ("onAction",):
                if key in element.attrib:
                    callbacks.add(element.attrib[key])
        self.assertTrue(callbacks, "Ribbon XML 没有任何回调")
        for callback in sorted(callbacks):
            self.assertRegex(
                vba_source,
                rf"Public\s+(Sub|Function)\s+{re.escape(callback)}\b",
                f"Ribbon 回调 {callback} 没有对应的 Public 过程",
            )
    def test_list_numbering_hanging_indent(self) -> None:
        """检查列表样式按层级设置左缩进，并统一使用 2 字符悬挂缩进。"""
        doc = self.doc
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        with ZipFile(self.docx_path) as zf:
            root = ET.fromstring(zf.read("word/numbering.xml"))
        # 一级不设左缩进；二、三级分别设 2、4 字符（1 字符 = 240 缇）。
        expected_left = {
            "List Number": None,
            "List Number 2": ("200", "480"),
            "List Number 3": ("400", "960"),
            "List Bullet": None,
            "List Bullet 2": ("200", "480"),
            "List Bullet 3": ("400", "960"),
        }
        for style_name, left in expected_left.items():
            style = doc.styles[style_name]
            style_ind = style.element.find(f"{qn('w:pPr')}/{qn('w:ind')}")
            self.assertIsNotNone(style_ind, style_name)
            if left is None:
                self.assertIsNone(style_ind.get(qn("w:leftChars")), style_name)
                self.assertIsNone(style_ind.get(qn("w:left")), style_name)
            else:
                self.assertEqual(style_ind.get(qn("w:leftChars")), left[0], style_name)
                self.assertEqual(style_ind.get(qn("w:left")), left[1], style_name)
            self.assertEqual(style_ind.get(qn("w:hangingChars")), "200", style_name)
            self.assertEqual(style_ind.get(qn("w:hanging")), "480", style_name)
            self.assertIsNone(style_ind.get(qn("w:firstLineChars")), style_name)

            num_id = style.element.find(
                f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}"
            ).get(qn("w:val"))
            num = root.find(f".//w:num[@w:numId='{num_id}']", ns)
            abs_id = num.find("w:abstractNumId", ns).get(qn("w:val"))
            abs_el = root.find(f".//w:abstractNum[@w:abstractNumId='{abs_id}']", ns)
            lvl = abs_el.find("w:lvl", ns)
            ind = lvl.find("w:pPr/w:ind", ns)
            self.assertIsNone(ind.get(qn("w:leftChars")), style_name)
            self.assertIsNone(ind.get(qn("w:left")), style_name)
            self.assertEqual(ind.get(qn("w:hangingChars")), "200", style_name)
            self.assertEqual(ind.get(qn("w:hanging")), "480", style_name)
            self.assertIsNone(ind.get(qn("w:firstLineChars")), style_name)
    def test_subfigure_caption_style(self) -> None:
        """检查子图注继承图注样式及段后间距。"""
        doc = self.doc
        style = doc.styles["子图注"]
        self.assertEqual(style.base_style.name, "图注")
        space_after = style.paragraph_format.space_after
        if space_after is None:
            space_after = style.base_style.paragraph_format.space_after
        self.assertEqual(space_after.pt, 12.0)
    def test_ribbon_ids_unique(self) -> None:
        """检查功能区控件 id 不重复。"""
        tree = ET.parse(self.ribbon_path)
        ids = [el.attrib["id"] for el in tree.iter() if "id" in el.attrib]
        self.assertTrue(ids)
        self.assertEqual(len(ids), len(set(ids)), "Ribbon id 重复")

    def test_export_menu(self) -> None:
        ns = {"m": "http://schemas.microsoft.com/office/2009/07/customui"}
        root = ET.parse(self.ribbon_path).getroot()
        group = root.find(".//m:group[@id='grpExport']", ns)
        self.assertIsNotNone(group)
        self.assertEqual(group.get("label"), "导出")
        menu, = group
        self.assertEqual(menu.tag, f"{{{ns['m']}}}menu")
        self.assertEqual(menu.get("label"), "导出")
        self.assertEqual(menu.get("imageMso"), "FileSaveAs")
        self.assertEqual(
            [(item.get("label"), item.get("onAction"), item.get("imageMso")) for item in menu],
            [("导出为 DOCX", "ExportDocx", "FileSaveAsWordDocx"),
             ("导出为 PDF", "ExportPdf", "FileSaveAsPdfOrXps")],
        )

    def test_footnote_marker_format(self) -> None:
        """脚注文本编号为普通文本，正文脚注标记仍为上标。"""
        with ZipFile(self.dotm_path) as zf:
            styles = zf.read("word/styles.xml").decode("utf-8")
            document = zf.read("word/document.xml").decode("utf-8")
            footnotes = zf.read("word/footnotes.xml").decode("utf-8")

        style = re.search(
            r'<w:style [^>]*w:styleId="[^"]*"[^>]*>\s*<w:name w:val="footnote reference".*?</w:style>',
            styles,
            re.S,
        )
        self.assertIsNotNone(style)
        self.assertNotIn("w:vertAlign", style.group(0))

        text_style = re.search(
            r'<w:style [^>]*w:styleId="[^"]*"[^>]*>\s*<w:name w:val="footnote text".*?</w:style>',
            styles,
            re.S,
        )
        self.assertIsNotNone(text_style)
        text_xml = text_style.group(0)
        # Word 在 SaveAs2 时会去掉值为 0 的 w:after（从 docDefaults 继承 0），
        # 所以属性缺省与 "0" 都表示段后 0 磅。
        self.assertNotRegex(text_xml, r'w:after="(?!0")\d+"')
        self.assertIn('w:hanging="270"', text_xml)
        self.assertIn('w:hangingChars="150"', text_xml)
        self.assertNotIn('w:leftChars="150"', text_xml)
        self.assertNotRegex(text_xml, r'w:left="(?!0")\d+"')

        first_note = re.search(r'<w:footnote [^>]*w:id="1".*?</w:footnote>', footnotes, re.S)
        self.assertIsNotNone(first_note)
        note_xml = first_note.group(0)
        self.assertNotIn("<w:ind", note_xml)
        self.assertNotIn("<w:spacing", note_xml)

        body_ref = re.search(r'<w:footnoteReference[^>]*/>', document)
        self.assertIsNotNone(body_ref)
        body_context = document[max(0, body_ref.start() - 260):body_ref.end()]
        self.assertIn('w:vertAlign w:val="superscript"', body_context)

        note_ref = re.search(r'<w:footnoteRef\s*/>', footnotes)
        self.assertIsNotNone(note_ref)
        note_context = footnotes[max(0, note_ref.start() - 260):note_ref.end()]
        self.assertNotIn('w:vertAlign w:val="superscript"', note_context)
        note_after = footnotes[note_ref.end():note_ref.end() + 260]
        self.assertIn("\u2002</w:t>", note_after)
        self.assertNotIn('xml:space="preserve"> </w:t>', note_after)


    def test_emphasis_and_keyword_styles(self) -> None:
        """强调样式、关键词、目录标题、文献引用都应复用样式而非直接格式。"""
        doc = self.doc
        for name, font, bold in (
            ("Subtle Emphasis", "Arial", False),
            ("Emphasis", "Times New Roman", True),
            ("Strong", "Arial", True),
        ):
            style = doc.styles[name]
            self.assertEqual(style.font.name, font)
            self.assertEqual(bool(style.font.bold), bold)
            self.assertFalse(bool(style.font.italic))

        kw = next(p for p in doc.paragraphs if p.text.startswith("关键词："))
        self.assertEqual(kw.runs[0].style.name, "Subtle Emphasis")
        en = next(p for p in doc.paragraphs if p.text.startswith("Keywords:"))
        self.assertEqual(en.runs[0].style.name, "Strong")

        for unused in ("名单页名单行", "名单页答辩行"):
            self.assertNotIn(unused, [st.name for st in doc.styles])

        toc_title = next(p for p in doc.paragraphs if p.text.strip() == "目　录")
        self.assertEqual(toc_title.style.name, "章标题（无编号）")
        outline = toc_title._p.xpath("./w:pPr/w:outlineLvl/@w:val")
        toc_rule = self.rules["documentParts"]["tableOfContents"]["heading"]
        self.assertEqual(outline, ["9"] if toc_rule.get("outlineLevel") == "bodyText" else [])

        cites = [r for p in doc.paragraphs for r in p.runs
                 if r.style is not None and r.style.name == "文献引用"]
        self.assertTrue(cites)
        self.assertFalse(cites[0]._r.xpath("./w:rPr/w:vertAlign"))

    def test_document_defaults(self) -> None:
        """文档默认：宋体/Times New Roman 小四，段前后 0、行距最小值 20 磅。"""
        styles_xml = read_part(self.docx_path, "word/styles.xml")
        defaults = re.search(r"<w:docDefaults>.*?</w:docDefaults>", styles_xml, re.S).group(0)
        fonts = re.search(r"<w:rFonts[^/]*/>", defaults).group(0)
        self.assertIn('w:ascii="Times New Roman"', fonts)
        self.assertIn('w:eastAsia="宋体"', fonts)
        self.assertIn('w:hint="eastAsia"', fonts)
        self.assertIn('<w:sz w:val="24"/>', defaults)
        ppr = re.search(r"<w:pPrDefault>.*?</w:pPrDefault>", defaults, re.S).group(0)
        spacing = re.search(r"<w:spacing[^/]*/>", ppr).group(0)
        self.assertIn('w:line="400"', spacing)
        self.assertIn('w:lineRule="atLeast"', spacing)
        for attr in ("before", "after"):
            value = spacing_attr(spacing, attr)
            self.assertIn(value, (None, "0"), f"文档默认段{attr}应为 0")

    def test_hyperlink_is_black_without_underline(self) -> None:
        """超链接沿用默认文字：黑色、无下划线、无主题色。"""
        block = style_block(self.dotm_styles, "Hyperlink")
        self.assertTrue(block, "dotm 缺少 Hyperlink 样式")
        self.assertIn('<w:color w:val="000000"/>', block)
        self.assertNotIn("themeColor", block)
        self.assertIn('<w:u w:val="none"/>', block)

    def test_footnote_reference_style(self) -> None:
        """脚注引用样式只管字形，上标由正文直接格式提供。"""
        block = style_block(self.dotm_styles, "footnote reference")
        self.assertTrue(block, "dotm 缺少 footnote reference 样式")
        self.assertIn('<w:color w:val="000000"/>', block)
        self.assertNotIn("vertAlign", block)

    def test_emphasis_styles_are_black(self) -> None:
        """三个强调字符样式都固定黑色。"""
        for name in ("Subtle Emphasis", "Emphasis", "Strong"):
            block = style_block(self.dotm_styles, name)
            self.assertIn('<w:color w:val="000000"/>', block, f"{name} 应为黑色")

    def test_three_line_table_style(self) -> None:
        """三线表表格样式：上下框线 + 首行条件底边线。"""
        block = style_block(self.dotm_styles, "三线表")
        self.assertIn('w:type="table"', block)
        self.assertIn("<w:tblBorders>", block)
        self.assertIn('w:type="firstRow"', block)

    def test_keep_next_for_unnumbered_headings(self) -> None:
        """章内无编号小节与图片要和下段同页；Word 存盘会剥掉，构建需补回。"""
        if not self.dotm_path.exists():
            self.skipTest(f"缺少 {self.variant} dotm：{self.dotm_path}")
        docx_styles = read_part(self.docx_path, "word/styles.xml")
        for name in ("章内无编号小节", "图片"):
            for source, xml in (("docx", docx_styles), ("dotm", self.dotm_styles)):
                block = style_block(xml, name)
                self.assertIn("<w:keepNext/>", block, f"{source} 的「{name}」缺少 keepNext")

    def test_all_rule_styles_exist_by_name(self) -> None:
        """规则里的每个样式键都能在产物中按名称找到。"""
        available = style_names(read_part(self.docx_path, "word/styles.xml"))
        missing = [
            f"{key}→{spec['name']}"
            for key, spec in self.style_map.items()
            if spec["name"].casefold() not in available
        ]
        self.assertEqual(missing, [], f"{self.variant} 产物缺少样式")
        self.assertEqual(set(self.style_map), set(self.rules["styles"]))

    def test_recommended_styles_and_gallery(self) -> None:
        implementation = yaml.safe_load(
            (ROOT / "configs" / "implementation-map.yaml").read_text(encoding="utf-8")
        )
        groups = implementation["stylePresentation"]
        names = {
            implementation["styleMap"][key]["name"].casefold(): (group["priority"], group.get("quickStyle", False))
            for group in groups for key in group["styles"] if key in self.rules["styles"]
        }
        self.assertEqual(len(names), 60)
        self.assertEqual(sum(gallery for _, gallery in names.values()), 22)
        for level in ("", " 2", " 3"):
            for family in ("list number", "list bullet"):
                self.assertEqual(names[family + level], (25, True))

        for path in (self.docx_path, self.dotm_path):
            with self.subTest(path=path.name):
                root = ET.fromstring(read_part(path, "word/styles.xml"))
                configured_priorities = {}
                other_priorities = []
                for style in root.findall(qn("w:style")):
                    name = style.find(qn("w:name")).get(qn("w:val")).casefold()
                    for spec in self.style_map.values():
                        if spec["name"].casefold() == name and "displayName" in spec:
                            self.assertEqual(style.find(qn("w:aliases")).get(qn("w:val")), spec["displayName"])
                            self.assertNotEqual(style.get(qn("w:customStyle")), "1")
                    actual = int(style.find(qn("w:uiPriority")).get(qn("w:val")))
                    self.assertIn(actual, range(100))
                    self.assertEqual(style.find(qn("w:unhideWhenUsed")).get(qn("w:val")), "0")
                    self.assertEqual(style.find(qn("w:semiHidden")).get(qn("w:val")), str(int(name not in names)))
                    preferred, gallery = names.get(name, (None, False))
                    self.assertEqual(style.find(qn("w:qFormat")).get(qn("w:val")), str(int(gallery)))
                    if preferred is None:
                        other_priorities.append(actual)
                    else:
                        self.assertEqual(configured_priorities.setdefault(preferred, actual), actual)

                ordered = [configured_priorities[key] for key in sorted(configured_priorities)]
                self.assertEqual(ordered, sorted(set(ordered)))
                self.assertGreater(min(other_priorities), ordered[-1])
                self.assertGreater(len(set(other_priorities)), 1)

                latent = root.find(qn("w:latentStyles"))
                self.assertEqual(latent.get(qn("w:defSemiHidden")), "1")
                self.assertEqual(latent.get(qn("w:defUnhideWhenUsed")), "0")
                self.assertEqual(latent.get(qn("w:defQFormat")), "0")
                self.assertGreater(int(latent.get(qn("w:defUIPriority"))), ordered[-1])
                for entry in latent.findall(qn("w:lsdException")):
                    if entry.get(qn("w:name")).casefold() not in names:
                        self.assertEqual(entry.get(qn("w:semiHidden")), "1")
                        self.assertEqual(entry.get(qn("w:unhideWhenUsed")), "0")
                        self.assertEqual(entry.get(qn("w:qFormat")), "0")
                        self.assertGreater(int(entry.get(qn("w:uiPriority"))), ordered[-1])

    def test_implementation_map_does_not_pin_style_ids(self) -> None:
        """样式映射只写名称：Word 保存后会重排 styleId。"""
        for key, spec in self.style_map.items():
            self.assertNotIn("styleId", spec, f"{key} 不应写死 styleId")
            self.assertTrue(spec.get("name"), f"{key} 缺少样式名")

    def test_unnumbered_headings_have_no_numbering(self) -> None:
        """章内无编号小节不带自动编号。"""
        for name in ("章内无编号小节", "名单页小节标题"):
            xml = self.doc.styles[name].element.xml
            self.assertNotIn('<w:numId w:val="1"', xml)


class GraduateStaticTest(_SharedStaticMixin, ProductCase):
    """研究生变体的静态格式与样式契约检查。"""
    variant = "graduate"


class UndergraduateStaticTest(_SharedStaticMixin, ProductCase):
    """本科变体的静态格式与样式契约检查。"""
    variant = "undergraduate"


class GraduateOnlyTest(ProductCase):
    """只在研究生模板里成立的结构：STYLEREF 页眉、完整封面、功能区布局、后置部分。"""

    variant = "graduate"

    def test_committee_tables_use_table_text(self) -> None:
        """名单页的 3 张答辩/评阅表格单元格统一套“表格文字”。"""
        doc = self.doc
        committee = [t for t in doc.tables
                     if t.style.name == "Normal Table"
                     and len(t.rows) == 4 and len(t.columns) in (3, 4)]
        self.assertEqual(len(committee), 3)
        for t in committee:
            for cell in t._cells:
                self.assertEqual(cell.paragraphs[0].style.name, "表格文字")
    def test_header_number_and_title_have_fullwidth_space(self) -> None:
        """检查研究生页眉章号与标题间用全角空格。"""
        texts = []
        with ZipFile(self.docx_path) as zf:
            for name in zf.namelist():
                if not (name.startswith("word/header") and name.endswith(".xml")):
                    continue
                xml = zf.read(name).decode("utf-8")
                texts.extend(re.findall(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", xml, re.S))
        joined = "".join(texts)
        self.assertRegex(joined, r"第1章　\S")
        self.assertRegex(joined, r"附录A　\S")
        self.assertNotIn("第1章引言", joined)
        self.assertNotIn("附录A模板检查清单", joined)
    def test_cover_direct_layout_is_complete(self) -> None:
        """直接构建的封面应完成默认排版，并带好级联所需的内容控件数据。"""
        doc = self.doc
        config = load_degree_config()
        from src.build.degree_catalog import discipline_options

        # --- 作者信息单元格：4 条必填行 + 2 个空行，标签宽度不变 ---------
        zh = doc.tables[0].rows[2].cells[0]
        self.assertEqual(len(zh.paragraphs), 6)
        expected = [
            "培养单位",
            "学科（专业学位改为专业领域）",
            "研究生（专业学位改为申请人）",
            "指导教师",
        ]
        tags = [
            sdt.find(".//" + qn("w:tag")).get(qn("w:val"))
            for sdt in zh._tc.findall(".//" + qn("w:sdt"))
        ]
        self.assertEqual(tags, expected)
        for paragraph in zh.paragraphs[:4]:
            fit = paragraph.runs[1]._r.rPr.find(qn("w:fitText"))
            self.assertIsNotNone(fit, paragraph.text)
            self.assertEqual(fit.get(qn("w:val")), "1280")
        for paragraph in zh.paragraphs[4:]:
            self.assertEqual(paragraph.text, "")

        # --- 题目与学位申请句同格：申请句紧接题目，不在单独行 ------------
        title_cell = doc.tables[0].rows[0].cells[0]
        self.assertEqual(len(title_cell.paragraphs), 2)
        self.assertIsNotNone(find_sdt(title_cell.paragraphs[0]._p, "论文题目"))

        # --- 学位类别：原生下拉列表，候选来自学位目录 --------------------
        degree_para = title_cell.paragraphs[1]
        degree = find_sdt(degree_para._p, "学位类别")
        self.assertIsNotNone(degree)
        self.assertIsNotNone(degree.find(".//" + qn("w:dropDownList")))
        items = degree.findall(".//" + qn("w:listItem"))
        self.assertEqual(len(items), sum(bool(item.get(key)) for group in ("academic", "professional")
                                        for item in load_degree_config()[group] for key in ("master_en", "doctor_en")))
        displays = [item.get(qn("w:displayText")) for item in items]
        self.assertIn("工学硕士学位", displays)
        self.assertIn("金融硕士专业学位", displays)
        self.assertIsNone(degree.find(".//" + qn("w:showingPlcHdr")))

        # --- 学科：包含本门类及跨门类可授工学学位的候选 ------------------
        discipline = find_sdt(zh._tc, "学科（专业学位改为专业领域）")
        self.assertIsNotNone(discipline)
        self.assertIsNotNone(discipline.find(".//" + qn("w:comboBox")))
        self.assertIsNone(discipline.find(".//" + qn("w:showingPlcHdr")))
        self.assertEqual(
            len(discipline.findall(".//" + qn("w:listItem"))),
            len(discipline_options(config, "工学")),
        )

        # --- 成文日期：中文日期控件就绪，无占位提示 -----------------------
        date = find_sdt(doc.tables[0].rows[4].cells[0]._tc, "成文日期")
        self.assertIsNotNone(date)
        self.assertIsNone(date.find(".//" + qn("w:showingPlcHdr")))

        # --- 英文导师段：制表位与内容控件 --------------------------------
        en = doc.tables[1].rows[5].cells[0].paragraphs[0]
        # Word 会把“零缩进”保存成无 w:ind；有效值等价于 0。
        self.assertTrue(en.paragraph_format.left_indent in (None, Pt(0)))
        tabs = {(t.position.pt, t.alignment) for t in en.paragraph_format.tab_stops}
        self.assertIn((195.5, WD_TAB_ALIGNMENT.RIGHT), tabs)
        self.assertIn((200.5, WD_TAB_ALIGNMENT.LEFT), tabs)
        sdt = en._p.find(".//" + qn("w:sdt"))
        self.assertIsNotNone(sdt)
        self.assertIsNone(sdt.find(".//" + qn("w:showingPlcHdr")))

        # --- 英文成文日期与英文学位名称：控件就绪，无占位提示 ------------
        en_date = find_sdt(doc.tables[1].rows[7].cells[0]._tc, "英文成文日期")
        self.assertIsNotNone(en_date)
        self.assertIsNone(en_date.find(".//" + qn("w:showingPlcHdr")))
        en_degree = find_sdt(doc.tables[1].rows[2].cells[0]._tc, "英文学位名称")
        self.assertIsNotNone(en_degree)
        self.assertIsNone(en_degree.find(".//" + qn("w:showingPlcHdr")))
    def test_ribbon_interaction_layout(self) -> None:
        """检查研究生功能区交互分组布局。"""
        tree = ET.parse(self.ribbon_path)
        ns = {"m": "http://schemas.microsoft.com/office/2009/07/customui"}
        tabs = list(tree.getroot().iter(f"{{{ns['m']}}}tab"))
        self.assertEqual(len(tabs), 1)
        groups = tabs[0].findall("m:group", ns)
        self.assertEqual(
            [g.attrib["label"] for g in groups],
            ["封面", "段落", "字体", "列表", "图片", "表格", "公式", "代码", "引用", "导出", "帮助"],
        )

        # 封面组：刷新封面（large 按钮）+ 选填信息（large 菜单，内含 3 个选填项）。
        cover_group = groups[0]
        self.assertEqual(cover_group.attrib.get("label"), "封面")
        self.assertEqual(
            [el.attrib.get("id") for el in cover_group],
            ["btnRefreshCover", "mnuOptionalInfo"],
        )
        self.assertEqual(
            [el.attrib.get("label") for el in cover_group],
            ["刷新封面", "选填信息"],
        )
        refresh_button, optional_menu = list(cover_group)
        self.assertEqual(refresh_button.attrib.get("size"), "large")
        self.assertEqual(refresh_button.attrib.get("imageMso"), "GroupAddins")
        self.assertEqual(optional_menu.attrib.get("size"), "large")
        self.assertEqual(optional_menu.attrib.get("imageMso"), "FilePrepareMenu")
        menu_buttons = list(optional_menu)
        self.assertEqual(
            [el.attrib.get("id") for el in menu_buttons],
            ["btnAssocSupervisor", "btnCoSupervisor", "btnSecurity"],
        )
        self.assertEqual(
            [el.attrib.get("label") for el in menu_buttons],
            ["副导师行", "联合导师行", "密级信息"],
        )
        for item in menu_buttons:
            self.assertNotIn("size", item.attrib, item.attrib.get("id"))
        cover_labels = [el.attrib.get("label", "") for el in cover_group.iter()]
        self.assertNotIn("检查封面", cover_labels)
        for label in ("刷新封面", "选填信息", "副导师行", "联合导师行", "密级信息"):
            self.assertIn(label, cover_labels)

        # 段落组：正文章节 large 菜单，附录章节与其他章节为 normal 小菜单。
        structure_group = next(g for g in groups if g.attrib["label"] == "段落")
        self.assertEqual(
            [el.attrib.get("id") for el in structure_group],
            ["mnuChapter", "mnuAppendix", "mnuOtherSection", "btnBodyText"],
        )
        chapter_menu, appendix_menu, other_menu, body_button = list(structure_group)
        self.assertEqual(chapter_menu.attrib.get("size"), "large")
        self.assertEqual(appendix_menu.attrib.get("size"), "normal")
        self.assertEqual(other_menu.attrib.get("size"), "normal")
        self.assertNotIn("size", body_button.attrib)

        # 公式组：行内公式（large 按钮）+ 行间公式菜单（large，内含 有编号/无编号）。
        formula_group = next(g for g in groups if g.attrib["label"] == "公式")
        self.assertEqual(
            [el.attrib.get("id") for el in formula_group],
            ["btnInlineFormula", "mnuDisplayFormula"],
        )
        self.assertEqual(
            [el.attrib.get("label") for el in formula_group],
            ["行内公式", "行间公式"],
        )
        inline_btn, display_menu = list(formula_group)
        self.assertEqual(inline_btn.attrib.get("size"), "large")
        self.assertEqual(display_menu.attrib.get("size"), "large")
        self.assertEqual(display_menu.attrib.get("imageMso"), "TableFormulaDialog")
        display_items = list(display_menu)
        self.assertEqual(
            [el.attrib.get("id") for el in display_items],
            ["btnDisplayFormulaNumbered", "btnDisplayFormulaPlain"],
        )
        self.assertEqual(
            [el.attrib.get("label") for el in display_items],
            ["有编号", "无编号"],
        )
        for item in display_items:
            self.assertNotIn("size", item.attrib, item.attrib.get("id"))
        all_ribbon_ids = {el.attrib.get("id", "") for el in tree.iter()}
        self.assertNotIn("btnDisplayFormula", all_ribbon_ids)

        # 列表组：调整编号（小菜单，内含 继续编号/重新开始于1）放在最后；
        # 参考文献小菜单与“缩略语”按钮在前。
        list_group = next(g for g in groups if g.attrib["label"] == "列表")
        self.assertEqual(
            [el.attrib.get("id") for el in list_group],
            [
                "mnuOrderedList",
                "mnuBulletList",
                "mnuReferenceFormat",
                "btnSymbolEntry",
                "mnuAdjustNumbering",
            ],
        )
        symbol_button = next(
            el for el in list_group if el.attrib.get("id") == "btnSymbolEntry"
        )
        self.assertEqual(symbol_button.attrib.get("label"), "缩略语")
        self.assertEqual(symbol_button.attrib.get("imageMso"), "SymbolInsertGallery")
        self.assertEqual(symbol_button.attrib.get("onAction"), "InsertSymbolEntry")
        self.assertNotIn("size", symbol_button.attrib)
        adjust_menu = next(
            el for el in list_group if el.attrib.get("id") == "mnuAdjustNumbering"
        )
        self.assertEqual(adjust_menu.attrib.get("size"), "normal")
        self.assertEqual(adjust_menu.attrib.get("imageMso"), "NumberingRestart")
        adjust_items = list(adjust_menu)
        self.assertEqual(
            [el.attrib.get("id") for el in adjust_items],
            ["btnListContinue", "btnListRestart"],
        )
        self.assertEqual(
            [el.attrib.get("label") for el in adjust_items],
            ["继续编号", "重新开始于1"],
        )
        for item in adjust_items:
            self.assertNotIn("size", item.attrib, item.attrib.get("id"))

        labels = {el.attrib.get("label", "") for el in tree.iter()}
        for removed in ("备份", "检查", "封面信息"):
            self.assertNotIn(removed, labels)
        for added in ("插入图片", "插入表格", "行内代码", "行间代码", "正文", "有序列表", "无序列表", "调整编号", "缩略语", "子图注", "行内公式", "行间公式", "不明显强调", "强调", "明显强调"):
            self.assertIn(added, labels)

        # 触发“类型不匹配”的 getEnabled 回调已全部移除。
        self.assertFalse(any("getEnabled" in el.attrib for el in tree.iter()))

        self.assertNotIn("mnuSubfigureNumbering", {el.attrib.get("id", "") for el in tree.iter()})
        self.assertNotIn("引用图表公式", labels)
        self.assertIn("引用内容", labels)

        ref_group = next(g for g in groups if g.attrib["label"] == "引用")
        ref_ids = [el.attrib.get("id") for el in ref_group]
        self.assertLess(ref_ids.index("mnuCrossRef"), ref_ids.index("btnCitation"))
        self.assertLess(ref_ids.index("btnCitation"), ref_ids.index("btnFootnote"))
        self.assertLess(ref_ids.index("btnFootnote"), ref_ids.index("btnUpdateFields"))
        self.assertEqual(ref_group.attrib.get("label"), "引用")
        sizes = {el.attrib.get("id"): el.attrib.get("size") for el in tree.iter()}
        self.assertEqual(sizes["btnRefreshCover"], "large")
        self.assertEqual(sizes["mnuOptionalInfo"], "large")
        self.assertEqual(sizes["mnuChapter"], "large")
        self.assertEqual(sizes["mnuAppendix"], "normal")
        self.assertEqual(sizes["mnuAdjustNumbering"], "normal")
        # 菜单内三项不写 size，Word 取默认（normal）。
        for normal in ("btnAssocSupervisor", "btnCoSupervisor", "btnSecurity"):
            self.assertIsNone(sizes.get(normal), normal)
        self.assertEqual(sizes["btnSubfigureCaption"], "normal")
        for big in ("mnuCrossRef", "btnUpdateFields", "btnFootnote", "btnAbout"):
            self.assertEqual(sizes[big], "large")

        # Word 不接受带 menu 的 splitButton；一旦出现会让整个 customUI 静默不加载。
        self.assertFalse(
            list(tree.iter(f"{{{ns['m']}}}splitButton")),
            "customUI 不能使用 splitButton，Word 会拒绝整个功能区",
        )

        # 空控件必须写成自闭合 <button .../>。写成 <button ...></button>（哪怕只含空白）
        # 会让 Word 丢弃整个 customUI，表现为“论文”选项卡消失。
        for button in tree.iter(f"{{{ns['m']}}}button"):
            button_id = button.attrib.get("id")
            self.assertIsNone(button.text, button_id)
            self.assertEqual(len(list(button)), 0, button_id)
    def test_p0_p1_document_fixes(self) -> None:
        """标题段距、编号格式、公式域、参考文献条目、无编号小节。"""
        doc = self.doc
        for name, before, after in (
            ("Heading 2", 24.0, 6.0),
            ("Heading 3", 12.0, 6.0),
            ("Heading 4", 12.0, 6.0),
        ):
            style = doc.styles[name]
            self.assertAlmostEqual(style_para(style, "space_before").pt, before, msg=name)
            self.assertAlmostEqual(style_para(style, "space_after").pt, after, msg=name)

        with ZipFile(self.docx_path) as zf:
            settings = zf.read("word/settings.xml").decode("utf-8")
        self.assertIn('w:numFmt w:val="decimalEnclosedCircle"', settings)
        self.assertIn('w:numRestart w:val="eachPage"', settings)

        formula = next(
            p for p in doc.paragraphs
            if p.style.name == "公式" and p._p.xpath(".//m:oMath")
        )
        self.assertTrue(formula._p.xpath(".//m:oMathPara"), "公式必须是块级 m:oMathPara")
        self.assertTrue(formula._p.xpath(".//m:eqArr"), "示例公式应使用 #() 编号布局")
        xml = formula._p.xml
        self.assertIn("STYLEREF 1", xml)
        self.assertIn("SEQ", xml)
        self.assertIn("式", xml)
        self.assertGreaterEqual(len(formula._p.xpath(".//w:fldChar")), 6)
        self.assertEqual(len(formula._p.xpath(".//w:tab")), 0, "公式不应使用制表符")
        self.assertFalse(formula.text.startswith("	"))

        ref_style = doc.styles["参考文献条目"]
        self.assertIsNotNone(ref_style.element.find(".//" + qn("w:numPr")))
        citation_run = next(
            run for p in doc.paragraphs for run in p.runs
            if run.text == "[1]" and run.style.name == "文献引用"
        )
        self.assertTrue(citation_run.font.superscript or doc.styles["文献引用"].font.superscript)

        # 10 项修订：无编号居中节标题、密级默认空、附录题注独立样式。
        unnum = doc.styles["章内无编号小节"]
        pPr = unnum.element.pPr
        self.assertEqual(pPr.find(qn("w:outlineLvl")).get(qn("w:val")), "9")
        self.assertEqual(unnum.paragraph_format.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(pPr.find(".//" + qn("w:numId")).get(qn("w:val")), "0")
        self.assertIsNotNone(doc.styles["附录图注"])
        self.assertIsNotNone(doc.styles["附录表注"])
        security_texts = []
        for sdt in doc.sections[0].header._element.iter(qn("w:sdt")):
            tag = sdt.find(".//" + qn("w:tag"))
            if tag is not None and tag.get(qn("w:val")) == "密级":
                security_texts.append("".join(t.text or "" for t in sdt.iter(qn("w:t"))))
        self.assertEqual(security_texts, [""])


UNDERGRADUATE = product_paths("undergraduate")
GRADUATE = product_paths("graduate")


class UndergraduateOnlyTest(ProductCase):
    """本科与研究生真正不同的结构。"""

    variant = "undergraduate"

    def test_cover_margins_and_gutter(self) -> None:
        """封面页边距按本科规范，装订线只对封面生效。"""
        cover = self.doc.sections[0]
        self.assertAlmostEqual(cover.top_margin.cm, 3.8, places=1)
        self.assertAlmostEqual(cover.bottom_margin.cm, 3.2, places=1)
        self.assertAlmostEqual(cover.left_margin.cm, 3.0, places=1)
        self.assertAlmostEqual(cover.right_margin.cm, 3.0, places=1)
        self.assertAlmostEqual(cover.header_distance.cm, 1.7, places=1)
        self.assertAlmostEqual(cover.footer_distance.cm, 1.5, places=1)
        self.assertAlmostEqual(cover.gutter.cm, 0.2, places=1)
        self.assertEqual(cover.header.paragraphs[0].style.style_id, self.doc.styles["Header"].style_id)
        table_width = self.doc.tables[0].columns[0].width
        body_width = cover.page_width - cover.left_margin - cover.right_margin - cover.gutter
        self.assertAlmostEqual(table_width / 360000, body_width / 360000, delta=0.01)
        self.assertAlmostEqual(self.doc.tables[0].rows[0].height.pt, 18.2, delta=0.1)
        for index, section in enumerate(self.doc.sections[1:], start=1):
            self.assertAlmostEqual(section.gutter.cm, 0.0, places=2,
                                   msg=f"第 {index} 节不应带装订线")

    def test_headers_are_empty(self) -> None:
        """本科页眉为“无”：任何一节都没有页眉文字与域。"""
        for index, section in enumerate(self.doc.sections):
            for header in (section.header, section.even_page_header, section.first_page_header):
                text = "".join(p.text for p in header.paragraphs).strip()
                self.assertEqual(text, "", f"第 {index} 节页眉不应有文字")
                instr = "".join(
                    t.text or "" for p in header.paragraphs for t in p._p.iter(qn("w:instrText"))
                )
                self.assertNotIn("STYLEREF", instr)

    def test_page_numbers_start_after_front_matter(self) -> None:
        """前两节（封面）不排页码，正文起每节页脚有 PAGE 域且居中。"""
        for index, section in enumerate(self.doc.sections):
            paragraph = section.footer.paragraphs[0]
            instr = "".join(t.text or "" for t in paragraph._p.iter(qn("w:instrText")))
            if index < 2:
                self.assertNotIn("PAGE", instr, f"第 {index} 节不应有页码")
                continue
            self.assertIn("PAGE", instr, f"第 {index} 节缺少页码")
            self.assertEqual(paragraph.style.name, "Footer")
        self.assertEqual(self.doc.styles["Footer"].paragraph_format.alignment,
                         WD_ALIGN_PARAGRAPH.CENTER)

    def test_header_footer_styles_have_no_tabs(self) -> None:
        """页眉页脚样式不带制表位（与页脚格式一致）。"""
        for name in ("Header", "Footer"):
            xml = self.doc.styles[name].element.xml
            self.assertNotIn("<w:tabs>", xml, f"{name} 样式不应有制表位")

    def test_header_and_footer_have_no_border(self) -> None:
        """本科页眉页脚都不带框线。"""
        with ZipFile(self.dotm_path) as zf:
            styles_xml = zf.read("word/styles.xml").decode("utf-8")
        for name in ("header", "footer"):
            block = style_block(styles_xml, name)
            self.assertTrue(block, f"dotm 缺少 {name} 样式")
            self.assertNotRegex(
                block,
                r'<w:(bottom|top|left|right) w:val="(single|double|thick|dashed)"',
            )

    def test_graduate_header_keeps_bottom_border(self) -> None:
        """对照项：研究生页眉仍有一条下框线 0.75 磅。"""
        graduate_dotm = GRADUATE["dotm"]
        if not graduate_dotm.exists():
            self.skipTest(f"缺少研究生 dotm：{graduate_dotm}")
        block = style_block(read_part(graduate_dotm, "word/styles.xml"), "header")
        self.assertIn('<w:bottom w:val="single" w:sz="6"', block)

    def test_ribbon_groups(self) -> None:
        """两版均有封面组，本科仅提供密级开关。"""
        undergraduate_xml = UNDERGRADUATE["ribbon"].read_text(encoding="utf-8")
        graduate_xml = GRADUATE["ribbon"].read_text(encoding="utf-8")
        self.assertEqual(undergraduate_xml.count("<group id="), 11)
        self.assertEqual(graduate_xml.count("<group id="), 11)
        self.assertIn('id="grpCover"', undergraduate_xml)
        self.assertIn('onAction="ToggleSecurity"', undergraduate_xml)
        self.assertNotIn('onAction="RefreshCover"', undergraduate_xml)
        self.assertIn('id="grpCover"', graduate_xml)


    def test_cover_logo_follows_vi_manual(self) -> None:
        """封面校徽与校名按《清华大学视觉形象识别系统手册》A-3-03 组合。"""
        # 画布 171×49 bp：校徽 49 bp、间距 10 bp、标准字 112×41 bp，
        # 即手册的 15X / 3X / 34.3X×12.6X；插入尺寸 6.03×1.73 cm。
        doc_xml = read_part(self.docx_path, "word/document.xml")
        extents = re.findall(r'<wp:extent cx="(\d+)" cy="(\d+)"', doc_xml)
        self.assertTrue(extents, "封面缺少组合标志")
        cx, cy = (int(value) for value in extents[0])
        self.assertAlmostEqual(cx / 360000, 6.03, places=2)
        self.assertAlmostEqual(cy / 360000, 1.73, places=2)

        svg = (ROOT / "assets" / "thu-cover-logo.svg").read_text(encoding="utf-8")
        self.assertIn('width="171" height="49"', svg)
        self.assertIn('viewBox="0 0 171 49"', svg)

        png = (ROOT / "assets" / "thu-cover-logo.png").read_bytes()
        self.assertEqual(png[:8], b"\x89PNG\r\n\x1a\n")
        size = (int.from_bytes(png[16:20], "big"), int.from_bytes(png[20:24], "big"))
        self.assertEqual(size, (684, 196), "PNG 兜底应为画布的 4 倍")


if __name__ == "__main__":
    unittest.main(verbosity=2)
