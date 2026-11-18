"""Word COM 测试：宏功能、导出产物、本科变体页眉页脚。

宏单元、导出与本科 Word 检查共用本文件：每个用例打开 dotm 临时副本，
不修改 output/ 下的交付文件。静态格式/样式契约见 tests/test_build.py。
"""

from __future__ import annotations

import hashlib
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from tests.support import BASE_COUNTS, WordDotmTest, product_paths



# 一条预置的合并引用域 QUOTE "THU_CITE:0001"：导出前由 UpdateAllFieldsNow
# 更新域并 NormalizeCitations 改写，再由 FreezeCitationFields 冻结成静态文字。
_CITE_PARAGRAPH = (
    '<w:p><w:r><w:t xml:space="preserve">注入引用：</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> QUOTE "THU_CITE:0001" </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r><w:t>THU_CITE:0001</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
)


def write_dotm_with_citation(src: Path, dst: Path) -> None:
    """在副本末尾加一条指向模板已有条目的未锁定引用。"""
    with ZipFile(src) as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with ZipFile(dst, "w") as zout:
        for info, data in items:
            if info.filename == "word/document.xml":
                root = etree.fromstring(data)
                body = root.find(qn("w:body"))
                cite = _CITE_PARAGRAPH.replace("<w:p>", f'<w:p xmlns:w="{root.nsmap["w"]}">', 1)
                body.find(qn("w:sectPr")).addprevious(etree.fromstring(cite.encode("utf-8")))
                data = etree.tostring(root, encoding="utf-8")
            zout.writestr(info, data)


class _CitationExportChecks:
    """导出冻结引用：在 dotm 副本里预置真实域，再走 ExportDocxFile。"""

    def setUp(self) -> None:
        # 不用 doc.Save()：同一 Word 实例在导出过别的文档后保存 dotm 副本会卡在
        # 隐藏的模态保存（见 .agents/docs/70-pitfalls.md 第 10 条）。
        """预置冻结引用域后打开 dotm 副本。"""
        self.doc_path = self.tmpdir / f"{self._testMethodName}.dotm"
        write_dotm_with_citation(product_paths(self.variant)["dotm"], self.doc_path)
        self.doc = self.app.Documents.Open(
            str(self.doc_path), ReadOnly=False, AddToRecentFiles=False
        )
        # Word 打开模板会刷新域并标成未保存。磁盘副本已预置本用例的全部改动。
        self.doc.Saved = True

    def test_export_docx_freezes_citation(self) -> None:
        """提交版 docx 无宏：引用域转成静态文字，打开后按 F9 不会回退成域代码。"""
        exported_path = self.run_macro("ExportDocxFile", True)
        self.assertTrue(exported_path, self.run_macro("GetLastExportError"))
        docx = Path(exported_path)
        with ZipFile(docx) as zf:
            xml = zf.read("word/document.xml").decode("utf-8")
        self.assertNotIn("THU_CITE:", xml)
        exported = Document(docx)
        self.assertTrue(any("注入引用：[1]" in p.text for p in exported.paragraphs))
        reopened = self.app.Documents.Open(str(docx), ReadOnly=True, AddToRecentFiles=False)
        try:
            reopened.Fields.Update()
            self.assertIn("注入引用：[1]", reopened.Content.Text)
            self.assertNotIn("THU_CITE:", reopened.Content.Text)
        finally:
            reopened.Close(0)

    def test_deleted_reference_is_reported_and_its_id_is_not_reused(self) -> None:
        """旧稿中的引用标识在删除或换条目后仍不能重新分配。"""
        target = "THU_REF_ENTRY_0001"
        self.run_macro("UpdateAllFieldsNow")
        self.assertIn("注入引用：[1]", self.doc.Content.Text)
        self.doc.Bookmarks.Item(target).Range.Paragraphs(1).Range.Delete()
        self.run_macro("UpdateAllFieldsNow")
        self.assertIn("注入引用：[引用失效]", self.doc.Content.Text)
        self.assertIn("引用=4|", self.run_macro("TestTemplateCheck"))

        self.doc.Content.InsertAfter("\r示例新增文献")
        self.doc.Paragraphs.Last.Range.Style = "参考文献条目"
        self.run_macro("UpdateAllFieldsNow")
        self.assertFalse(self.doc.Bookmarks.Exists(target))
        self.assertIn("注入引用：[引用失效]", self.doc.Content.Text)

    def test_direct_save_as_preserves_refreshed_citation_after_field_update(self) -> None:
        """先用模板更新，直接另存为无宏稿后的 F9 也不会改写文献引用。"""
        self.run_macro("UpdateAllFieldsNow")
        path = self.tmpdir / "manual.docx"
        self.app.DisplayAlerts = 0
        self.doc.SaveAs2(FileName=str(path), FileFormat=12, AddToRecentFiles=False)
        self.doc.Close(0)
        self.doc = self.app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
        self.doc.Fields.Update()
        self.assertIn("注入引用：[1]", self.doc.Content.Text)
        self.assertNotIn("THU_CITE:", self.doc.Content.Text)


class _ExportChecks:
    """两变体的导出产物与写作稿完整性。"""

    def test_export_preserves_hyphenation(self) -> None:
        # 夹具未修改磁盘副本；Word打开模板后更新域可能将其标记为未保存。
        self.doc.Saved = True
        exported = self.run_macro("ExportDocxFile", True)
        self.assertTrue(exported, self.run_macro("GetLastExportError"))
        reopened = self.app.Documents.Open(str(exported), ReadOnly=True, AddToRecentFiles=False)
        try:
            self.assertTrue(reopened.AutoHyphenation)
            self.assertFalse(reopened.HyphenateCaps)
            self.assertEqual(reopened.ConsecutiveHyphensLimit, 2)
            self.assertFalse(reopened.Styles("参考文献条目").ParagraphFormat.Hyphenation)
            self.assertEqual(reopened.Styles("参考文献条目").ParagraphFormat.WordWrap, -1)
        finally:
            reopened.Close(0)

    def test_security_survives_export(self) -> None:
        """密级随 DOCX/PDF 导出保留，且仅出现在封面。"""
        from pypdf import PdfReader

        self.doc.Close(0)
        self.doc = self.app.Documents.Add(Template=str(self.doc_path), Visible=False)
        self.doc.SaveAs2(FileName=str(self.tmpdir / "security.docx"), FileFormat=12, AddToRecentFiles=False)
        result = self.run_macro("TestCoverSecurityToggle")
        self.assertTrue(result.startswith("ok|秘密★5年|False|True|秘密★5年|"), result)
        self.doc.Save()
        exported = self.run_macro("ExportDocxFile", True)
        self.assertTrue(exported, self.run_macro("GetLastExportError"))
        document = Document(exported)
        texts = lambda h: "".join(h._element.itertext())
        if self.variant == "undergraduate":
            # Word 在空单元格重建控件后可能将 sdt 包在 tc 外；按 tag 查找。
            controls = document.tables[0]._tbl.xpath('.//w:sdt[w:sdtPr/w:tag[@w:val="密级"]]')
            self.assertEqual(len(controls), 1)
            self.assertIn("秘密★5年", "".join(t.text or "" for t in controls[0].iter(qn("w:t"))))
            self.assertNotIn("秘密★5年", texts(document.sections[0].header))
        else:
            self.assertIn("秘密★5年", texts(document.sections[0].header))
        for section in list(document.sections)[1:]:
            self.assertNotIn("秘密★5年", texts(section.header))
        pdf = self.run_macro("ExportPdfFile", True)
        self.assertTrue(pdf, self.run_macro("GetLastExportError"))
        pages = PdfReader(pdf).pages
        security_line = r"(?m)^[ \t]*秘[ \t]*密[ \t]*★[ \t]*5[ \t]*年[ \t]*$"
        self.assertRegex(pages[0].extract_text(), security_line)
        # 正文使用说明包含带引号的示例文字，不能按全文子串判断页眉泄漏。
        for page in pages[1:]:
            self.assertNotRegex(page.extract_text(), security_line)

    def test_export_docx_and_pdf(self) -> None:
        """从模板新建写作稿，导出后提交版解除关联，写作稿保留模板链接。"""
        self.doc.Close(0)
        self.doc = self.app.Documents.Add(Template=str(self.doc_path), Visible=False)
        manuscript = self.tmpdir / "manuscript.docx"
        self.doc.UpdateStylesOnOpen = True
        self.doc.SaveAs2(FileName=str(manuscript), FileFormat=12, AddToRecentFiles=False)
        before = hashlib.sha256(manuscript.read_bytes()).hexdigest()
        pdf_path = self.run_macro("ExportPdfFile", True)
        self.assertTrue(pdf_path, self.run_macro("GetLastExportError"))
        docx_path = self.run_macro("ExportDocxFile", True)
        self.assertTrue(docx_path, self.run_macro("GetLastExportError"))
        pdf, docx = Path(pdf_path), Path(docx_path)
        after = hashlib.sha256(manuscript.read_bytes()).hexdigest()

        self.assertEqual(before, after, "导出不应修改写作稿")
        self.assertEqual(self.doc.AttachedTemplate.Name, self.doc_path.name)
        self.assertTrue(self.doc.UpdateStylesOnOpen)
        self.assertTrue(pdf.is_file(), pdf)
        self.assertTrue(docx.is_file(), docx)
        self.assertEqual(pdf.read_bytes()[:4], b"%PDF")
        with ZipFile(docx) as zf:
            names = zf.namelist()
            self.assertFalse(any("vbaProject" in n for n in names), names)
            self.assertFalse(any(n.startswith("customUI/") for n in names), names)
            self.assertNotIn("word/_rels/settings.xml.rels", names)
            settings = etree.fromstring(zf.read("word/settings.xml"))
            self.assertIsNone(settings.find(qn("w:attachedTemplate")))
            update_styles = settings.find(qn("w:linkStyles"))
            if update_styles is not None:
                self.assertIn(update_styles.get(qn("w:val")), ("0", "false", "off"))

        exported = Document(docx)
        self.assertEqual(len(exported.sections), BASE_COUNTS[self.variant]["sections"])
        self.assertTrue(any(s.name == "文献引用" for s in exported.styles))

    def test_export_removes_embedded_thesis_ribbon(self) -> None:
        """写作稿同时带内嵌功能区和模板关联时，提交版移除二者。"""
        manuscript = self.tmpdir / "duplicated-ribbon.docx"
        self.doc.SaveAs2(FileName=str(manuscript), FileFormat=12, AddToRecentFiles=False)
        self.doc.AttachedTemplate = str(self.doc_path)
        self.doc.Save()
        with ZipFile(manuscript) as zf:
            self.assertIn("customUI/customUI14.xml", zf.namelist())
            self.assertIn("word/_rels/settings.xml.rels", zf.namelist())
        before = manuscript.read_bytes()

        result = self.run_macro("ExportDocxFile", True)
        self.assertTrue(result, self.run_macro("GetLastExportError"))
        self.assertEqual(manuscript.read_bytes(), before)
        with ZipFile(result) as zf:
            names = zf.namelist()
            self.assertFalse(any(n.startswith("customUI/") for n in names), names)
            self.assertNotIn(b"/ui/extensibility", zf.read("_rels/.rels"))
            self.assertNotIn("word/_rels/settings.xml.rels", names)
            self.assertNotIn("word/vbaProject.bin", names)

    def test_export_overwrites_existing_target(self) -> None:
        """检查重复导出覆盖同一目标文件。"""
        self.doc.Saved = True
        first_path = self.run_macro("ExportDocxFile", True)
        self.assertTrue(first_path, self.run_macro("GetLastExportError"))
        first = Path(first_path)
        first_mtime = first.stat().st_mtime_ns
        second_path = self.run_macro("ExportDocxFile", True)
        self.assertTrue(second_path, self.run_macro("GetLastExportError"))
        second = Path(second_path)
        self.assertEqual(first, second)
        self.assertGreaterEqual(second.stat().st_mtime_ns, first_mtime)


class GraduateCitationExportTest(_CitationExportChecks, WordDotmTest):
    """研究生变体的导出冻结引用检查。"""
    variant = "graduate"


class UndergraduateCitationExportTest(_CitationExportChecks, WordDotmTest):
    """本科变体的导出冻结引用检查。"""
    variant = "undergraduate"


class GraduateExportTest(_ExportChecks, WordDotmTest):
    """研究生变体的导出产物与母版完整性检查。"""
    variant = "graduate"


class UndergraduateExportTest(_ExportChecks, WordDotmTest):
    """本科变体的导出产物与母版完整性检查。"""
    variant = "undergraduate"


class _StylePresentationChecks:
    def test_hyphenation_defaults_and_style_exceptions_survive_reopening(self) -> None:
        self.doc.Close(0)
        self.doc = self.app.Documents.Add(Template=str(self.doc_path), Visible=False)
        path = self.tmpdir / "hyphenation.docx"
        self.doc.SaveAs2(FileName=str(path), FileFormat=12, AddToRecentFiles=False)
        self.doc.Close(0)
        self.doc = self.app.Documents.Open(str(path), AddToRecentFiles=False)
        self.assertTrue(self.doc.AutoHyphenation)
        self.assertTrue(self.app.CommandBars.GetPressedMso("HyphenationAutomatic"))
        self.assertFalse(self.app.CommandBars.GetPressedMso("HyphenationNone"))
        self.assertFalse(self.doc.HyphenateCaps)
        self.assertEqual(self.doc.ConsecutiveHyphensLimit, 2)
        # Word 的实际排版中，COM 的 -1 对应按词换行，0 会在字符间拆词。
        for first in self.doc.StoryRanges:
            story = first
            while story is not None:
                self.assertEqual(story.ParagraphFormat.WordWrap, -1, story.StoryType)
                story = story.NextStoryRange
        # 检查 Word 实际提供的全部样式，包括未使用和隐式内置样式。
        for style in self.doc.Styles:
            if style.Type in (2, 4):  # 字符/编号定义不提供 ParagraphFormat；包内定义由静态测试覆盖。
                continue
            self.assertEqual(style.ParagraphFormat.WordWrap, -1, style.NameLocal)
        exceptions = {"行间代码", "公式", "参考文献条目", "参考文献条目（著者出版年）"}
        for name in ("正文", "章标题（有编号）", "一级节标题", "表格文字", "脚注文本", *exceptions):
            with self.subTest(style=name):
                paragraph = self.doc.Paragraphs.Add()
                paragraph.Range.Text = "Automatically formatted English text."
                paragraph.Style = self.doc.Styles(name)
                self.assertEqual(bool(paragraph.Format.Hyphenation), name not in exceptions)
                self.assertEqual(paragraph.Format.WordWrap, -1, name)
        # 正文即使含URL，也仅沿用正文样式，不因内容改变断字设置。
        paragraph = self.doc.Paragraphs.Add()
        paragraph.Range.Text = "https://example.org/resource"
        paragraph.Style = self.doc.Styles("正文")
        self.assertTrue(paragraph.Format.Hyphenation)
        self.assertEqual(paragraph.Format.WordWrap, -1)

    def test_western_words_move_to_next_line_without_midword_wrapping(self) -> None:
        workspace = self.doc.Sections.Last.PageSetup
        text = "例文 approximation communication"
        paragraph = self.doc.Paragraphs.Add()
        paragraph.Range.Text = text
        paragraph.Style = self.doc.Styles("正文")
        paragraph.Range.Font.Name = "Calibri"
        paragraph.Range.Font.Size = 10
        paragraph.Range.LanguageID = 1033
        paragraph.Format.FirstLineIndent = 0
        paragraph.Format.LeftIndent = 0
        paragraph.Format.RightIndent = (
            workspace.PageWidth - workspace.LeftMargin - workspace.RightMargin - workspace.Gutter - 92
        )
        # 暂停此测试段的自动断字，以区分正常断字和无连字符的拆词。
        paragraph.Format.Hyphenation = False
        self.doc.Repaginate()
        start = paragraph.Range.Start
        line_numbers = [self.doc.Range(start + i, start + i).Information(10) for i in range(len(text))]
        first = text.index("approximation")
        last = first + len("approximation") - 1
        self.assertGreater(line_numbers[first], line_numbers[0], line_numbers)
        self.assertEqual(line_numbers[first], line_numbers[last], line_numbers)

    def test_native_heading_styles_supply_separator_and_appendix_references(self) -> None:
        styles = ("章标题（有编号）", "一级节标题", "二级节标题", "三级节标题",
                  "附录标题", "附录一级节标题", "附录二级节标题", "附录三级节标题")
        for style in styles:
            paragraph = self.doc.Paragraphs.Add()
            paragraph.Range.Text = "手动新增标题"
            paragraph.Style = self.doc.Styles(style)
            level = paragraph.Range.ListFormat.ListLevelNumber
            definition = paragraph.Range.ListFormat.ListTemplate.ListLevels(level)
            self.assertTrue(definition.NumberFormat.endswith("　"), (style, definition.NumberFormat))
            self.assertEqual(paragraph.Range.Text.strip(), "手动新增标题")
            self.assertFalse(paragraph.Range.Text.startswith("　"))
        for style in styles[4:]:
            self.assertTrue(self.run_macro("IsChapterReferenceStyle", style))

    def test_long_url_preserves_text_without_midword_wrapping(self) -> None:
        url = "https://example.org/" + "a" * 240 + "/resource"
        for style in ("参考文献条目", "参考文献条目（著者出版年）"):
            paragraph = self.doc.Paragraphs.Add()
            paragraph.Range.Text = url
            paragraph.Style = self.doc.Styles(style)
            self.doc.Repaginate()
            self.assertEqual(paragraph.Format.WordWrap, -1, style)
            self.assertFalse(paragraph.Format.Hyphenation, style)
            self.assertEqual(paragraph.Range.Text.rstrip("\r"), url)

    def test_check_table_header_stays_with_first_data_row(self) -> None:
        self.doc.Repaginate()
        for table in self.doc.Tables:
            if not table.Range.Text.startswith("检查项\r"):
                continue
            first = table.Rows(1).Range.Duplicate
            first.Collapse(1)
            data = table.Rows(2).Range.Duplicate
            data.Collapse(1)
            self.assertEqual(first.Information(3), data.Information(3))

    def test_style_presentation_after_reopening_new_document(self) -> None:
        self.doc.Close(0)
        self.doc = self.app.Documents.Add(Template=str(self.doc_path), Visible=False)
        draft = self.tmpdir / "draft.docx"
        self.doc.SaveAs2(FileName=str(draft), FileFormat=12, AddToRecentFiles=False)
        self.doc.Close(0)
        self.doc = self.app.Documents.Open(str(draft), AddToRecentFiles=False)
        self.doc.FormattingShowFilter = 5  # wdShowFilterFormattingRecommended
        self.doc.StyleSortMethod = 1  # wdStyleSortRecommended

        styles = self.doc.Styles
        for alias in ("章标题（有编号）", "一级节标题", "二级节标题", "三级节标题",
                      "目录一级条目", "目录二级条目", "目录三级条目", "清单条目", "明显强调",
                      "列表编号1", "列表编号2", "列表编号3",
                      "列表项目符号1", "列表项目符号2", "列表项目符号3"):
            style = styles(alias)
            self.assertTrue(style.BuiltIn, alias)
            self.assertIn(alias, style.NameLocal.split(","))
        cover = styles("封面作者信息")
        table = styles("三线表")
        hidden_heading = styles(-6)  # wdStyleHeading5
        self.assertEqual(sum(bool(style.QuickStyle) for style in styles), 22)
        self.assertLess(styles(-1).Priority, styles(-2).Priority)
        self.assertLess(styles("图注").Priority, cover.Priority)
        self.assertLess(cover.Priority, hidden_heading.Priority)
        # Word 的 COM Visibility 在写回文件后对应 w:semiHidden 的布尔值。
        self.assertFalse(cover.Visibility)
        self.assertFalse(cover.QuickStyle)
        self.assertFalse(table.Visibility)
        self.assertEqual(table.Type, 3)  # wdStyleTypeTable
        for name in ("列表编号", "列表项目符号"):
            for suffix in ("", " 2", " 3"):
                style = styles(name + suffix)
                self.assertTrue(style.QuickStyle, style.NameLocal)
                self.assertEqual(style.Priority, 25)

        self.doc.Paragraphs.Last.Style = hidden_heading
        self.assertTrue(hidden_heading.Visibility)
        self.assertFalse(hidden_heading.UnhideWhenUsed)
        self.assertFalse(hidden_heading.QuickStyle)


class GraduateStylePresentationTest(_StylePresentationChecks, WordDotmTest):
    variant = "graduate"


class UndergraduateStylePresentationTest(_StylePresentationChecks, WordDotmTest):
    variant = "undergraduate"


class _AppendixSequenceChecks:
    def test_second_appendix_restarts_formula_figure_and_table_numbers(self) -> None:
        from pythoncom import VT_DISPATCH
        from win32com.client import VARIANT

        doc = self.doc
        end = doc.Content.End - 1
        labels = (("图一", "附录图"), ("表一", "附录表"), ("图二", "附录图"), ("表二", "附录表"))
        doc.Range(end, end).InsertAfter("\r第二份附录\r\r" + "\r".join(text for text, _ in labels))
        heading = next(p for p in doc.Paragraphs if p.Range.Text.strip() == "第二份附录")
        heading.Range.Select()
        self.assertTrue(self.run_macro("ApplyHeadingStyleByName", "附录标题"))
        self.assertIn("B", heading.Range.ListFormat.ListString)

        heading.Next().Range.Select()
        self.run_macro("RunNormalizeFormula", VARIANT(VT_DISPATCH, None), True)
        for text, label in labels:
            paragraph = next(p for p in doc.Paragraphs if p.Range.Text.strip() == text)
            end = paragraph.Range.End - 1
            doc.Fields.Add(doc.Range(end, end), -1, f"SEQ {label} \\* ARABIC \\s 1", False)
        self.run_macro("UpdateAllFieldsNow")

        formula = next(p for p in doc.Paragraphs
                       if p.Range.Start >= heading.Range.End and p.Range.OMaths.Count)
        self.assertEqual([f.Result.Text.strip() for f in formula.Range.Fields], ["B", "1"])
        for label in ("附录图", "附录表"):
            results = [f.Result.Text.strip() for f in doc.Fields
                       if f.Code.Start > heading.Range.End and f.Code.Text.strip().startswith(f"SEQ {label} ")]
            self.assertEqual(results, ["1", "2"], label)


class _CoverSecurityChecks:
    """密级切换与标题行数变化不移动后续信息。"""

    def test_security_header_does_not_move_cover(self) -> None:
        header = self.doc.Sections(1).Headers(1)
        self.assertEqual(header.Range.Text.strip(), "")
        if self.variant == "undergraduate":
            security = self.doc.Tables(1).Range.Cells(1).Range
            self.assertEqual(header.Range.ContentControls.Count, 0)
        else:
            security = header.Range
        self.assertEqual(security.ParagraphFormat.Alignment, 0)
        self.assertEqual(security.ContentControls.Count, 1)
        body = self.doc.Content.Text
        self.doc.Repaginate()
        pages = self.doc.ComputeStatistics(2)

        def positions():
            return [cell.Range.Information(6) for cell in self.doc.Tables(1).Range.Cells]

        before = positions()
        if self.variant == "undergraduate":
            # logo 顶部留白参考 ThuThesis，作者与日期位置参考官方本科示例。
            logo_top = self.doc.InlineShapes(1).Range.Information(6)
            self.assertAlmostEqual(logo_top - self.doc.Sections(1).PageSetup.TopMargin, 39.2, delta=0.75)
            for tag, expected in (("系别", 473.2), ("成文日期", 680.5)):
                cc = self.doc.SelectContentControlsByTag(tag)(1)
                self.assertAlmostEqual(cc.Range.Information(6), expected, delta=1)
            self.assertEqual(self.doc.Sections(1).Range.Characters.Last.Information(3), 1)
        for expected in ("秘密★5年", "", "秘密★5年"):
            self.assertEqual(self.run_macro("TestCoverSecurityClick"), expected)
            self.doc.Repaginate()
            self.assertEqual(security.Text.strip("\r\x07"), expected)
            self.assertEqual(security.ContentControls.Count, 1 if expected else 0)
            content = self.doc.Content.Text
            if self.variant == "undergraduate" and expected:
                content = content.replace(expected, "", 1)
            # Word 首次编辑表格控件时会补齐空单元格的段落结束标记。
            normalize = lambda text: text.replace("\r", "").replace("\x07", "").replace("，不包含涉及国家秘密的内容", "")
            self.assertEqual(normalize(content), normalize(body))
            if self.variant == "graduate":
                self.assertEqual("不包含涉及国家秘密的内容" in content, not bool(expected))
            self.assertEqual(self.doc.ComputeStatistics(2), pages)
            self.assertEqual(positions(), before)
            for section in list(self.doc.Sections)[1:]:
                self.assertNotIn("秘密★5年", section.Headers(1).Range.Text)

        title = self.doc.SelectContentControlsByTag("论文题目")(1)
        for text in ("模板标题", "模板标题\v第二行标题"):
            title.Range.Text = text
            self.doc.Repaginate()
            self.assertEqual(positions(), before)
            if self.variant == "graduate":
                first = title.Range.Duplicate
                first.End = first.Start + 1
                self.assertAlmostEqual(first.Information(6), self.doc.Sections(1).PageSetup.TopMargin, delta=0.3)


class GraduateMacroTest(_CoverSecurityChecks, _AppendixSequenceChecks, WordDotmTest):
    """研究生模板宏单元：公式、引用、封面级联、交互、题注与样式。"""

    def test_degree_switch_then_repeated_refresh_keeps_author_information(self) -> None:
        from pythoncom import VT_DISPATCH
        from win32com.client import VARIANT
        controls = self.doc.SelectContentControlsByTag
        null = VARIANT(VT_DISPATCH, None)
        self.run_macro("ToggleAssociateSupervisor", null)
        self.run_macro("ToggleCoSupervisor", null)
        tags = ("培养单位", "研究生（专业学位改为申请人）", "指导教师",
                "副指导教师（可选）", "联合指导教师（可选）")
        values = {tag: "示例信息" + str(i) for i, tag in enumerate(tags)}
        for tag, value in values.items():
            controls(tag)(1).Range.Text = value
        for degree in ("工学博士学位", "金融硕士专业学位", "资源与环境博士专业学位", "工学硕士学位"):
            with self.subTest(degree=degree):
                cc = controls("学位类别")(1)
                cc.Range.Select()
                next(entry for entry in cc.DropdownListEntries if entry.Text == degree).Select()
                # 不先离开下拉框；直接调用功能区实际入口，再连续刷新。
                for _ in range(3):
                    self.run_macro("RefreshCover", null)
                    cell = controls("培养单位")(1).Range.Cells(1)
                    for tag, value in values.items():
                        self.assertEqual(controls(tag).Count, 1, tag)
                        self.assertEqual(controls(tag)(1).Range.Text, value)
                        self.assertTrue(controls(tag)(1).Range.InRange(cell.Range))
                    self.assertEqual(cell.Range.Paragraphs.Count, 6)

    def test_author_rebuild_failure_restores_original_information(self) -> None:
        import sys
        sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
        try:
            from build_dotm import Vbom
        finally:
            sys.path.pop(0)
        controls = self.doc.SelectContentControlsByTag
        before = controls("培养单位")(1).Range.Cells(1).Range.Text
        # 只向当前临时副本注入故障：第一行已重建、第二行开始前失败。
        with Vbom():
            module = self.doc.VBProject.VBComponents("modCover").CodeModule
            source = module.Lines(1, module.CountOfLines)
            marker = 'FillSupervisorRow cell.Range.Paragraphs(row), LABEL_AUTHOR, TAG_AUTHOR_CN, values(2)'
            self.assertIn(marker, source)
            source = source.replace(marker, 'Err.Raise vbObjectError + 999, "test", "模拟控件创建失败"', 1)
            module.DeleteLines(1, module.CountOfLines)
            module.AddFromString(source)
        controls("学位类别")(1).Range.Select()
        self.run_macro("TestCoverSelectDegree", "金融硕士专业学位")
        self.assertEqual(controls("培养单位")(1).Range.Cells(1).Range.Text, before)
        for tag in ("培养单位", "学科（专业学位改为专业领域）", "研究生（专业学位改为申请人）", "指导教师"):
            self.assertEqual(controls(tag).Count, 1, tag)

    def test_english_discipline_prefix_is_separate_and_survives_switching(self) -> None:
        controls = self.doc.SelectContentControlsByTag
        cc = controls("英文学科")(1)
        self.assertEqual(cc.Range.Text, "Computer Science and Technology")
        self.assertEqual(cc.Range.Paragraphs.Count, 1)
        prefix = cc.Range.Paragraphs(1).Previous()
        self.assertEqual(prefix.Range.Text.strip("\r\x07"), "in")
        self.assertEqual(prefix.Range.ContentControls.Count, 0)
        self.assertTrue(self.run_macro("TestCoverSelectDegree", "金融硕士专业学位"))
        self.assertEqual(prefix.Range.Text.strip("\r\x07"), "")
        self.assertTrue(self.run_macro("TestCoverSelectDegree", "工学硕士学位"))
        cc = controls("英文学科")(1)
        self.assertEqual(cc.Range.Text, "Computer Science and Technology")
        self.assertNotIn("\x0b", cc.Range.Text)
        self.assertEqual(cc.Range.Paragraphs(1).Previous().Range.Text.strip("\r\x07"), "in")

    def test_degree_change_preserves_author_controls_while_dropdown_selected(self) -> None:
        controls = self.doc.SelectContentControlsByTag
        tags = ("培养单位", "学科（专业学位改为专业领域）", "研究生（专业学位改为申请人）", "指导教师")
        before = {tag: (controls(tag)(1).ID, controls(tag)(1).Range.Text) for tag in tags}
        for degree in ("工学博士学位", "工学硕士学位", "文学硕士学位"):
            with self.subTest(degree=degree):
                controls("学位类别")(1).Range.Select()
                self.assertTrue(self.run_macro("TestCoverSelectDegree", degree))
                for tag in tags:
                    self.assertEqual(controls(tag).Count, 1, tag)
                    self.assertEqual((controls(tag)(1).ID, controls(tag)(1).Range.Text), before[tag])

    def test_degree_event_updates_english_without_leaving_dropdown(self) -> None:
        controls = self.doc.SelectContentControlsByTag
        for degree, english in (("工学博士学位", "Doctor of Philosophy"),
                                ("文学硕士学位", "Master of Arts"),
                                ("金融硕士专业学位", "Master of Finance"),
                                ("工学硕士学位", "Master of Science")):
            with self.subTest(degree=degree):
                cc = controls("学位类别")(1)
                cc.Range.Select()
                next(entry for entry in cc.DropdownListEntries if entry.Text == degree).Select()
                # 确认选择后选区收拢，仍在学位控件内；不调用级联测试入口。
                caret = cc.Range.Duplicate
                caret.Collapse(1)
                caret.Move(1, 1)
                caret.Select()
                self.assertEqual(controls("英文学位名称")(1).Range.Text, english)
                self.assertTrue(self.app.Selection.InRange(cc.Range))
                self.assertEqual(controls("研究生（专业学位改为申请人）")(1).Range.Text, "薛瑞尼")

    def test_optional_supervisors_are_anchored_when_cover_control_selected(self) -> None:
        from pythoncom import VT_DISPATCH
        from win32com.client import VARIANT
        controls = self.doc.SelectContentControlsByTag
        null = VARIANT(VT_DISPATCH, None)
        tags = ("培养单位", "学科（专业学位改为专业领域）", "研究生（专业学位改为申请人）", "指导教师")
        before = {tag: controls(tag)(1).Range.Text for tag in tags}
        for selection in ("学位类别", *tags, "body"):
            with self.subTest(selection=selection):
                if selection == "body":
                    self.doc.Range(self.doc.Content.End - 1, self.doc.Content.End - 1).Select()
                else:
                    controls(selection)(1).Range.Select()
                self.run_macro("ToggleCoSupervisor", null)
                self.run_macro("ToggleAssociateSupervisor", null)
                cell = controls("培养单位")(1).Range.Cells(1)
                assoc = controls("副指导教师（可选）")
                co = controls("联合指导教师（可选）")
                self.assertEqual(assoc.Count, 1)
                self.assertEqual(co.Count, 1)
                self.assertTrue(assoc(1).Range.InRange(cell.Range))
                self.assertTrue(co(1).Range.InRange(cell.Range))
                self.assertLess(controls("指导教师")(1).Range.End, assoc(1).Range.Start)
                self.assertLess(assoc(1).Range.End, co(1).Range.Start)
                self.assertEqual(cell.Range.Paragraphs.Count, 6)
                for tag in tags:
                    self.assertEqual(controls(tag)(1).Range.Text, before[tag])
                self.run_macro("ToggleAssociateSupervisor", null)
                self.run_macro("ToggleCoSupervisor", null)

    def test_cover_rebuild_checks_locked_controls_before_deleting_data(self) -> None:
        controls = self.doc.SelectContentControlsByTag
        author = controls("研究生（专业学位改为申请人）")(1)
        author.LockContentControl = True
        before = self.doc.SelectContentControlsByTag("培养单位")(1).Range.Cells(1).Range.Text
        controls("学位类别")(1).Range.Select()
        self.run_macro("TestCoverSelectDegree", "金融硕士专业学位")
        self.assertEqual(controls("研究生（专业学位改为申请人）")(1).Range.Text, "薛瑞尼")
        self.assertEqual(controls("培养单位")(1).Range.Cells(1).Range.Text, before)

    def test_cover_type_changes_keep_filled_optional_rows(self) -> None:
        from pythoncom import VT_DISPATCH
        from win32com.client import VARIANT
        controls = self.doc.SelectContentControlsByTag
        null = VARIANT(VT_DISPATCH, None)
        self.run_macro("ToggleAssociateSupervisor", null)
        self.run_macro("ToggleCoSupervisor", null)
        values = {"培养单位": "示例培养单位", "研究生（专业学位改为申请人）": "示例作者",
                  "指导教师": "示例导师", "副指导教师（可选）": "示例副导师",
                  "联合指导教师（可选）": "示例联合导师"}
        for tag, value in values.items():
            controls(tag)(1).Range.Text = value
        for degree in ("金融硕士专业学位", "资源与环境博士专业学位", "工学博士学位", "工学硕士学位"):
            with self.subTest(degree=degree):
                controls("学位类别")(1).Range.Select()
                self.assertTrue(self.run_macro("TestCoverSelectDegree", degree))
                for tag, value in values.items():
                    self.assertEqual(controls(tag).Count, 1, tag)
                    self.assertEqual(controls(tag)(1).Range.Text, value)
                    cell = controls("培养单位")(1).Range.Cells(1)
                    self.assertTrue(controls(tag)(1).Range.InRange(cell.Range))
                self.assertEqual(cell.Range.Paragraphs.Count, 6)

    def test_degree_events_in_new_document_and_reopened_writing_copy(self) -> None:
        template = str(self.doc_path)
        self.doc.Close(False)
        self.doc = self.app.Documents.Add(Template=template, Visible=False)
        for reopen, degree, english in ((False, "工学博士学位", "Doctor of Philosophy"),
                                        (True, "文学硕士学位", "Master of Arts")):
            with self.subTest(reopen=reopen):
                if reopen:
                    path = self.tmpdir / "cover-writing-copy.docx"
                    self.doc.SaveAs2(FileName=str(path), FileFormat=12, AddToRecentFiles=False)
                    self.doc.Close(False)
                    self.doc = self.app.Documents.Open(str(path), AddToRecentFiles=False)
                controls = self.doc.SelectContentControlsByTag
                cc = controls("学位类别")(1)
                cc.Range.Select()
                next(entry for entry in cc.DropdownListEntries if entry.Text == degree).Select()
                caret = cc.Range.Duplicate
                caret.Collapse(1)
                caret.Move(1, 1)
                caret.Select()
                self.assertEqual(controls("英文学位名称")(1).Range.Text, english)
                self.assertEqual(controls("研究生（专业学位改为申请人）")(1).Range.Text, "薛瑞尼")

    def test_cover_events_ignore_unrelated_documents(self) -> None:
        # 先启动模板事件接收器，再在不关联此模板的文档中使用同名控件。
        self.doc.SelectContentControlsByTag("学位类别")(1).Range.Select()
        unrelated = self.app.Documents.Add(Visible=False)
        try:
            unrelated.Activate()
            unrelated.Range(0, 0).Select()
            degree = unrelated.ContentControls.Add(4, unrelated.Range(0, 0))
            degree.Tag = "学位类别"
            degree.DropdownListEntries.Add("工学博士学位").Select()
            unrelated.Content.InsertAfter("\r保留的文字")
            before = unrelated.Content.Text
            degree.Range.Select()
            unrelated.Range(unrelated.Content.End - 1, unrelated.Content.End - 1).Select()
            self.assertEqual(unrelated.Content.Text, before)
            self.assertEqual(unrelated.ContentControls.Count, 1)
        finally:
            unrelated.Close(False)

    def test_cover_type_transitions_preserve_names_and_english_fields(self) -> None:
        controls = self.doc.SelectContentControlsByTag
        academic = controls("英文学科")(1).Range.Text
        author = controls("研究生（专业学位改为申请人）")(1).Range.Text
        self.assertTrue(self.run_macro("TestCoverSelectDegree", "金融硕士专业学位"))
        self.assertEqual(controls("学科（专业学位改为专业领域）").Count, 0)
        self.assertEqual(controls("英文学科")(1).Range.Text.strip(), "")
        self.assertTrue(self.run_macro("TestCoverSelectDegree", "资源与环境硕士专业学位"))
        self.assertEqual(controls("学科（专业学位改为专业领域）").Count, 1)
        controls("英文专业领域")(1).Range.Text = "(Environmental Engineering)"
        self.assertTrue(self.run_macro("TestCoverSelectDegree", "金融硕士专业学位"))
        self.assertEqual(controls("英文专业领域")(1).Range.Text.strip(), "")
        self.assertTrue(self.run_macro("TestCoverSelectDegree", "资源与环境硕士专业学位"))
        self.assertEqual(controls("英文专业领域")(1).Range.Text, "(Environmental Engineering)")
        self.assertTrue(self.run_macro("TestCoverSelectDegree", "工学硕士学位"))
        self.assertEqual(controls("英文学科")(1).Range.Text, academic)
        self.assertEqual(controls("研究生（专业学位改为申请人）")(1).Range.Text, author)

    def test_associate_supervisor_english_row(self) -> None:
        from pythoncom import VT_DISPATCH
        from win32com.client import VARIANT
        null = VARIANT(VT_DISPATCH, None)
        self.run_macro("ToggleAssociateSupervisor", null)
        cc = self.doc.SelectContentControlsByTag("指导教师（英文）")(1)
        self.assertIn("Associate Supervisor:", cc.Range.Text)
        cc.Range.Text = cc.Range.Text.replace("Professor ×××", "Professor Example")
        self.run_macro("ToggleCoSupervisor", null)
        self.assertIn("Professor Example", cc.Range.Text)
        self.run_macro("ToggleCoSupervisor", null)
        self.assertIn("Professor Example", cc.Range.Text)
        self.run_macro("ToggleAssociateSupervisor", null)
        self.assertNotIn("Associate Supervisor:", cc.Range.Text)

    def test_normalize_and_reference(self) -> None:
        """检查公式规范化与引用编号。"""
        first = self.run_macro("TestFormulaNormalize")
        self.assertIn("ok=True", first)
        second = self.run_macro("TestFormulaNormalize")
        self.assertIn("ok=True", second)
        self.assertEqual(second.count("#"), 1, second)

        ref = self.run_macro("TestFormulaReference")
        self.assertTrue(ref.startswith("ok|"), ref)
        self.assertRegex(ref, r"式\(\d+\.\d+\)$")

    def test_template_check(self) -> None:
        """检查宏应核对题注与引用，且不因长文档卡死。"""
        result = self.run_macro("TestTemplateCheck")
        self.assertEqual(result, "图注=0|表注=0|引用=0|提示=1")  # 表4.3由范围引用涵盖
        long_result = self.run_macro("TestCheckLong", 2000)
        elapsed = float(long_result.split("elapsed=")[1].split("|")[0])
        self.assertLess(elapsed, 3.0, long_result)

    def test_reference_check_report_groups_issues_without_editing(self) -> None:
        self.doc.Content.InsertAfter("\r(a)　手写子图注")
        self.doc.Paragraphs.Last.Range.Style = "子图注"
        self.doc.Content.InsertAfter("\r表9.9　手写表注")
        self.doc.Paragraphs.Last.Range.Style = "表注"
        end = self.doc.Content.End - 1
        self.doc.Fields.Add(self.doc.Range(end, end), -1, "REF THU_MISSING", False)
        self.doc.Saved = True
        before = self.doc.Content.Text

        report = self.run_macro("TestReferenceCheckReport")

        self.assertIn("图注（含子图注） · 1 项", report)
        self.assertIn("子图注缺少编号域", report)
        self.assertIn("表注编号缺少域", report)
        self.assertIn("引用 · 1 项", report)
        self.assertIn("未引用提示（不计为错误） · 1 项", report)
        self.assertIn("图表题注未被引用：THU_TAB_0004", report)
        self.assertEqual(self.doc.Content.Text, before)
        self.assertTrue(self.doc.Saved)

    def test_display_formula_modes(self) -> None:
        """无编号=显示 OMath（有 oMathPara、无编号域/书签）；有编号=显示 OMath 且带编号。"""
        result = self.run_macro("TestDisplayFormulaModes")
        self.assertEqual(result, "ok|True|0|0|True|True|True|0|0")

    def test_appendix_formula_numbering(self) -> None:
        """附录内公式按附录字母编号（式(C.1)），引用域回指同一编号。"""
        result = self.run_macro("TestAppendixFormulaNumbering")
        self.assertNotEqual(result, "NO_APPENDIX", result)
        self.assertTrue(result.startswith("ok|"), result)
        num, has_style, bm_text, ref_text = result.split("|")[1:5]

        # 章号是附录字母，不是正文章号。
        self.assertRegex(num, r"^[A-Z]\.\d+$")
        self.assertEqual(bm_text, num)
        self.assertEqual(ref_text, num)
        self.assertEqual(has_style, "True")

    def test_formula_edit_preserves_math_and_existing_ref_target(self) -> None:
        from pythoncom import VT_DISPATCH
        from win32com.client import VARIANT

        control = VARIANT(VT_DISPATCH, None)
        doc = self.doc
        old_name = "THU_EQ_0001"
        original = doc.Bookmarks(old_name).Range.Paragraphs(1)
        original.Range.InsertParagraphBefore()
        expression = "z=f(1)+2"
        start = original.Range.Start - 1
        doc.Range(start, start).InsertAfter(expression)
        formula = doc.OMaths.Add(doc.Range(start, start + len(expression))).OMaths(1)
        formula.BuildUp()
        formula.Range.Select()
        self.run_macro("RunNormalizeFormula", control, True)
        self.run_macro("UpdateAllFieldsNow")

        old_bookmark = doc.Bookmarks(old_name)
        self.assertEqual(old_bookmark.Range.Text.strip(), "4.2")
        self.assertIn("𝑚", old_bookmark.Range.Paragraphs(1).Range.Text)
        self.assertIn("(1)", doc.Content.Text)
        reference = next(f for f in doc.Fields if f.Code.Text.strip().startswith(f"REF {old_name} "))
        self.assertEqual(reference.Result.Text.strip(), "4.2")

        old_bookmark.Range.Paragraphs(1).Range.Select()
        self.run_macro("RunNormalizeFormula", control, True)
        self.run_macro("UpdateAllFieldsNow")
        self.assertEqual(doc.Bookmarks(old_name).Range.Text.strip(), "4.2")

    def test_numbered_heading_preserves_bookmark(self) -> None:
        start = self.doc.Content.End - 1
        self.doc.Range(start, start).InsertAfter("\r示例章标题")
        paragraph = self.doc.Paragraphs.Last
        target = self.doc.Range(paragraph.Range.Start, paragraph.Range.End - 1)
        self.doc.Bookmarks.Add("THU_CH_TEST", target)
        target.Select()
        self.assertTrue(self.run_macro("ApplyHeadingStyleByName", "章标题（有编号）"))
        self.assertTrue(self.doc.Bookmarks.Exists("THU_CH_TEST"))
        self.assertEqual(paragraph.Range.Text.strip(), "示例章标题")
        self.assertFalse(paragraph.Range.Text.startswith("　"))

    def test_citation_superscript(self) -> None:
        """检查文献引用标记为上标。"""
        result = self.run_macro("TestCitation")
        self.assertTrue(result.startswith("ok|"), result)
        self.assertIn("[1]", result)

    def test_citation_compression(self) -> None:
        """纯函数：连续区间、逗号加空格、单值、乱序去重、空串。"""
        result = self.run_macro("TestCitationCompress")
        self.assertEqual(result, "ok|2-3|7, 9, 11-13|1|1-3|[]")

    def test_citation_merge(self) -> None:
        """多选插入一个域：连续、两条连续、非连续、重复 id 都合并成上标序号。"""
        result = self.run_macro("TestCitationMerge")
        self.assertEqual(
            result,
            "ok|连续引用：[1-3]|两条连续：[2-3]|非连续引用：[1, 3]|重复引用：[2]|再整理一次：[1-3]",
        )

    def test_citation_update_round_trip(self) -> None:
        """普通 F9 与模板更新均保留正确显示的文献编号。"""
        result = self.run_macro("TestCitationUpdateRoundTrip")
        parts = result.split("|")
        self.assertEqual(parts[0], "ok", result)
        self.assertEqual(parts[1], "更新域测试：[1]", result)
        self.assertEqual(parts[2], "更新域测试：[1]", result)
        self.assertEqual(parts[3], "更新域测试：[1]", result)

    def test_citation_freeze(self) -> None:
        """导出前冻结：引用域数量归零，显示文字保持 [1]。"""
        result = self.run_macro("TestCitationFreeze")
        self.assertEqual(result, "ok|6|0|冻结测试：[1]", result)

    def test_degree_cascade_academic(self) -> None:
        """检查学术学位级联的封面显示值。"""
        result = self.run_macro("TestCoverDegreeCascade")
        # spec §11 的期望串末尾“|”之后还有第 8 段（学科候选数），故逐字断言前 7 段。
        self.assertTrue(
            result.startswith(
                "ok|工学硕士学位|Thesis submitted to|for the degree of|Master of Science|False|True|"
            ),
            result,
        )
        # 末项=学科候选数（工学 disciplines 数量），仅断言 >0
        parts = result.split("|")
        self.assertGreater(len(parts), 7, result)
        self.assertGreater(int(parts[7]), 0, result)
        disciplines = self.doc.SelectContentControlsByTag("学科（专业学位改为专业领域）")(1)
        names = [entry.Text for entry in disciplines.DropdownListEntries]
        self.assertIn("集成电路科学与工程", names)
        self.assertIn("管理科学与工程", names)
        text = self.doc.Content.Text
        self.assertIn("学科：", text)
        self.assertNotIn("专业领域：", text)

    def test_degree_cascade_professional(self) -> None:
        """检查专业学位级联的封面显示值。"""
        result = self.run_macro("TestCoverDegreeCascadeProfessional")
        self.assertEqual(
            result,
            "ok|金融硕士专业学位|Thesis submitted to|for the professional degree of|Master of Finance|True|True|0",
        )
        text = self.doc.Content.Text
        self.assertNotIn("专业领域：", text)
        self.assertIn("申请人：", text)

    def test_degree_cascade_doctor(self) -> None:
        """检查博士学位级联的封面显示值。"""
        result = self.run_macro("TestCoverDegreeCascadeDoctor")
        self.assertEqual(
            result,
            "ok|工学博士学位|Dissertation submitted to|for the degree of|Doctor of Philosophy|True",
        )

    def test_date_cascade(self) -> None:
        """检查封面中英文日期联动。"""
        result = self.run_macro("TestCoverDateCascade")
        self.assertEqual(result, "ok|二〇二六年十月|October, 2026|True")

    def test_refresh_cover(self) -> None:
        """检查刷新封面后学位与日期取值。"""
        result = self.run_macro("TestCoverRefresh")
        self.assertEqual(
            result,
            "ok|金融硕士专业学位|for the professional degree of|Master of Finance|二〇二七年三月|March, 2027",
        )

    def test_toggle_supervisor_rows(self) -> None:
        """检查导师行增删后的封面结构。"""
        result = self.run_macro("TestCoverToggleRows")
        self.assertEqual(result, "ok|6|4|1280|5|1600|6|1920|5|1920|4|1280")

    def test_english_cover_format(self) -> None:
        """检查英文封面字体与制表位格式。"""
        result = self.run_macro("TestEnglishCoverFormat")
        self.assertEqual(result, "ok|-1|-1|True|True")  # 学位名/日期加粗；导师 195.5 制表位；密级空

    def test_security_toggle(self) -> None:
        """检查密级控件的增删切换。"""
        result = self.run_macro("TestCoverSecurityToggle")
        # 存在性开关：空白→添加；已有内容→删除控件与文字；再点→重新添加。
        # Document.ContentControls 只计正文；密级开关不增删正文控件。
        self.assertEqual(result, "ok|秘密★5年|False|True|秘密★5年|16|16|16")

    def test_english_co_supervisor_sync(self) -> None:
        """检查英文副导师信息同步。"""
        result = self.run_macro("TestCoverEnglishCoSupervisorSync")
        self.assertEqual(result, "ok|False|True|True|True|True|True")

    def test_supervisor_toggle(self) -> None:
        """检查导师控件存在性切换。"""
        result = self.run_macro("TestCoverSupervisorToggle")
        # 控件存在时再点即删除（含已填内容）；不存在时插入空行。
        self.assertEqual(result, "ok|True|6|4|True|True||True|6|5")

    def test_formula_click_path(self) -> None:
        """检查功能区公式按钮调用链路。"""
        self.assertEqual(self.run_macro("TestFormulaClickPath"), "ok")

    def test_formula_three_states(self) -> None:
        """检查公式三种状态的切换结果。"""
        self.assertEqual(self.run_macro("TestFormulaThreeStates"), "ok|True|True|True")

    def test_formula_buttons(self) -> None:
        """行内公式按钮负责插入/调字体；空白行回退行间；行间按钮在文字后新建段。"""
        self.assertEqual(self.run_macro("TestFormulaButtons"), "ok|True|True|True|True")

    def test_table_caption_outside_table(self) -> None:
        """检查表格外插入表注的结果。"""
        result = self.run_macro("TestTableCaptionClick")
        self.assertTrue(result.startswith("ok|False|"), result)
        self.assertIn("表", result)

    def test_figure_caption_click(self) -> None:
        """检查图注按钮插入结果。"""
        result = self.run_macro("TestFigureCaptionClick")
        self.assertTrue(result.startswith("ok|图"), result)

    def test_caption_blank_line(self) -> None:
        """空白行无选中文字时，图注/表注/子图注也应能直接插入。"""
        result = self.run_macro("TestCaptionBlankLine")
        self.assertTrue(result.startswith("ok|"), result)
        self.assertIn("图注=图", result)
        self.assertIn("图题", result)
        self.assertIn("表注=表", result)
        self.assertIn("表题", result)
        self.assertRegex(result, r"子图注=\([a-z]\)")
        self.assertIn("子图题", result)

    def test_figure_caption_numbering(self) -> None:
        """图注编号格式：图<章>.<序> + 一个全角空格 + 题名，前后无多余空格。"""
        result = self.run_macro("TestFigureCaption")
        self.assertRegex(result, r"^ok\|图\d+\.\d+　测试图题$")

    def test_table_caption_numbering(self) -> None:
        """表注编号格式：表<章>.<序> + 一个全角空格 + 题名。"""
        result = self.run_macro("TestTableCaption")
        self.assertRegex(result, r"^ok\|表\d+\.\d+　测试表题$")

    def test_appendix_table_caption_numbering(self) -> None:
        """附录表注走附录 SEQ 分支，编号带附录字母（表A.1 而非 表2.1）。"""
        result = self.run_macro("TestAppendixTableCaption")
        self.assertRegex(result, r"^ok\|表[A-Z]\.\d+　测试附录表题$")

    def test_subfigure_caption(self) -> None:
        """检查子图注按序编号。"""
        result = self.run_macro("TestSubfigureCaption")
        self.assertEqual(
            result,
            "ok|(c)　甲|(d)　乙|(e)　丙|(f)　丁|(g)　戊|(h)　己|(a)　丙|(b)　丁",
        )

    def test_structure_macros(self) -> None:
        """检查结构宏生成的标题层级。"""
        self.assertEqual(self.run_macro("TestStructureMacros"), "ok|2")

    def test_picker_layout(self) -> None:
        """选择器窗体可加载并填充条目，布局随窗口尺寸收敛。"""
        form = self.run_macro("TestPickerForm")
        self.assertEqual(form, "ok|2|1|False|True|已选择 1 项|1")
        result = self.run_macro("TestPickerLayout")
        self.assertEqual(result, "ok|True|True|True|True")

    def test_reference_check_dialog_tabs(self) -> None:
        self.assertEqual(self.run_macro("TestReferenceDialog"), "4|表注 (1)|1. 检查示例|True")

    def test_table_reference_items(self) -> None:
        """检查表引用候选条目数量。"""
        result = self.run_macro("TestTableReferenceItems")
        self.assertTrue(result.startswith("ok|"), result)
        self.assertGreaterEqual(int(result.split("|")[1]), 1, result)

    def test_table_reference_inserts_ref(self) -> None:
        """检查表引用插入 REF 域。"""
        result = self.run_macro("TestTableRefText")
        self.assertTrue(result.startswith("ok|引用：表"), result)

    def test_chapter_reference_field_code(self) -> None:
        """章引用候选覆盖各级标题与附录，域代码用 REF 书签而非 STYLEREF。"""
        items = self.run_macro("TestChapterReferenceItems")
        self.assertTrue(items.startswith("ok|"), items)
        item_parts = items.split("|")
        self.assertGreaterEqual(int(item_parts[1]), 2, items)
        self.assertGreaterEqual(int(item_parts[2]), 1, items)
        self.assertGreaterEqual(int(item_parts[3]), 1, items)
        self.assertGreaterEqual(int(item_parts[4]), 1, items)
        self.assertGreaterEqual(int(item_parts[5]), 1, items)

        result = self.run_macro("TestChapterReferenceCode")
        self.assertTrue(result.startswith("ok|"), result)
        self.assertIn("REF THU_CH_0001", result)
        self.assertIn(chr(92) + "n", result)
        self.assertNotIn("STYLEREF", result)

    def test_chapter_reference_targets_selected_chapter(self) -> None:
        """检查章引用指向所选章。"""
        result = self.run_macro("TestChapterReferenceSelection")
        self.assertTrue(result.startswith("ok|"), result)
        parts = result.split("|")
        self.assertEqual(len(parts), 5, result)
        self.assertEqual(parts[1], parts[2], result)
        self.assertNotEqual(parts[2], parts[3], result)
        self.assertEqual(parts[4], "True", result)

    def test_repair_legacy_appendix_reset_fields(self) -> None:
        """检查修复旧版附录编号域。"""
        result = self.run_macro("TestRepairAppendixResetFields")
        self.assertTrue(result.startswith("ok|"), result)
        parts = result.split("|")
        self.assertGreaterEqual(int(parts[1]), 1, result)
        self.assertEqual(parts[2], "True", result)
        self.assertIn("\\r 0 \\h", result)

    def test_heading_styles(self) -> None:
        """检查标题样式数量。"""
        self.assertEqual(self.run_macro("TestHeadingStyles"), "ok|8")

    def test_formula_reference_existing_numbered(self) -> None:
        """检查引用已有编号公式。"""
        result = self.run_macro("TestFormulaReferenceExisting")
        self.assertTrue(result.startswith("ok|"), result)
        self.assertIn("THU_EQ_", result)
        self.assertIn("4.1", result)

    def test_formula_reference_is_field(self) -> None:
        """检查公式引用以域实现。"""
        self.assertIn("ok=True", self.run_macro("TestFormulaNormalize"))
        result = self.run_macro("TestFormulaRefText")
        self.assertIn("式(", result)
        self.assertIn("|fields=1", result)
        self.assertIn("|font=XITS Math", result)

    def test_body_style_selection(self) -> None:
        """检查正文样式选择候选数量。"""
        result = self.run_macro("TestBodyStyleSelection")
        parts = result.split("|")
        self.assertEqual(parts[0], "ok", result)
        self.assertGreaterEqual(int(parts[1]), 3, result)
        self.assertEqual(parts[2], "True", result)

    def test_footnote_insert(self) -> None:
        """检查脚注插入结果。"""
        result = self.run_macro("TestFootnoteInsert")
        parts = result.split("|")
        self.assertEqual(parts[0], "ok", result)
        self.assertEqual(parts[1], "True", result)
        self.assertEqual(parts[2], "True", result)

    def test_apply_emphasis(self) -> None:
        """选中文字后，三种强调入口应分别套用对应的字符样式。"""
        self.assertEqual(self.run_macro("TestApplyEmphasis"), "ok|3")

    def test_code_and_lists(self) -> None:
        """检查代码与列表宏结果。"""
        result = self.run_macro("TestCodeAndList")
        self.assertEqual(result, "ok|True|True|True|True|True")

    def test_list_numbering_adjust(self) -> None:
        """重新开始于1 后，继续编号应恢复到原来的编号。"""
        result = self.run_macro("TestListNumberingAdjust")
        parts = result.split("|")
        self.assertEqual(parts[0], "ok", result)
        self.assertEqual(parts[2], "1,2,3", result)
        self.assertEqual(parts[1], parts[3], result)
        self.assertEqual(parts[4], "True", result)
        self.assertEqual(parts[5], "True", result)
        self.assertEqual(parts[6], "True", result)

    def test_symbol_entry(self) -> None:
        """缩略语：空行写入、有文字补制表符、已有制表符只套样式。"""
        result = self.run_macro("TestSymbolEntry")
        parts = result.split("|")
        self.assertEqual(parts[0], "ok", result)
        self.assertEqual(parts[1], "缩略语\t说明", result)
        self.assertEqual(parts[2], "CPU\t说明", result)
        self.assertEqual(parts[3], "GPU\t图形处理器", result)
        self.assertEqual(parts[4], "True", result)

    def test_appendix_numbering(self) -> None:
        """检查附录各级标题编号。"""
        result = self.run_macro("TestAppendixNumbering")
        self.assertTrue(result.startswith("ok|"), result)
        self.assertIn("附录一级节标题=B.1", result)
        self.assertIn("附录二级节标题=B.1.1", result)

    def test_figure_caption_without_image(self) -> None:
        """检查无图时插入图注的样式。"""
        result = self.run_macro("TestFigureCaptionTextOnly")
        parts = result.split("|")
        self.assertEqual(parts[0], "ok", result)
        self.assertIn(parts[1], ("图注", "附录图注"), result)
        self.assertTrue(parts[2].startswith("图"), result)

    def test_table_caption_without_table(self) -> None:
        """检查无表时插入表注的样式。"""
        result = self.run_macro("TestTableCaptionTextOnly")
        parts = result.split("|")
        self.assertEqual(parts[0], "ok", result)
        self.assertIn(parts[1], ("表注", "附录表注"), result)
        self.assertTrue(parts[2].startswith("表"), result)


class UndergraduateWordTest(_CoverSecurityChecks, _AppendixSequenceChecks, WordDotmTest):
    """本科 dotm 的宏级检查：Word 中确认页眉为空、页码编排正确。"""

    variant = "undergraduate"

    def test_header_empty_and_footer_numbered(self) -> None:
        """Word 中确认页眉为空、前两节无页码、正文页脚有页码。"""
        for index in (1, 2):
            self.assertEqual(self.doc.Sections(index).Headers(1).Range.Text.strip(), "")
            self.assertEqual(self.doc.Sections(index).Footers(1).Range.Text.strip(), "")
        front = self.doc.Sections(3).Footers(1)
        self.assertRegex(front.Range.Text.strip(), r"[IVXLCDM]+")
        self.assertEqual(front.Range.ParagraphFormat.Alignment, 1)  # 居中对齐（wdAlignParagraphCenter）
        body = self.doc.Sections(self.doc.Sections.Count).Footers(1)
        self.assertRegex(body.Range.Text.strip(), r"\d+")
        self.assertEqual(body.Range.ParagraphFormat.Alignment, 1)

    def test_template_check(self) -> None:
        """本科变体的检查宏同样不应出现题注或引用错误。"""
        result = self.run_macro("TestTemplateCheck")
        self.assertEqual(result, "图注=0|表注=0|引用=0|提示=1")  # 表4.3由范围引用涵盖


if __name__ == "__main__":
    unittest.main(verbosity=2)
