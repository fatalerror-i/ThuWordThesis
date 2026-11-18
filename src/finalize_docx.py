# -*- coding: utf-8 -*-
"""生成模板 docx，用 Word 更新域，后处理，导出 PDF。"""
import argparse
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

from office_metadata import sanitize_office_metadata, sanitize_pdf_metadata
from build.style_presentation import finalize_style_presentation
from build.variants import choices, resolve

SRC = Path(__file__).resolve().parent


def word_update_fields(docx: Path):
    """在 Word 里打开母版并更新目录、题注与公式编号域。"""
    import win32com.client as win32
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx), ReadOnly=False)
        # 给正文示例补一个真实脚注；模板生成阶段不依赖 Word 的脚注部件。
        if "【脚注示例】" in doc.Content.Text and doc.Footnotes.Count == 0:
            for p in doc.Paragraphs:
                if "【脚注示例】" not in p.Range.Text:
                    continue
                # 只删除标记文本本身；给整段 Range 赋文本会改写段落标记，
                # 使正文段继承前一标题的样式并进入目录。
                marker = p.Range
                find = marker.Find
                find.ClearFormatting()
                find.Text = "【脚注示例】"
                find.Forward = True
                find.Wrap = 0  # 到文档末尾停止（wdFindStop）
                if not find.Execute():
                    continue
                pos = marker.Start
                marker.Delete()
                doc.Footnotes.Add(
                    Range=doc.Range(pos, pos),
                    Text="脚注编号按页重新开始，正文标记为上标带圈数字。",
                )
                fn = doc.Footnotes(doc.Footnotes.Count)
                # 去掉 Word 从正文段落带入的左侧缩进/段后距；由“脚注文本”样式统一。
                pf = fn.Range.Paragraphs(1).Range.ParagraphFormat
                pf.LeftIndent = 13.5
                pf.FirstLineIndent = -13.5
                pf.CharacterUnitLeftIndent = 0
                pf.CharacterUnitFirstLineIndent = 0
                pf.SpaceBefore = 0
                pf.SpaceAfter = 0
                # 正文里的脚注标记单独设回上标。
                fn.Reference.Font.Superscript = True
                # 编号后只保留一个 0.5 字符宽的空格（U+2002）。
                # Word 在编号后会自带一个普通空格，它在 fn.Range 之外，
                # 需要把 Range 起点前移一格再删除。
                fr = fn.Range
                while fr.Start > 0:
                    sp = fr.Duplicate
                    sp.Start = sp.Start - 1
                    sp.End = sp.Start + 1
                    if sp.Text not in (" ", "\u2002", "\u00a0", "\t"):
                        break
                    sp.Delete()
                    fr = fn.Range
                fr.InsertBefore("\u2002")
                break
        # 给示例公式补 Word 原生 #() 编号：这是模板示例的一部分，必须和用户
        # 点“公式”宏后的结果一致。前缀域按所在位置取：正文用“STYLEREF 1 \s”
        # 得到章号（3.1），附录用“STYLEREF "附录标题" \s”得到附录字母（C.3），
        # 与宏 BuildNumberWithFields 的 bApp 分支同源。
        in_appendix = False
        for p in doc.Paragraphs:
            style_name = p.Style.NameLocal
            if style_name == "附录标题":
                in_appendix = True
            elif style_name == "章标题（无编号）":
                in_appendix = False
            if style_name != "公式":
                continue
            if p.Range.OMaths.Count < 1:
                continue
            if "#" in p.Range.Text:
                continue
            om = p.Range.OMaths(1)
            p.Style = doc.Styles("公式")
            p.Alignment = 1  # 居中对齐（wdAlignParagraphCenter）
            p.TabStops.ClearAll()
            p.FirstLineIndent = 0
            om.Linearize()
            if "#" in om.Range.Text:
                continue
            om.Range.Select()
            word.Selection.Collapse(0)  # 折叠到末尾（wdCollapseEnd）
            word.Selection.TypeText("#(")
            chapter_field = ('STYLEREF "附录标题" \s' if in_appendix
                             else "STYLEREF 1 \s")
            doc.Fields.Add(word.Selection.Range, -1, chapter_field, False)
            word.Selection.Collapse(0)
            word.Selection.TypeText(".")
            doc.Fields.Add(word.Selection.Range, -1, "SEQ 式 \* ARABIC \s 1", False)
            word.Selection.Collapse(0)
            word.Selection.TypeText(")")
            om.BuildUp()
            p.Range.Fields.Update()

        # 公式编号加书签，供正文的 REF 域引用。书签名用宏里的 BM_EQ_PREFIX
        # （THU_EQ_），宏重跑时按前缀匹配到已有书签，不会重复创建。
        try:
            eq_index = 0
            for p in doc.Paragraphs:
                if p.Style.NameLocal != "公式":
                    continue
                if p.Range.OMaths.Count < 1:
                    continue
                styleref = None
                seq = None
                for i in range(1, p.Range.Fields.Count + 1):
                    f = p.Range.Fields(i)
                    fc = f.Code.Text.strip().upper()
                    if fc.startswith("STYLEREF") and styleref is None:
                        styleref = f
                    elif fc.startswith("SEQ") and seq is None:
                        seq = f
                if styleref is None or seq is None:
                    continue
                eq_start = styleref.Result.Start
                eq_end = seq.Result.End
                if eq_end <= eq_start:
                    continue
                eq_index += 1
                bm_name = "THU_EQ_%04d" % eq_index
                if doc.Bookmarks.Exists(bm_name):
                    doc.Bookmarks(bm_name).Delete()
                doc.Bookmarks.Add(bm_name, doc.Range(eq_start, eq_end))
        except Exception as exc:  # 书签是尽力而为，失败不影响构建
            print("warning: formula bookmark skipped:", exc)

        for _ in range(2):
            for f in list(doc.Fields):
                code = f.Code.Text.strip()
                upper = code.upper()
                if upper.startswith(("SEQ", "STYLEREF")):
                    f.Update()
                elif upper.startswith("REF "):
                    # 只有书签存在时才更新，否则 Word 会把结果写成「错误!未定义书签」。
                    parts = code.split()
                    if len(parts) > 1 and doc.Bookmarks.Exists(parts[1]):
                        f.Update()
            for toc in doc.TablesOfContents:
                toc.Update()
            for tof in doc.TablesOfFigures:
                tof.Update()
            for sec in doc.Sections:
                sec.Headers(1).Range.Fields.Update()
                sec.Footers(1).Range.Fields.Update()
        # 脚注编号：①、②、③，每页重新编号。Word 只认对象模型里的设置，
        # 直接写 settings.xml 的 footnotePr 会在保存时被丢弃。
        doc.Footnotes.NumberingRule = 2  # 每页重排（wdRestartPage）
        doc.Footnotes.NumberStyle = 18   # 带圈数字（wdNoteNumberStyleNumberEnclosedCircle）
        doc.Save()
        doc.Close(False)
    finally:
        word.Quit()


def word_export_pdf(docx: Path, target: Path):
    """用 Word 把母版导出为预览 PDF。"""
    import win32com.client as win32
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx), ReadOnly=True, AddToRecentFiles=False)
        doc.ExportAsFixedFormat(OutputFileName=str(target), ExportFormat=17, IncludeDocProps=False)
        print("pages:", doc.ComputeStatistics(2))
        doc.Close(False)
    finally:
        word.Quit()


def main():
    parser = argparse.ArgumentParser(description="生成 DOCX 示例和 PDF 预览。")
    parser.add_argument("--variant", required=True, choices=choices())
    parser.add_argument("--docx", type=Path, help="DOCX 输出路径，默认取变体配置")
    parser.add_argument("--pdf", type=Path, help="PDF 输出路径，默认取变体配置")
    parser.add_argument("--for-dotm", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()

    variant = resolve(args.variant)
    docx = (args.docx or variant.docx).resolve()
    docx_args = ["--variant", variant.name, "--docx", str(docx)]
    subprocess.run(
        [sys.executable, str(SRC / "build_docx.py"), *docx_args, "--defer-style-presentation"],
        check=True,
    )
    # 第一次后处理修正编号，第二次更新域刷新目录缓存并重新套用格式。
    for _ in range(2):
        word_update_fields(docx)
        subprocess.run([sys.executable, str(SRC / "postprocess_docx.py"), *docx_args], check=True)

    if not args.for_dotm:
        finalize_style_presentation(docx, variant)
    sanitize_office_metadata(docx)

    if not args.for_dotm:
        pdf = (args.pdf or variant.pdf).resolve()
        pdf.parent.mkdir(parents=True, exist_ok=True)
        with TemporaryDirectory(dir=pdf.parent) as directory:
            preview = Path(directory) / pdf.name
            word_export_pdf(docx, preview)
            sanitize_pdf_metadata(preview)
            preview.replace(pdf)
    print("saved:", docx)


if __name__ == "__main__":
    main()
