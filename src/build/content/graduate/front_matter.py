# -*- coding: utf-8 -*-
"""前置部分 内容。构建器通过ctx提供doc、样式与版式helper。"""
from __future__ import annotations

from types import ModuleType

from build.content.examples import write_abstracts, write_symbol_entries


def build(ctx: ModuleType) -> None:
    # --- 名单页 ------------------------------------------------------------
    """写入研究生版前置部分：名单页、授权说明、摘要与目录等。"""
    ctx.chapter_section(None, False)
    ctx.H("学位论文指导小组、公开评阅人和答辩委员会名单", "章标题（无编号、不进目录）")

    def _borderless_table(rows, cols, cell_texts):
        """名单页用无框线表格排版，单元格文字套默认“表格文字”样式。"""
        t = ctx.doc .add_table(rows=rows, cols=cols)
        t .alignment = ctx.WD_TABLE_ALIGNMENT .CENTER
        t .autofit = False
        tblPr = t ._tbl .tblPr
        borders = ctx.OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = ctx.OxmlElement(f"w:{edge}")
            e .set(ctx.qn("w:val"), "none")
            e .set(ctx.qn("w:sz"), "0")
            e .set(ctx.qn("w:space"), "0")
            borders .append(e)
        tblPr .append(borders)
        for row, values in zip(t .rows, cell_texts):
            for cell, value in zip(row .cells, values):
                cell .width = ctx.Cm(15.0 / cols)
                cell .vertical_alignment = ctx.WD_ALIGN_VERTICAL .CENTER
                para = cell .paragraphs[0]
                para .style = ctx.get_style("tableText")
                # 名单页表格内容按指南要求为小四（12 pt）："表格文字"样式是
                # 11 pt（供正文示例表格用），这里只对名单页做直接格式覆盖，
                # 不新增样式键。
                _run = para .add_run(value)
                _run .font .size = ctx.Pt(12)

    for sub in ("指导小组名单", "公开评阅人名单", "答辩委员会名单"):
        ctx.H(sub, "committeeSectionTitle")
        if sub == "答辩委员会名单":
            roles = ("主席", "委员", "委员", "秘书")
            rows = [[roles[i], "×××", "×××", "×××"]for i in range(4)]
            _borderless_table(4, 4, rows)
        else:
            rows = [["×××", "×××", "×××"]for _ in range(4)]
            _borderless_table(4, 3, rows)

            # --- 授权说明 --------------------------------------------------------
    ctx.chapter_section(None, False)
    ctx.H("关于学位论文使用授权的说明", "授权说明标题")
    ctx.P("本人完全了解清华大学有关保留、使用学位论文的规定，即：", "授权说明文字")
    ctx.P("清华大学拥有在著作权法规定范围内学位论文的使用权，其中包括但不限于：（1）已获学位的研究生"
          "必须按学校规定提交学位论文，学校可以采用影印、缩印或其他复制手段保存研究生上交的"
          "学位论文；（2）为教学和科研目的，学校可以将公开的学位论文作为资料在图书馆、资料室"
          "等场所供校内师生阅读，或在校园网上供校内师生浏览部分内容；（3）根据《中华人民共和国"
          "学位法》及上级教育主管部门要求，报送相应的学位论文。", "授权说明文字")
    ctx.P("本人保证遵守上述规定。", "授权说明文字")
    ctx.P("", "授权说明文字")
    p = ctx.P("", "授权说明文字")
    p .add_run("作者签名：＿＿＿＿＿＿　　导师签名：＿＿＿＿＿＿")
    p = ctx.P("", "授权说明文字")
    p .add_run("日　　期：＿＿＿＿＿＿　　日　　期：＿＿＿＿＿＿")

    # --- 前置部分：每个章级标题各自起一节 --------------------------------
    write_abstracts(ctx, "章标题（无编号）",
                    "本模板将研究生学位论文常用的封面、摘要、目录、正文、图表、公式、参考文献和附录编排为可直接编辑的文档部件。使用者可从DOTM新建写作稿，通过“论文”功能区的宏设置章节和样式、插入图表与公式、生成编号和交叉引用，并在定稿时更新域、检查引用和导出提交文件。宏将重复的排版操作集中到按钮中，便于保持全文格式和编号一致。DOCX示例保留相同的内容部件和样式，也可通过Word内置功能编辑。本文概述两种使用方式，并展示相应的排版效果。",
                    "This template provides editable components for graduate theses, including title pages, abstracts, contents, chapters, figures, tables, equations, references and appendices. Users can create a document from the DOTM template and use macros on the Thesis ribbon to apply styles, organize chapters, insert objects, and manage numbering and cross-references. Field updates, reference checks and export commands support preparation of the final document. These macros simplify repeated formatting tasks and help maintain consistent styles and numbering. The DOCX example contains the same components and styles for editing with Word's built-in tools. This document outlines both workflows and demonstrates the resulting layout.",
                    "学位论文；Word模板；宏；自动化排版",
                    "graduate theses; Word template; macros; automated formatting")

    ctx.chapter_section("unchaptered", True, "upperRoman")
    ctx.H("目　录", ctx.RULES["documentParts"]["tableOfContents"]["heading"]["style"])
    p = ctx.doc .add_paragraph()
    p .paragraph_format .alignment = ctx.WD_ALIGN_PARAGRAPH .LEFT
    ctx.add_field(p, 'TOC \\o "1-3" \\t "章标题（无编号）,1,附录标题,1" \\h \\z \\u',
                  "（右键“更新域”生成目录）")

    ctx.chapter_section("unchaptered", True, "upperRoman")
    ctx.H("插图清单", "章标题（无编号）")
    p = ctx.doc .add_paragraph()
    ctx.add_field(p, 'TOC \\h \\z \\c "图"', "（更新域后生成插图清单）")

    ctx.chapter_section("unchaptered", True, "upperRoman")
    ctx.H("附表清单", "章标题（无编号）")
    p = ctx.doc .add_paragraph()
    ctx.add_field(p, 'TOC \\h \\z \\c "表"', "（更新域后生成附表清单）")

    ctx.chapter_section("unchaptered", True, "upperRoman")
    ctx.H("符号和缩略语说明", "章标题（无编号）")
    write_symbol_entries(ctx)
