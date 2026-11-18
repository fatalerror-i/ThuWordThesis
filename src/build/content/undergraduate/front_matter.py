# -*- coding: utf-8 -*-
r"""前置部分 内容（本科综合论文训练）。

组成顺序（规范 2.1）：封面 → 授权说明 → 摘要 → Abstract → 目录 →
插图清单 → 附表清单 → 符号和缩略语说明 → 正文。

与研究生版差异：删去名单页；授权说明压成规范规定的 1 段，签名与日期
用两段纯文本（不用表格，与研究生版一致）；目录域去掉 \t 开关
（官方样本 `TOC \o "1-3" \h \z \u`）；
图/表清单改用 \c "图"/\c "表" SEQ 题注域。摘要、Abstract、符号说明
沿用研究生版结构与样式（用户明确正文可复用研究生模板）。
"""
from __future__ import annotations

from types import ModuleType

from build.content.examples import write_abstracts, write_symbol_entries


def build(ctx: ModuleType) -> None:
    # --- 授权说明（封面后单设一页，规范 2.2）------------------------------
    """写入本科版前置部分：授权说明、摘要、目录与符号说明等。"""
    ctx.chapter_section(None, False)
    ctx.H("关于论文使用授权的说明", "授权说明标题")
    ctx.P("本人完全了解清华大学有关保留、使用综合论文训练论文的规定，即：学校有权保留"
          "论文的复印件，允许论文被查阅和借阅；学校可以公布论文的全部或部分内容，"
          "可以采用影印、缩印或其他复制手段保存论文。", "授权说明文字")
    # 提交时作者与指导教师必须签名：两段纯文本，不用表格（与研究生版一致）。
    ctx.P("", "授权说明文字")
    p = ctx.P("", "授权说明文字")
    p.add_run("作者签名：＿＿＿＿＿＿　　导师签名：＿＿＿＿＿＿")
    p = ctx.P("", "授权说明文字")
    p.add_run("日　　期：＿＿＿＿＿＿　　日　　期：＿＿＿＿＿＿")

    # --- 摘要 --------------------------------------------------------------
    write_abstracts(ctx, "章标题（无编号、不进目录）",
                    "本模板将本科综合论文训练论文常用的封面、摘要、目录、正文、图表、公式、参考文献和附录编排为可直接编辑的文档部件。使用者可从DOTM新建写作稿，通过“论文”功能区的宏设置章节和样式、插入图表与公式、生成编号和交叉引用，并在定稿时更新域、检查引用和导出提交文件。宏将重复的排版操作集中到按钮中，便于保持全文格式和编号一致。DOCX示例保留相同的内容部件和样式，也可通过Word内置功能编辑。本文概述两种使用方式，并展示相应的排版效果。",
                    "This template provides editable components for undergraduate theses, including a cover, abstracts, contents, chapters, figures, tables, equations, references and appendices. Users can create a document from the DOTM template and use macros on the Thesis ribbon to apply styles, organize chapters, insert objects, and manage numbering and cross-references. Field updates, reference checks and export commands support preparation of the final document. These macros simplify repeated formatting tasks and help maintain consistent styles and numbering. The DOCX example contains the same components and styles for editing with Word's built-in tools. This document outlines both workflows and demonstrates the resulting layout.",
                    "综合论文训练；Word模板；宏；自动化排版",
                    "undergraduate theses; Word template; macros; automated formatting")

    # --- 目录（官方样本：TOC \o "1-3" \h \z \u，无 \t 开关）---------------
    ctx.chapter_section("unchaptered", True, "upperRoman")
    heading_rule = ctx.RULES["documentParts"]["tableOfContents"]["heading"]
    heading = ctx.H("目　录", heading_rule["style"])
    # 本科目录仍从正文开始。保留章标题样式，仅覆盖本段的大纲级别。
    if heading_rule.get("outlineLevel") == "bodyText":
        ctx._child(heading._p.get_or_add_pPr(), "w:outlineLvl", val="9")
    p = ctx.doc.add_paragraph()
    p.paragraph_format.alignment = ctx.WD_ALIGN_PARAGRAPH.LEFT
    ctx.add_field(p, 'TOC \\o "1-3" \\h \\z \\u',
                  "（右键“更新域”生成目录）")

    # --- 插图清单 / 附表清单（官方样本用 SEQ 题注 \c 开关）----------------
    ctx.chapter_section("unchaptered", True, "upperRoman")
    ctx.H("插图清单", "章标题（无编号、不进目录）")
    p = ctx.doc.add_paragraph()
    ctx.add_field(p, 'TOC \\h \\z \\c "图"', "（更新域后生成插图清单）")

    ctx.chapter_section("unchaptered", True, "upperRoman")
    ctx.H("附表清单", "章标题（无编号、不进目录）")
    p = ctx.doc.add_paragraph()
    ctx.add_field(p, 'TOC \\h \\z \\c "表"', "（更新域后生成附表清单）")

    # --- 符号和缩略语说明 --------------------------------------------------
    ctx.chapter_section("unchaptered", True, "upperRoman")
    ctx.H("符号和缩略语说明", "章标题（无编号、不进目录）")
    write_symbol_entries(ctx)
