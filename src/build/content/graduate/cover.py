# -*- coding: utf-8 -*-
"""封面 内容。构建器通过ctx提供doc、样式与版式helper。"""
from __future__ import annotations

from types import ModuleType

from build.content import cover_example
from build.degree_catalog import (
    cover_defaults,
    degree_display_options,
    degree_english,
    discipline_options,
    load_degree_config,
    parse_degree_display,
)


def cover_initial_values(config):
    """封面默认值：学位下拉、学科候选、英文学位名。"""
    defaults = cover_defaults(config)
    level, dtype, category = parse_degree_display(defaults["degree_cn"])
    return {
        "degree": defaults["degree_cn"],
        "degree_options": degree_display_options(config),
        "level": level,
        "degree_type": dtype,
        "category": category,
        "discipline": defaults["discipline"],
        "discipline_en": defaults["discipline_en"],
        "disciplines": discipline_options(config, category, level),
        "degree_en": degree_english(config, category, level),
    }


def build(ctx: ModuleType) -> None:
    """按研究生规范写入中英文封面与学位信息控件。"""
    _initial = cover_initial_values(load_degree_config())
    sec = ctx.doc .sections[0]
    # 指南 2.3.1：上/下 6.0cm，左/右 4.0cm。A4 正文宽 = 21 - 8 = 13cm = 368.5pt。
    # 标题从版心顶部起排，密级独立放入页眉。
    margins = ctx.section_margins("chineseCover")
    ctx.setup_section(sec, ctx.margin_cm(margins, "top"), ctx.margin_cm(margins, "bottom"),
                      ctx.margin_cm(margins, "left"), ctx.margin_cm(margins, "right"),
                      ctx.margin_cm(margins, "header"), ctx.margin_cm(margins, "footer"),
                      None, False, gutter=ctx.gutter_cm(margins))
    ctx.add_cover_security_header(sec)

    # 0 题目 + 学位说明，1 空行，2 作者信息（6×26pt），3 空行，4 日期。
    # 作者信息与日期按官方封面位置校准，不随密级有无或标题行数变化。
    # 题目与学位申请句同格：无论题目排一行还是两行，申请句都紧接题目之后。
    layout = ctx.RULES["cover"]["chinese"]["layout"]
    _zh_rows = tuple(row["height"]["value"] for row in layout["rows"])
    _zh = ctx.add_cover_layout_table(_zh_rows, layout["width"]["value"])

    _title_cell = _zh .rows[0].cells[0]
    p = ctx.cover_cell_paragraph(_title_cell, "coverTitleChinese",
                                 align=ctx.WD_ALIGN_PARAGRAPH .CENTER)
    ctx.add_content_control(p, "清华大学研究生学位论文\nWord模板", "论文题目",
                            kind="richText", showing_placeholder=False)

    # 学位申请句放在题目同一单元格的下一段，紧接题目，不单独占行。
    p = _title_cell .add_paragraph()
    ctx.format_cover_cell_paragraph(
        p,
        "coverDegreeInfo",
        align=ctx.WD_ALIGN_PARAGRAPH .CENTER,
        line_pt=22,
        before=16,
        after=16)
    p .add_run("(申请清华大学")
    ctx.add_content_control(
        p,
        _initial["degree"],
        "学位类别",
        kind="dropdownList",
        items=_initial["degree_options"],
        showing_placeholder=False)
    p .add_run("论文)")

    _cover_rows = (
        ("培养单位", cover_example.DEPARTMENT, "培养单位"),
        ("学科", _initial["discipline"], "学科（专业学位改为专业领域）"),
        ("研究生", cover_example.AUTHOR, "研究生（专业学位改为申请人）"),
        ("指导教师", cover_example.SUPERVISOR, "指导教师"),
    )
    _info_cell = _zh .rows[2].cells[0]
    for _i, (_label, _text, _alias) in enumerate(_cover_rows):
        _p = _info_cell .paragraphs[0]if _i == 0 else _info_cell .add_paragraph(
        )
        if _alias == "学科（专业学位改为专业领域）":
            ctx.add_cover_info_paragraph(
                _p, _label, _text, _alias,
                kind="comboBox",
                items=_initial["disciplines"],
                showing_placeholder=False)
        else:
            ctx.add_cover_info_paragraph(_p, _label, _text, _alias, showing_placeholder=False)

        # 固定行高 156pt = 6 × 26pt；默认 4 行，余下 2 行用空段落补足。
    for _ in range(2):
        _p = _info_cell .add_paragraph()
        ctx.format_cover_cell_paragraph(_p, "coverAuthorInfo", line_pt=26)
        _p .paragraph_format .tab_stops .add_tab_stop(
            ctx.Pt(159), ctx.WD_TAB_ALIGNMENT .RIGHT)
        _p .paragraph_format .tab_stops .add_tab_stop(
            ctx.Pt(164), ctx.WD_TAB_ALIGNMENT .LEFT)

    p = ctx.cover_cell_paragraph(
        _zh .rows[4].cells[0],
        "coverDate",
        align=ctx.WD_ALIGN_PARAGRAPH .CENTER,
        line_pt=20,
        before=4,
        after=16)
    ctx.add_content_control(p, cover_example.DATE_CHINESE, "成文日期",
                            showing_placeholder=False)

    # --- section 2: 英文封面 -------------------------------------------------
    sec = ctx.doc.add_section(ctx.WD_SECTION.NEW_PAGE)
    # 指南 2.3.2：上 5.5cm、下 5.0cm、左/右 3.6cm；正文宽 13.8cm = 391.2pt。
    margins = ctx.section_margins("englishCover")
    ctx.setup_section(sec, ctx.margin_cm(margins, "top"), ctx.margin_cm(margins, "bottom"),
                      ctx.margin_cm(margins, "left"), ctx.margin_cm(margins, "right"),
                      ctx.margin_cm(margins, "header"), ctx.margin_cm(margins, "footer"),
                      None, False, gutter=ctx.gutter_cm(margins))

    # 0 题目，1 空行，2 学位块，3 作者，4 空行，5 导师，
    # 6 空行，7 日期
    layout = ctx.RULES["cover"]["english"]["layout"]
    _en_rows = tuple(row["height"]["value"] for row in layout["rows"])
    _en = ctx.add_cover_layout_table(_en_rows, layout["width"]["value"])

    p = ctx.cover_cell_paragraph(_en .rows[0].cells[0], "coverTitleEnglish",
                                 align=ctx.WD_ALIGN_PARAGRAPH .CENTER)
    ctx.add_content_control(p, "Tsinghua University Graduate Thesis Template for Word",
                            "英文论文题目", kind="richText", showing_placeholder=False)

    _degree_cell = _en .rows[2].cells[0]
    _cover_en = (
        (None, "Thesis submitted to", False, "英文提交声明"),
        ("Tsinghua University", None, True, None),
        ("in partial fulfillment of the requirement", None, False, None),
        (None, "for the degree of", False, "英文degree行"),
        (None, _initial["degree_en"], True, "英文学位名称"),
    )
    for _i, (_text, _ph, _bold, _alias) in enumerate(_cover_en):
        p = (_degree_cell .paragraphs[0]if _i == 0
             else _degree_cell .add_paragraph())
        ctx.format_cover_cell_paragraph(
            p,
            "coverInfoEnglish",
            align=ctx.WD_ALIGN_PARAGRAPH .CENTER,
            line_pt=28)
        if _text is not None:
            r = p .add_run(_text)
            if _bold:
                ctx.set_char_style(
                    r, "strong"if _alias == "英文学位名称"else "emphasis")
        else:
            ctx.add_content_control(
                p, _ph, _alias,
                char_style=("strong"if _alias == "英文学位名称"else "emphasis")
                if _bold else None,
                showing_placeholder=False)

    p = _degree_cell.add_paragraph()
    ctx.format_cover_cell_paragraph(p, "coverInfoEnglish",
                                    align=ctx.WD_ALIGN_PARAGRAPH.CENTER, line_pt=20)
    p.add_run("in")
    p = _degree_cell.add_paragraph()
    ctx.format_cover_cell_paragraph(p, "coverInfoEnglish",
                                    align=ctx.WD_ALIGN_PARAGRAPH.CENTER, line_pt=20)
    ctx.add_content_control(p, _initial["discipline_en"], "英文学科",
                            showing_placeholder=False)

    _author_cell = _en .rows[3].cells[0]
    # “by + 英文作者”之前空一行。
    p = ctx.format_cover_cell_paragraph(
        _author_cell .paragraphs[0],
        "coverInfoEnglish",
        align=ctx.WD_ALIGN_PARAGRAPH .CENTER,
        line_pt=14)
    p = _author_cell .add_paragraph()
    ctx.format_cover_cell_paragraph(
        p,
        "coverInfoEnglish",
        align=ctx.WD_ALIGN_PARAGRAPH .CENTER,
        line_pt=28)
    r = p .add_run("by")
    ctx.set_char_style(r, "subtleEmphasis")
    p = _author_cell .add_paragraph()
    ctx.format_cover_cell_paragraph(
        p,
        "coverInfoEnglish",
        align=ctx.WD_ALIGN_PARAGRAPH .CENTER,
        line_pt=28)
    ctx.add_content_control(p, cover_example.AUTHOR_ENGLISH, "英文作者",
                            char_style="strong", showing_placeholder=False)
    p = _author_cell.add_paragraph()
    ctx.format_cover_cell_paragraph(p, "coverInfoEnglish",
                                    align=ctx.WD_ALIGN_PARAGRAPH.CENTER, line_pt=20)
    ctx.add_content_control(p, "", "英文专业领域", kind="richText",
                            char_style="strong", showing_placeholder=False)

    p = ctx.cover_cell_paragraph(
        _en .rows[5].cells[0],
        "coverSupervisorEnglish",
        line_pt=28)
    p .paragraph_format .left_indent = ctx.Pt(0)
    p .paragraph_format .tab_stops .clear_all()
    # 清掉样式“英文封面导师信息”自带的 183.7/195.6pt 制表位，否则首个
    # 制表符会先落到样式的右制表位上，冒号与姓名之间只剩 5pt。
    p .paragraph_format .tab_stops .add_tab_stop(
        ctx.Pt(183.7), ctx.WD_TAB_ALIGNMENT .CLEAR)
    p .paragraph_format .tab_stops .add_tab_stop(
        ctx.Pt(195.6), ctx.WD_TAB_ALIGNMENT .CLEAR)
    p .paragraph_format .tab_stops .add_tab_stop(
        ctx.Pt(195.5), ctx.WD_TAB_ALIGNMENT .RIGHT)
    p .paragraph_format .tab_stops .add_tab_stop(
        ctx.Pt(200.5), ctx.WD_TAB_ALIGNMENT .LEFT)
    ctx.add_content_control(p, f"Thesis Supervisor: {cover_example.SUPERVISOR_ENGLISH}",
                            "指导教师（英文）", kind="richText",
                            font="Times New Roman", size=15,
                            parts=["\tThesis Supervisor:", "\t" + cover_example.SUPERVISOR_ENGLISH],
                            showing_placeholder=False)

    p = ctx.cover_cell_paragraph(
        _en .rows[7].cells[0],
        "coverInfoEnglish",
        align=ctx.WD_ALIGN_PARAGRAPH .CENTER,
        line_pt=20,
        before=10,
        after=20)
    ctx.add_content_control(p, cover_example.DATE_ENGLISH, "英文成文日期",
                            char_style="strong", showing_placeholder=False)

    # 空白行不再保留空白单元格：把每个 spacer 行合并到上方单元格。
    ctx.merge_spacer_rows(_zh, (1, 3))
    ctx.merge_spacer_rows(_en, (1, 4, 6))
