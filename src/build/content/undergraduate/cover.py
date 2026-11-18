# -*- coding: utf-8 -*-
"""封面 内容（本科综合论文训练）。

与研究生版不同：本科封面是静态单页，没有英文封面、没有学位类别下拉与
学科级联；密级控件放在版心顶部的独立表格行。其下为校徽 + 校名组合图（SVG），
下面依次是“综 合 论 文 训 练”、论文题目、系别/专业/姓名/指导教师、
成文日期。页边距与装订线从本科规则读取（上 3.8 / 下 3.2 / 左右 3.0 cm，
装订线 0.2 cm 在左）。
"""
from __future__ import annotations

import os
from types import ModuleType

from build.content import cover_example


def build(ctx: ModuleType) -> None:
    """按本科规范写入静态单页封面与校徽校名组合图。"""
    sec = ctx.doc.sections[0]
    margins = ctx.section_margins("chineseCover")
    ctx.setup_section(sec,
                      ctx.margin_cm(margins, "top"),
                      ctx.margin_cm(margins, "bottom"),
                      ctx.margin_cm(margins, "left"),
                      ctx.margin_cm(margins, "right"),
                      ctx.margin_cm(margins, "header"),
                      ctx.margin_cm(margins, "footer"),
                      None, False,
                      gutter=ctx.gutter_cm(margins))
    # 恢复官方封面的空白“页眉”样式；密级仍在版心顶部的表格行内。
    sec.header.paragraphs[0].style = ctx.get_style("header")
    # 规范 2.3.1：本科封面装订线 0.2 cm，位置在左；正文各节装订线为 0，
    # 由 setup_section 按各自 pageMargins 显式写入，不会从封面节继承。

    # 单列无框版式表。A4 正文宽 = 21 - 3.0 - 3.0 - 0.2(装订线) = 14.8cm
    # ≈ 419.5pt；表格占满版心。
    # 0 密级, 1 校徽+校名, 2 spacer, 3 “综合论文训练”, 4 spacer, 5 题目,
    # 6 spacer, 7 信息块(4行×26pt + 2空行), 8 spacer, 9 日期
    layout = ctx.RULES["cover"]["chinese"]["layout"]
    _rows = tuple(row["height"]["value"] for row in layout["rows"])
    _t = ctx.add_cover_layout_table(_rows, ctx.Cm(layout["width"]["value"]).pt)

    p = ctx.cover_cell_paragraph(_t.rows[0].cells[0], "coverSecurity")
    ctx.add_content_control(p, "", "密级", showing_placeholder=False)

    # 直接嵌入 assets/thu-cover-logo.svg 组合图，PNG 为兼容版本。
    # 依据《清华大学视觉形象识别系统手册》A-3-03，171 × 49 bp 的画布左侧为
    # 直径 49 bp 的校徽，右侧间隔 10 bp 放置 112 × 41 bp 的校名标准字，
    # 校名底部比校徽底部高 4.5 bp。封面插入尺寸为 6.03 × 1.73 cm。
    _assets = os.path.join(ctx.ROOT, "assets")
    p = ctx.cover_cell_paragraph(
        _t.rows[1].cells[0], "coverTitleChinese",
        align=ctx.WD_ALIGN_PARAGRAPH.CENTER, line_pt=80,
        before=layout["topSpace"]["value"])
    ctx.add_svg_picture(
        p,
        os.path.join(_assets, "thu-cover-logo.png"),
        os.path.join(_assets, "thu-cover-logo.svg"),
        ctx.Cm(6.03), ctx.Cm(1.73))

    # --- “综 合 论 文 训 练” ----------------------------------------------
    # 这是一行静态装饰性封面大题，不是正文内容，模板中不为其另立全局样式；
    # 以“封面论文题目”样式为底，仅在本段覆盖字号与加粗。
    p = ctx.cover_cell_paragraph(
        _t.rows[3].cells[0], "coverTitleChinese",
        align=ctx.WD_ALIGN_PARAGRAPH.CENTER, line_pt=44)
    r = p.add_run("综 合 论 文 训 练")
    r.bold = True
    r.font.size = ctx.Pt(36)

    # --- 论文题目（黑体一号，可换两行，1.25 倍行距）------------------------
    p = ctx.cover_cell_paragraph(
        _t.rows[5].cells[0], "coverTitleChinese",
        align=ctx.WD_ALIGN_PARAGRAPH.CENTER)
    ctx.add_content_control(p, "清华大学本科综合论文训练\nWord模板", "论文题目",
                            kind="richText", showing_placeholder=False)

    # --- 系别 / 专业 / 姓名 / 指导教师（仿宋三号）--------------------------
    _info = (
        ("系别", cover_example.DEPARTMENT, "系别"),
        ("专业", cover_example.DISCIPLINE, "专业"),
        ("姓名", cover_example.AUTHOR, "姓名"),
        ("指导教师", cover_example.SUPERVISOR, "指导教师"),
    )
    _info_cell = _t.rows[7].cells[0]
    for _i, (_label, _text, _alias) in enumerate(_info):
        _p = (_info_cell.paragraphs[0] if _i == 0
              else _info_cell.add_paragraph())
        ctx.add_cover_info_paragraph(_p, _label, _text, _alias, showing_placeholder=False)
    # 固定行高 156pt = 6 × 26pt；补 2 行空段落把下方日期推到版面下部。
    for _ in range(2):
        _p = _info_cell.add_paragraph()
        ctx.format_cover_cell_paragraph(_p, "coverAuthorInfo", line_pt=26)
        _p.paragraph_format.tab_stops.add_tab_stop(
            ctx.Pt(159), ctx.WD_TAB_ALIGNMENT.RIGHT)
        _p.paragraph_format.tab_stops.add_tab_stop(
            ctx.Pt(164), ctx.WD_TAB_ALIGNMENT.LEFT)

    info_rules = ctx.RULES["cover"]["chinese"]["authorInfo"]
    for _p in _info_cell.paragraphs:
        tabs = _p.paragraph_format.tab_stops
        tabs.clear_all()
        tabs.add_tab_stop(ctx.Pt(info_rules["labelRightTab"]["value"]), ctx.WD_TAB_ALIGNMENT.RIGHT)
        tabs.add_tab_stop(ctx.Pt(info_rules["valueLeftTab"]["value"]), ctx.WD_TAB_ALIGNMENT.LEFT)

    # --- 成文日期（宋体三号，不用阿拉伯数字）------------------------------
    p = ctx.cover_cell_paragraph(
        _t.rows[9].cells[0], "coverDate",
        align=ctx.WD_ALIGN_PARAGRAPH.CENTER,
        line_pt=24, before=4, after=12)
    ctx.add_content_control(p, cover_example.DATE_CHINESE, "成文日期",
                            showing_placeholder=False)

    # 空白 spacer 行合并到上方单元格，避免留下空表格行。
    ctx.merge_spacer_rows(_t, (2, 4, 6, 8))
