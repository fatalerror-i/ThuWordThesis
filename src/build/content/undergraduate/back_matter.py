# -*- coding: utf-8 -*-
"""后置部分 内容（本科综合论文训练）。

组成顺序（规范 2.1）：正文 → 参考文献 → 附录 → 致谢 → 声明 →
在学期间参加课题的研究成果（如有）。综合论文训练记录表按用户要求排除。

与研究生版差异：删去指导教师（小组）评语、答辩委员会决议书、个人简历；
声明改用本科规范固定文字（“综合论文训练论文”）；“在学期间参加课题的
研究成果”只给“学术论文”一条示例小节（成果分类属写作内容，模板不预设）。
"""
from __future__ import annotations

from types import ModuleType

from build.content.appendix import build as build_appendix


def build(ctx: ModuleType) -> None:
    # --- 参考文献 ----------------------------------------------------------
    """写入本科版后置部分：参考文献、附录、致谢与声明等。"""
    ctx.chapter_section("unchaptered", True)
    ctx.H("参考文献", "章标题（无编号）")
    ctx.ref_entry("清华大学教务处. 清华大学综合论文训练写作规范（试行）[Z]. 北京: 清华大学.", 1)
    ctx.ref_entry("国家质量监督检验检疫总局, 中国国家标准化管理委员会. 信息与文献 参考文献著录规则: "
                  "GB/T 7714—2015[S]. 北京: 中国标准出版社, 2015.", 2)
    ctx.ref_entry("清华大学TUNA协会. ThuThesis: 清华大学学位论文LaTeX模板[CP/OL]. "
                  "https://github.com/tuna/thuthesis.", 3)

    build_appendix(ctx)

    ctx.chapter_section("unchaptered", True)
    ctx.H("致　谢", "章标题（无编号）")
    ctx.P("致谢对象包括资助研究工作的基金、组织或个人，协助完成研究工作的组织或个人，"
          "在研究工作中提出重要建议或提供重要帮助的组织或个人以及给予转载和引用权的资料、"
          "图片、文献、研究思想和设想的所有者等。篇幅限一页。")

    # --- 声明（本科规范固定文字，规范 2.13）--------------------------------
    ctx.chapter_section("unchaptered", True)
    ctx.H("声　明", "章标题（无编号）")
    ctx.P("本人郑重声明：所呈交的综合论文训练论文，是本人在导师指导下，独立进行研究工作"
          "所取得的成果。尽我所知，除文中已经注明引用的内容外，本论文的研究成果不包含"
          "任何他人享有著作权的内容。对本论文所涉及的研究工作做出贡献的其他个人和集体，"
          "均已在文中以明确方式标明。")
    ctx.P("")
    p = ctx.P("")
    p.alignment = ctx.WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run("签　名：＿＿＿＿＿＿　　日　期：＿＿＿＿＿＿")

    # --- 在学期间参加课题的研究成果（规范 2.14）----------------------------
    # 模板只给“学术论文”一条示例小节；成果分类属写作内容，不在模板里预设。
    ctx.chapter_section("unchaptered", True)
    ctx.H("在学期间参加课题的研究成果", "章标题（无编号）")
    _p = ctx.doc.add_paragraph(style=ctx.get_style("章内无编号小节"))
    _p.add_run("学术论文：")
    _pp = ctx.doc.add_paragraph(style=ctx.get_style("referenceEntry"))
    ctx.set_paragraph_numbering(_pp, ctx.NUM["achievements"][1])
    _pp.add_run("作者. 论文题目[J]. 期刊名, 2026.")
