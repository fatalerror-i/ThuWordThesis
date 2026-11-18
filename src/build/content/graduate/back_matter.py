# -*- coding: utf-8 -*-
"""后置部分（back matter）：参考文献、附录、致谢、声明、简历、评语、决议。"""
from __future__ import annotations

from types import ModuleType

from build.content.appendix import build as build_appendix


def build(ctx: ModuleType) -> None:
    # --- 参考文献 ----------------------------------------------------------
    """写入研究生版后置部分：参考文献、附录、致谢与声明等。"""
    ctx.chapter_section("unchaptered", True)
    ctx.H("参考文献", "章标题（无编号）")
    ctx.ref_entry("清华大学研究生院. 清华大学研究生学位论文写作指南: 202605[Z]. 北京: 清华大学, 2026.", 1)
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

    ctx.chapter_section("unchaptered", True)
    ctx.H("声　明", "章标题（无编号）")
    ctx.P("本人郑重声明：所呈交的学位论文，是本人在导师指导下，独立进行研究工作所取得的成果，不包含涉及国家秘密的内容。"
          "尽我所知，除文中已经注明引用的内容外，本学位论文的研究成果不包含任何他人享有著作权的内容。"
          "对本论文所涉及的研究工作做出贡献的其他个人和集体，均已在文中以明确方式标明。")
    ctx.P("")  # 声明文字与签字区之间的空行
    p = ctx.P("")
    p .alignment = ctx.WD_ALIGN_PARAGRAPH .RIGHT
    p .add_run("签　名：＿＿＿＿＿＿　　日　期：＿＿＿＿＿＿")

    ctx.chapter_section("unchaptered", True)
    ctx.H("个人简历、在学期间完成的相关学术成果", "章标题（无编号）")
    # 章内两个小节标题：个人简历、在学期间完成的相关学术成果。
    # 学术成果的具体分类属写作内容，模板只给一条示例，不预设分类。
    _p = ctx.doc .add_paragraph(style=ctx.get_style("章内无编号小节"))
    _p .add_run("个人简历")
    ctx.P("×××，×年×月生，×年×月于××大学获学士学位。")
    _p = ctx.doc .add_paragraph(style=ctx.get_style("章内无编号小节"))
    _p .add_run("在学期间完成的相关学术成果")
    _pp = ctx.doc .add_paragraph(style=ctx.get_style("referenceEntry"))
    ctx.set_paragraph_numbering(_pp, ctx.NUM["achievements"][1])
    _pp .add_run("作者. 论文题目[J]. 期刊名, 2026.")

    ctx.chapter_section("unchaptered", True)
    ctx.H("指导教师（小组）评语", "章标题（无编号）")
    ctx.P("（此处由指导教师或指导小组填写评语。）")

    ctx.chapter_section("unchaptered", True)
    ctx.H("答辩委员会决议书", "章标题（无编号）")
    ctx.P("（此处填写答辩委员会决议书。）")
