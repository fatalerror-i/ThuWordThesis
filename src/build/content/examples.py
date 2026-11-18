# -*- coding: utf-8 -*-
"""两种模板共用的摘要布局和正文示例。"""
from __future__ import annotations

from types import ModuleType


def write_abstracts(ctx: ModuleType, title_style: str, chinese: str, english: str,
                    keywords_cn: str, keywords_en: str) -> None:
    ctx.chapter_section("unchaptered", True, "upperRoman", 1)
    ctx.H("摘　要", title_style)
    ctx.P(chinese)
    p = ctx.P()
    ctx.set_char_style(p.add_run("关键词："), "subtleEmphasis")
    p.add_run(keywords_cn)

    ctx.chapter_section("unchaptered", True, "upperRoman")
    ctx.H("Abstract", title_style)
    ctx.P(english)
    p = ctx.P()
    ctx.set_char_style(p.add_run("Keywords: "), "strong")
    p.add_run(keywords_en)

def write_symbol_entries(ctx: ModuleType) -> None:
    # 先列质能关系中的物理量，再列写作文件与纸张规格。
    for symbol, description in (("E", "能量"), ("m", "质量"), ("c", "真空中的光速")):
        p = ctx.P(style="符号说明条目")
        ctx.add_inline_formula(p, symbol)
        p.add_run("\t" + description)
    for symbol, description in (
        ("DOTM", "支持宏的Word模板格式"),
        ("DOCX", "Word文档格式"),
        ("PDF", "固定版式文档格式"),
        ("A4", "纸张幅面为210毫米×297毫米"),
    ):
        ctx.P(symbol + "\t" + description, "符号说明条目")


def _table(ctx: ModuleType, title: str, header, rows, widths) -> None:
    """写入带编号题注的三线表。"""
    ctx.add_caption("tableCaption", "表", "1", "表", title)
    table = ctx.three_line_table(len(rows) + 1, len(header))
    for i, row in enumerate((header, *rows)):
        for j, text in enumerate(row):
            ctx.cell_text(table.rows[i].cells[j], text, 11)
    ctx.set_table_widths(table, widths)


def build(ctx: ModuleType) -> None:
    _introduction(ctx)
    _macro_workflow(ctx)
    _docx_workflow(ctx)
    _format_examples(ctx)


def _introduction(ctx: ModuleType) -> None:
    ctx.chapter_section("numbered", True, "decimal", 1)
    ctx.H("引言", "chapterTitle")
    p = ctx.P("写作时可直接使用模板中的章节、图表和引文样式。论文格式以对应的写作指南或规范为准")
    ctx.add_citation(p, 1)
    p.add_run("。参考文献示例采用GB/T 7714—2015的著录格式")
    ctx.add_citation(p, 2)
    p.add_run("。")

    ctx.H("文件用途", "sectionTitle1")
    p = ctx.P()
    ctx.ref_caption(p, "表", "表")
    p.add_run("列出写作与提交所用的文件。提交格式以院系要求为准。")
    _table(ctx, "写作与提交文件的用途", ("文件格式", "用途"),
           (("DOTM", "从模板新建写作稿，并提供“论文”选项卡。"),
            ("DOCX", "编辑、审阅或提交论文。"),
            ("PDF", "核对页面编排或提交定稿。")),
           widths=(3.4, 11.6))


def _macro_workflow(ctx: ModuleType) -> None:
    ctx.chapter_section("numbered", True, "decimal")
    ctx.H("DOTM宏模板的使用", "chapterTitle")
    ctx.P("模板已包含论文写作常用的封面、章节、图表、公式、引文等文档部件，"
          "并在“论文”功能区提供相应的自动化宏。以下按功能组介绍主要操作，详细步骤见项目《快速入门》。")
    ctx.H("新建写作稿", "sectionTitle1")
    ctx.P("双击对应的DOTM模板新建文档，将新文档保存为DOCX，再关闭并重新打开。"
          "首次启用宏和更换模板的步骤见项目《快速入门》。")

    ctx.H("封面", "sectionTitle1")
    ctx.P("研究生版可用“刷新封面”更新学位与日期相关信息，用“选填信息”增删副导师、联合导师和密级信息。"
          "本科版直接填写封面信息，用“选填信息”填写密级。密级在本科封面首行或研究生封面页眉，默认留空，启用示例为“秘密★5年”，可按实际要求修改。")

    ctx.H("段落", "sectionTitle1")
    ctx.P("通过“正文章节”“附录章节”和“其他章节”设置章标题、节标题及摘要、致谢等部分的标题，"
          "章号和节号随样式生成。章级标题会另起一节，“正文”按钮可将选中段落恢复为正文样式。")

    ctx.H("字体", "sectionTitle1")
    ctx.P("选中文字后，在“强调”菜单中选择“不明显强调”“强调”或“明显强调”，统一应用对应的字符样式。")

    ctx.H("列表", "sectionTitle1")
    ctx.P("“有序列表”和“无序列表”提供一至三级列表样式，“调整编号”可接续编号或重新从1开始。"
          "“参考文献”可设置顺序编码制或著者出版年制条目，“缩略语”可整理符号及其说明。")

    ctx.H("图片", "sectionTitle1")
    ctx.P("用“插入图片”添加嵌入型图片并应用图片样式，用“图注”在图下生成带章号的题注。"
          "多幅分图可用“子图注”添加字母编号，再为整组图片添加总图注。")

    ctx.H("表格", "sectionTitle1")
    ctx.P("用“插入表格”生成三线表，用“表注”在表上添加带章号的题注。"
          "跨页表格重复表头，保留同一张表。")

    ctx.H("公式", "sectionTitle1")
    ctx.P("“行内公式”插入或调整正文中的公式。“行间公式”可选“有编号”或“无编号”，"
          "前者生成编号与引用书签，后者使用同样的居中公式样式。")

    ctx.H("代码", "sectionTitle1")
    ctx.P("“行内代码”用于文件名、命令和短代码，“行间代码”用于独立程序段，"
          "可统一字体、段落格式和代码行号。")

    ctx.H("引用", "sectionTitle1")
    ctx.P("“引用内容”插入图、表、公式或章节的交叉引用，“引用文献”插入一条或多条上标引文，"
          "“脚注”添加页下说明。“更新域”刷新目录、题注、公式编号和引用，"
          "“引用检查”核对图表题注及引用目标，并列出未引用的图表提示。")

    ctx.H("导出", "sectionTitle1")
    ctx.P("完成编辑后先更新域并检查引用，保存写作稿，再通过“导出”生成DOCX或PDF提交文件。"
          "导出的DOCX将文献引用固定为当时的编号，后续修改可继续在写作稿中进行。")

    ctx.H("帮助", "sectionTitle1")
    ctx.P("“关于”显示模板名称和版本。完整按钮说明和常见问题见项目用户文档。")


def _docx_workflow(ctx: ModuleType) -> None:
    ctx.chapter_section("numbered", True, "decimal")
    ctx.H("DOCX模板的使用", "chapterTitle")
    ctx.P("打开DOCX示例并另存为写作稿，替换摘要、关键词、正文和参考文献。"
          "文档已保存常用部件的样式、编号和域，以下介绍通过Word内置功能完成相应编辑的方法。")

    ctx.H("封面与论文结构", "sectionTitle1")
    ctx.P("在封面内容控件中填写题目、作者、导师和日期；研究生另填英文封面，并核对中英文信息。"
          "按院系要求填写授权说明、声明和成果页，保留所需的清单及附录，完成签名和日期。")

    ctx.H("段落与字体", "sectionTitle1")
    ctx.P("在“开始→样式”中选择“正文”“章标题（有编号）”、各级节标题或附录标题等样式，沿用模板的多级编号。"
          "新章前通过“布局→分隔符→分节符（下一页）”另起一节，并核对页眉和页码。"
          "选中文字可应用“不明显强调”“强调”或“明显强调”字符样式。")

    ctx.H("列表与符号说明", "sectionTitle1")
    ctx.P("在样式窗格中选择模板的编号列表或项目符号列表样式，右键编号可继续编号或设置起始值。"
          "符号说明按“符号、制表符、说明”输入，再应用“符号说明条目”样式，使续行与说明列对齐。")

    ctx.H("图片与图注", "sectionTitle1")
    ctx.P("通过“插入→图片”添加图片，选择“嵌入型”布局并为所在段落应用“图片”样式。"
          "选中图片后用“引用→插入题注”，标签选“图”，位置选图片下方；"
          "在“编号”中包含章节号，章节起始样式选“章标题（有编号）”，分隔符选小数点，再应用“图注”样式。附录图表可复制已有附录题注，保留其编号域并更新，引用时另建唯一书签。"
          "分图题可手动填写字母标记并应用“子图注”样式，整组图片另设总图注。")

    ctx.H("表格", "sectionTitle1")
    ctx.P("通过“插入→表格”创建表格，在“表格设计”中应用模板的“三线表”样式，"
          "单元格文字套“表格文字”。用“引用→插入题注”添加标签为“表”、位于表上方且包含章号的题注，"
          "再应用“表注”样式。跨页表格在“表格布局”中设置重复标题行，"
          "保留同一张表并核对表题、表头与首条数据是否同页。")

    ctx.H("公式与编号", "sectionTitle1")
    ctx.P("按Alt+=插入Word公式，段内公式使用“行内公式”字符样式，独立公式所在段落应用“公式”样式。"
          "需要编号时可复制第4章的带编号公式，保留编号域并替换公式内容，为新编号另建唯一书签。"
          "也可在“引用→插入题注”中新建“式”标签、设置包含章号的编号，再按公式样例调整位置。"
          "引用编号前，用“插入→书签”为该编号建立唯一书签。")

    ctx.H("代码", "sectionTitle1")
    ctx.P("在样式窗格中为短代码应用“行内代码”字符样式，为独立程序段应用“行间代码”段落样式。"
          "复制程序时保留缩进，代码行号随模板列表生成。")

    ctx.H("引用、脚注与参考文献", "sectionTitle1")
    ctx.P("用“引用→交叉引用”选择图、表、标题或书签，插入自动引用。"
          "参考文献可应用顺序编码制或著者出版年制条目样式，顺序编码制引文可通过交叉引用插入编号项，"
          "再设为上标。用“引用→插入脚注”添加页下说明，按模板要求设置编号样式和每页重新编号。")

    ctx.H("目录、更新与导出", "sectionTitle1")
    ctx.P("右键目录选择“更新域→更新整个目录”，图表清单也分别更新。正文可按Ctrl+A、F9更新域，"
          "页眉、页脚和脚注中的域分别检查。核对引用结果和分页后，用“文件→另存为”保存DOCX，"
          "或通过“文件→导出”生成PDF提交文件。")


def _format_examples(ctx: ModuleType) -> None:
    ctx.chapter_section("numbered", True, "decimal")
    ctx.H("模板格式与使用范例", "chapterTitle")
    ctx.P("下文展示常用的标题、图表、公式和引文样式。具体字号、间距和两版差异可查项目《格式定义》。")

    ctx.H("页面与页码", "sectionTitle1")
    ctx.P("正文采用A4纸张，四边页边距均为3.0厘米，页脚距边界1.5厘米。"
          "摘要至符号说明使用大写罗马数字编页，正文从第1页起使用阿拉伯数字。")

    ctx.H("标题与层级", "sectionTitle1")
    ctx.P("章标题采用三号黑体居中排列。一级节标题采用四号黑体，二级节标题为13磅黑体，三级节标题为小四黑体。"
          "章和节的编号随标题样式生成。参考文献和致谢等部分采用无编号章标题。")
    ctx.H("二级节标题示例", "sectionTitle2")
    ctx.P("二级节标题的编号由章号、一级节序号和本级序号组成。")
    ctx.H("三级节标题示例", "sectionTitle3")
    ctx.P("三级节标题在二级节标题编号后增加本级序号。")

    ctx.H("正文", "sectionTitle1")
    ctx.P("正文采用小四号宋体，西文和数字采用Times New Roman。段落两端对齐，首行缩进2字符，行距最小值20磅。")

    ctx.H("列表", "sectionTitle1")
    ctx.P("有序列表用于呈现具有确定顺序的条目。以下三项依次采用一至三级编号样式。")
    ctx.P("界定研究问题", "listNumber1")
    ctx.P("确定数据来源", "listNumber2")
    ctx.P("核验样本完整性", "listNumber3")
    ctx.P("无序列表同样按层级缩进。以下示例采用一至三级项目符号样式。")
    ctx.P("图表与公式", "listBullet1")
    ctx.P("题注与编号", "listBullet2")
    ctx.P("正文中的交叉引用", "listBullet3")

    ctx.H("插图", "sectionTitle1")
    ctx.P("图片居中，图注置于图下。图序与图题之间用一个全角空格分隔。")
    p = ctx.P(style="image")
    ctx.add_placeholder_shape(p, ctx.Cm(10), ctx.Cm(4.7))
    ctx.add_caption("figureCaption", "图", "1", "图", "示例插图")
    p = ctx.P()
    ctx.ref_caption(p, "图", "图", offset=-1)
    p.add_run("展示单幅插图及其题注。分图置于同一组图内，各分图在下方列出字母编号和题名，组末列出总图注。")
    subfigures = (("THU_SUBFIG_A", "a", "处理前的原始数据"),
                  ("THU_SUBFIG_B", "b", "处理后的结果"))
    p = ctx.P()
    for (bookmark, letter, title), punctuation in zip(subfigures, ("，", "。")):
        ctx.ref_caption(p, "图", "图")
        ctx.add_ref_field(p, bookmark, letter, "(")
        p.add_run(")示意" + title + punctuation)
    for bookmark, letter, title in subfigures:
        p = ctx.P(style="image")
        ctx.add_placeholder_shape(p, ctx.Cm(5), ctx.Cm(3), "分图示意")
        p = ctx.P("(", "subfigureCaption")
        bookmark_id = ctx.add_bookmark_start(p, bookmark)
        ctx.add_field(p, r"SEQ 子图 \* alphabetic", letter)
        ctx.add_bookmark_end(p, bookmark_id)
        p.add_run(")　" + title)
    ctx.add_caption("figureCaption", "图", "1", "图", "分图示例")
    ctx.P("注：占位框用于示意分图位置。", "figureTableNote")

    ctx.H("表格", "sectionTitle1")
    p = ctx.P("表格采用三线表，表注置于表上。表格上下边线为1.5磅，表头下边线为1磅。"
              "单元格文字采用11磅，段前、段后各3磅。")
    ctx.ref_caption(p, "表", "表")
    p.add_run("列出正文的几项排版参数。")
    _table(ctx, "三线表示例", ("项目", "取值", "说明"),
           (("页边距", "3.0厘米", "正文四边取值相同"),
            ("正文字号", "12磅", "中文宋体，西文Times New Roman"),
            ("正文行距", "最小值20磅", "段前、段后均为0")),
           widths=(3.2, 4.2, 7.6))
    ctx.P("注：表中取值适用于正文。", "figureTableNote")

    ctx.H("公式", "sectionTitle1")
    p = ctx.P("短公式可与正文同排，如质能关系")
    ctx.add_formula_omml(p, block=False)
    p.add_run("。")
    p = ctx.P("变量")
    ctx.add_inline_formula(p, "x")
    p.add_run("、")
    ctx.add_inline_formula(p, "y")
    p.add_run("与函数")
    ctx.add_inline_formula(p, "f(x)")
    p.add_run("采用行内公式格式。")
    ctx.P("行间公式居中排列，段前、段后各6磅，最小行距20磅。编号在公式右端，以圆括号标示。")
    p = ctx.P(style="formula")
    ctx.add_formula_omml(p, block=True)
    p = ctx.P()
    ctx.ref_formula(p, prefix_text="式(")
    p.add_run(")给出质能关系的行间表示。")

    ctx.H("脚注", "sectionTitle1")
    ctx.P("脚注用于补充说明【脚注示例】。正文中的脚注标记为上标带圈数字，按页重新编号。"
          "脚注文本采用小五号、单倍行距，悬挂缩进1.5字符。")

    ctx.H("参考文献与引文", "sectionTitle1")
    p = ctx.P("顺序编码制以方括号上标标注文献序号。可单独引用写作规范")
    ctx.add_citation(p, 1)
    p.add_run("或排版示例")
    ctx.add_citation(p, 3)
    p.add_run("，也可将两项来源合并标注")
    ctx.add_citation(p, 1, 3)
    p.add_run("。文献表条目采用五号，固定行距16磅，悬挂缩进1厘米。")
    ctx.P("著者出版年制按作者和年份编排文献表。"
          "以下两例的正文引文为作者甲和作者乙（2024）及AUTHOR（2023）。")
    ctx.P("作者甲, 作者乙. 期刊论文题名[J]. 期刊名, 2024, 1(1): 1-8.", "referenceEntryAuthorYear")
    ctx.P("AUTHOR A. Book title[M]. City: Publisher, 2023.", "referenceEntryAuthorYear")

    ctx.H("代码与强调", "sectionTitle1")
    p = ctx.P("短文件名可使用行内代码样式，如")
    ctx.set_char_style(p.add_run("sample.csv"), "inlineCode")
    p.add_run("。较长的程序片段使用行间代码样式并保留缩进。")
    ctx.P("for record in records:", "blockCode")
    ctx.P("    validate(record)", "blockCode")
    p = ctx.P("正文可按强调程度选用")
    ctx.set_char_style(p.add_run("不明显强调"), "subtleEmphasis")
    p.add_run("、")
    ctx.set_char_style(p.add_run("强调"), "emphasis")
    p.add_run("或")
    ctx.set_char_style(p.add_run("明显强调"), "strong")
    p.add_run("样式。")

    ctx.H("前置与后置部分", "sectionTitle1")
    ctx.P("中英文摘要分别成页。摘要后列出关键词，关键词之间用分号分隔。"
          "目录收录章标题、一级和二级节标题，插图清单和附表清单分别列出正文图注和表注。"
          "附录中的图、表和公式使用附录字母与序号组合编号。")

    _style_index(ctx)


def _style_index(ctx: ModuleType) -> None:
    ctx.H("样式索引", "sectionTitle1")
    p = ctx.P()
    ctx.ref_caption(p, "表", "表")
    p.add_run("至")
    ctx.ref_caption(p, "表", "表", offset=2)
    p.add_run("列出各类内容对应的样式。")
    header = ("样式", "用途")
    widths = (6.0, 9.0)
    _table(ctx, "章节与结构样式", header,
           (("章标题（有编号）", "正文各章的自动编号标题"),
            ("章标题（无编号）", "参考文献、致谢等章级标题"),
            ("章标题（无编号、不进目录）", "名单页等无需收录的章级标题"),
            ("一级节标题", "正文一级节标题，如1.1"),
            ("二级节标题", "正文二级节标题，如1.1.1"),
            ("三级节标题", "正文三级节标题，如1.1.1.1"),
            ("章内无编号小节", "章内小节标题"),
            ("名单页小节标题", "名单页分组标题"),
            ("附录标题", "附录各章标题"),
            ("附录一级节标题", "附录一级节标题，如A.1"),
            ("附录二级节标题", "附录二级节标题，如A.1.1"),
            ("附录三级节标题", "附录三级节标题，如A.1.1.1")),
           widths)
    _table(ctx, "文档内容样式", header,
           (("正文", "普通段落"),
            ("表格文字", "表格单元格"),
            ("图表附注", "图或表下方的附注"),
            ("图片", "插图所在段落"),
            ("图注", "正文插图下方的题注"),
            ("子图注", "分图下方的题注"),
            ("表注", "正文表格上方的题注"),
            ("附录图注", "附录插图下方的题注"),
            ("附录表注", "附录表格上方的题注"),
            ("公式", "行间公式"),
            ("行内公式", "段内公式"),
            ("脚注文本", "脚注区文字"),
            ("页眉", "页眉文字"),
            ("页脚", "页脚和页码"),
            ("目录一级条目", "目录中的章级条目"),
            ("目录二级条目", "目录中的一级节条目"),
            ("目录三级条目", "目录中的二级节条目"),
            ("符号说明条目", "符号和缩略语说明"),
            ("参考文献条目", "顺序编码制文献表"),
            ("参考文献条目（著者出版年）", "著者出版年制文献表"),
            ("文献引用", "顺序编码制的上标引文"),
            ("不明显强调", "轻度强调"),
            ("强调", "一般强调"),
            ("明显强调", "重点强调"),
            ("行内代码", "文件名、命令和简短代码"),
            ("行间代码", "独立代码段"),
            ("列表编号1", "一级编号列表：1."),
            ("列表编号2", "二级编号列表：(1)"),
            ("列表编号3", "三级编号列表：a."),
            ("列表项目符号1", "一级项目符号列表：•"),
            ("列表项目符号2", "二级项目符号列表：◦"),
            ("列表项目符号3", "三级项目符号列表：*"),
            ("授权说明标题", "授权页标题"),
            ("授权说明文字", "授权页正文")),
           widths)
    _table(ctx, "封面样式", header,
           (("封面论文题目", "中文封面题目"),
            ("封面学位类别", "封面学位申请信息"),
            ("封面作者信息", "中文封面作者与导师信息"),
            ("封面密级", "封面密级信息"),
            ("封面日期", "中文封面成文日期"),
            ("英文封面题目", "英文封面题目"),
            ("英文封面信息", "英文封面学位信息"),
            ("英文封面导师信息", "英文封面导师行")),
           widths)
