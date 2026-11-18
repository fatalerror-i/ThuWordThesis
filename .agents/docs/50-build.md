# 生成器与构建管线

本册讲「规则怎么落到 Word 里」。规则怎么写见 `20-rules.md`，产物长什么样见 `30-docx-structure.md`。

## 构建链

```text
build_dotm.py --variant graduate | undergraduate
  └─ finalize_docx.py 调用 build_docx.py 生成 DOCX，更新域并后处理
       ├─ 保留 DOCX 示例
       ├─ 用 Word 导出 PDF 预览
       └─ 用 Word 另存为 DOTM，注入宏、窗体和功能区，运行冒烟测试
            └─ 将三份成品写入 output/
```

`build_dotm.py` 在临时目录构建，整个流程共用一份完成更新域和后处理的 DOCX。DOTM 冒烟测试通过后输出三份文件，结束时清理临时目录。

```powershell
python src/validate_rules.py --all
python src/build_dotm.py --variant graduate
python src/build_dotm.py --variant undergraduate
python -m unittest discover -s tests -v
```

## 源码与目录

| 路径 | 作用 |
| --- | --- |
| `configs/variants.yaml` / `src/build/variants.py` | 双变体注册表（`--variant` 必选、无默认值），提供各变体规则、产物、Ribbon、ThisDocument 路径 |
| `configs/thesis-format-rules-graduate.yaml` / `configs/thesis-format-rules-undergraduate.yaml` | 研究生 / 本科页面、样式、编号、题注、封面等格式规则 |
| `configs/implementation-map.yaml` | 规则样式名到 Word 样式 ID、基于样式、回退样式的映射；继承只写在这里（两变体共享） |
| `src/build_docx.py` | 构建入口：读取当前变体规则，创建样式、节、页眉页脚，再按变体动态导入内容模块组装正文 |
| `src/build/content/graduate/` / `src/build/content/undergraduate/` | 按变体维护封面、摘要文字、前置页和专属后置章节。两份后置模块均调用共用的附录生成器。 |
| `src/build/content/examples.py` | 共用摘要布局辅助函数、符号说明和四章正文。摘要文字由各变体的 `front_matter.py` 提供。 |
| `src/build/content/appendix.py` | 共用附录检查清单、附录图表和公式示例。 |
| `src/build/` | assemble、variants、Word 阶段等构建支撑；学位配置在 `src/build/degree_catalog.py`（学位显示串、学科候选、英文名称与 VBA 学位函数生成） |
| `src/postprocess_docx.py` | Word 更新域后的补救：目录、图表目录、题注、列表、参考文献、样式编号、页眉页脚框线（两变体都清框线，仅研究生按规则加下框线）及样式分页控制 |
| `src/finalize_docx.py` | 生成 DOCX，用 Word 更新域并完成后处理，再导出 PDF。测试只重建 DOTM 时通过内部参数 `--for-dotm` 跳过 PDF |
| `src/build_dotm.py` | 复用 DOCX 生成 DOTM，写入宏、窗体和功能区，运行冒烟测试并输出三份成品 |
| `src/office_metadata.py` | 清除 Word 和 PDF 构建产物的作者、本机路径及构建环境元数据 |
| `src/dotm/vba/*.bas` | 9 个 VBA 标准模块（含构建时生成的 `modDegreeOptions`），两变体共用 |
| `src/dotm/vba/ThisDocument-graduate.txt` / `ThisDocument-undergraduate.txt` | 研究生文档级事件（内容控件退出时调用封面级联）/ 本科无事件最小文件，构建时按变体写入 |
| `src/dotm/forms/*` | 选择器和引用检查窗体。两者使用系统配色和统一按钮布局。选择器支持拖动调整大小，引用检查使用四个可滚动的选项卡。 |
| `src/dotm/customUI14-graduate.xml` / `customUI14-undergraduate.xml` | 两版“论文”选项卡均 11 组；本科封面组仅有密级菜单 |
| `tests/` | 静态格式/样式契约与 VBA 宏单元测试 |
| `references/icons/*.py` | 图标 imageMso ID 校验与本机预览图库导出脚本，不参与构建 |

所有脚本从自身位置推导 `ROOT`：

```python
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
```

图标工具脚本在 `references/icons/`，比仓库根深两层，`ROOT = Path(__file__).resolve().parents[2]`；读 `ROOT/src/dotm/...` 与 `ROOT/references/icons/...`，导出的图库也写回 `ROOT/references/icons/...`。

## build_docx.py

主要区块：

| 区块/函数 | 作用 |
| --- | --- |
| `STYLE_NAMES` | 规则语义键 → Word 显示名，键数与规则文件 `styles` 一致 |
| `BUILTIN_IDS`/`_ensure_builtin_paragraph` | 复用/确保内置样式：Normal、Heading 1–4、Caption、TOC1–4、FootnoteText、Header、Footer、List Number/Bullet 1–5 |
| `set_doc_defaults()` | 段落与字符默认值（`pageSetup.documentDefaults`），同时写 `w:rPrDefault` 与 `w:pPrDefault` |
| `apply_character()` / `apply_paragraph()` | 从 YAML 写样式 |
| `build_numbering()` | outline/appendix/references 三套多级列表；六个列表样式的编号层写悬挂缩进 |
| `set_style_numbering()` / `disable_style_numbering()` | 样式挂编号/抑制编号 |
| `setup_section()` | 页面尺寸、边距、页眉页脚距离、装订线、页码格式 |
| `set_header()` / `set_footer()` / `set_pgnum()` | 页眉域、页脚页码、页码续接 |
| `add_field()` | TOC/STYLEREF/SEQ/PAGE/REF 复杂域；`result` 参数是域的缓存结果 |
| `add_bookmark_start()` / `add_bookmark_end()` | 在段落内开合书签，返回/接收书签 id |
| `_caption_scan()` / `peek_caption()` | 扫描已写题注，推算下一题的书签名与编号（`THU_FIG_0001`、`1.1`）|
| `add_content_control()` | `w:sdt` 内容控件，支持 text/richText/dropdownList/comboBox、tag/alias/占位；下拉与组合框用 `items` 写 `w:listItem` |
| `_cover_table_properties()` / `add_cover_layout_table()` | 封面单列无框线固定行高表格 |
| `format_cover_cell_paragraph()` / `cover_cell_paragraph()` | 封面单元格段落、制表位 |
| `add_cover_security_header()` | 研究生中文封面独立页眉，写入默认空白的密级控件 |
| `add_cover_info_paragraph()` | 封面作者信息行 |
| `add_caption()` | 题注：STYLEREF + SEQ + 全角空格，并把编号文本包进书签 |
| `ref_caption()` / `add_ref_field()` | 正文写 `REF` 域引用题注编号；`offset=-1` 引用刚建好的上一条题注 |
| `ref_formula()` | 引用公式编号 `THU_EQ_0001`，书签由 `finalize_docx.py` 补建 |
| `add_formula_omml()` | OMML 公式示例 |
| `add_table_style()` / `three_line_table()` | 三线表样式/示例 |
| `add_seq_reset()` | 在附录示例导语段末尾写 `SEQ <标识符> \r 0 \h`，供预览构建时初始化编号。运行时按附录重排的逻辑见 [VBA 实现](80-dotm.md#附录编号)。 |
| 正文说明与示例 | 第1章引言，第2章按功能组介绍DOTM宏模板，第3章介绍DOCX的Word手动操作，第4章展示格式与样式。行内和行间公式示例集中在第4章“公式”小节。 |
| 样式索引 | 表4.2列章节结构，表4.3“文档内容样式”包括引用与列表，表4.4仅列封面样式；显示中文替换名，列表只列1–3级。 |
| 符号表与附录 | 符号按语义组织，物理量用行内OMML。共用附录分为A.1文档部件示例、A.2内容检查清单、A.3格式检查清单，仅后两节各含一张表。两张清单同时演示附录表格及其交叉引用。 |

封面表格关键参数：

- 中文：宽 368.5pt，5 行，行高 `(130,142.5,156,34.25,36)`（题目与学位申请句同一行）；密级在独立页眉左侧，页眉距页顶 2.8 cm，不占版心。
- 英文：宽 391pt，8 行，行高 `(110,8.2,180,56,46.8,60,25.6,50)`。
- 本科：宽 14.8cm，与版心一致，10 行，行高 `(18.2,101,18,50,26,136,16.5,156,47.3,40)`；直接读取本科 `cover.chinese.layout`，宽度单位 cm、行高单位 pt。封面空白页眉使用内置“页眉”样式。首行密级为四号黑体，logo 段前 21pt 包含在第二行高度内；作者与日期分别约距页顶 473.2pt、680.5pt。
- 表格无边框、固定布局、单元格边距 0、行高 EXACTLY、`cantSplit`。
- 封面内容控件：`学位类别` 为 `dropDownList`（`DEGREE_DISPLAY_OPTIONS`，按已配置学位层次展开），`学科（专业学位改为专业领域）` 为 `comboBox`（`DEFAULT_DISCIPLINE_OPTIONS`，默认真实内容可输入）；`成文日期`、`英文提交声明`、`英文degree行`、`英文学位名称`、`英文成文日期` 构建时写真实内容，不用 `showingPlcHdr`。学位/日期级联不在构建时写死，由 ThisDocument 事件在用户操作后调用 `modCover`。

生成器与后处理均调用 `build.hyphenation.apply_document_text_flow()` 恢复文档断字默认值和样式例外；只操作 `settings.xml` 与样式定义，不检查正文内容。

## 规则到 Word 的映射

- 样式规则先在规则 YAML 里定义中文字体、西文字体、字号、段距、行距、缩进、编号、大纲级别等。
- `configs/implementation-map.yaml` 把规则键映射到 Word 样式：`name` 与 `fallback`（继承）。规则文件本身不写继承。
- 映射里不写 `styleId`。Word 打开并保存一次之后会按内部规则重排自定义样式的 `styleId`（`footnote text` 这类内置样式也会被换成 `afff0` 之类），写死存不下来。代码、VBA 与测试一律按**样式名**查找样式；`build_docx.py` 建内置样式时用 `add_style(..., builtin=True)` 让 Word 认。

  验证样式是否落地时不要比对 `styleId`，只比 `w:name`。`postprocess_docx.py` 里按 `style_id` 匹配段落是允许的：那里取的是当前文档对象上的实时值，不是常量。
- `src/build_docx.py` 读取两份 YAML，创建或改写样式、编号、节、页眉页脚和内容控件；`src/postprocess_docx.py` 处理 Word 更新域后的目录、图表目录、样式编号与框线。页边距、页眉/页脚距离与装订线都直接读 `pageSetup`（封面读 `pageSetup.sections.*.pageMargins`），生成脚本不再写死数值。
- 段落默认值写进 `w:rPrDefault` 与 `w:pPrDefault`（`pageSetup.documentDefaults` → `set_doc_defaults()` + `_set_ppr_default()` + `apply_paragraph_to_ppr()`）。`apply_paragraph()` 直接改样式层的 `w:pPr`，`apply_paragraph_to_ppr()` 改裸 pPr（docDefaults 用），两者的行距映射一致：`minimum → lineRule="atLeast"`，`fixed → "exact"`。
- 样式分三类循环处理：`CHARACTER_STYLE_KEYS`（`styleType: character`）走 `apply_character()`，`TABLE_STYLE_KEYS`（`styleType: table`）由 `add_table_style()` 单独建 `w:type="table"`，其余走段落循环。新增表格样式时必须靠 `TABLE_STYLE_KEYS` 过滤，否则会被当成段落样式重建。
- `apply_character()` 支持 `color`（只写 `w:val`，不写 `themeColor`，避免点击后变色）与 `underline`（`none` 显式写 `w:val="none"`）；`apply_paragraph()` 的 `tabs: []` 表示「先清空继承来的制表位，再加规则里的」，空列表即清空。
- 注意 Word 的冗余剥离：`SaveAs2` 会把「与继承值相同」的属性当冗余删掉（例如 `docDefaults` 已声明的 `w:before="0"`）。有效值不变，但**验证时必须用 Word COM 读 `Paragraph.Format` / `Style` 的有效值，不能只看样式自身 XML**。
- Word 保存时可能省略样式里的分页控制。DOCX 后处理显式恢复图片样式的 `widowControl`。DOTM 包处理调用 `restore_style_pagination()`，按规则在 `word/styles.xml` 中恢复 `keepWithNext` 对应的 `w:keepNext` 和 `widowControl` 对应的 `w:widowControl`。图片与题注的跨页效果仍随实际稿件变化。
- 规则校验：`python src/validate_rules.py` 检查键顺序、样式索引、继承键、单位等；构建前必须通过。

## postprocess_docx.py

| 函数/常量 | 作用 |
| --- | --- |
| `set_fonts()` / `set_indent()` | 统一目录/题注/页眉页脚字体与缩进 |
| `clear_direct_run_format()` | 清 Word 更新域后留下的直接字体格式 |
| `_clear_pbdr()` / `_nil_pbdr()` | 清段落边框；页脚/封面页眉用 nil |
| `_header_pbdr()` | 正文页眉底边 single 0.75pt、space=1，其余 nil |
| `BASE_MAP` | 恢复样式继承链 |
| `ensure_no_numbering()` | 给无编号章标题/授权说明标题/名单页小节标题写 `numId=0` |
| 列表样式恢复 | 按规则给六个列表样式写 `w:hangingChars=200`、`w:hanging=480`；二级加 `w:leftChars=200`/`w:left=480`，三级加 `w:leftChars=400`/`w:left=960`；删除 `firstLine/firstLineChars` |
| 参考文献 | 样式与每条段落都写段前 3pt、段后 0pt；Word 保存会删同值样式属性，所以双重写入 |

## finalize_docx.py

1. `build_docx.py`；
2. `word_update_fields()`：DispatchEx Word，隐藏窗口；两轮更新 SEQ/STYLEREF、TablesOfContents、TablesOfFigures、页眉页脚域；
   给公式编号补 `THU_EQ_0001` 书签（公式编号在这一步才由 `#()` 生成，构建期算不出），更新 `REF` 域时先用 `Bookmarks.Exists()` 判断书签是否存在；
   公式的章节号域按段落位置切换：正文用 `STYLEREF 1 \s`，附录用 `STYLEREF "附录标题" \s`，遍历时以样式名 `附录标题` 置位、`章标题（无编号）` 复位；
3. `postprocess_docx.py`；
4. `word_update_fields()`；
5. `postprocess_docx.py`；
6. 清理 DOCX 文档属性。`word_export_pdf()` 以只读方式打开母版，调用 `ExportAsFixedFormat` 导出预览并清除 PDF 的文档信息和 XMP。

**两轮不可省**：后处理会改编号/样式，第一次目录缓存可能仍是旧章号。

## build_dotm.py

1. 在临时目录调用 `src/build/assemble.py` 生成 DOCX 及配套 PDF；
2. Word COM 打开 docx 并 `SaveAs2(path, FileFormat=15)` 生成 `.dotm`；
3. `VBProject.VBComponents.Import` 依次导入 `src/dotm/vba/*.bas`（UTF-8 转 GBK）与构建生成的 `modDegreeOptions.bas`。生成模块还含 `THU_UNDERGRADUATE` 及供旧稿只读识别使用的续表前后缀；
4. 按变体写入 `ThisDocument` 事件源：研究生用 `src/dotm/vba/ThisDocument-graduate.txt`（`Document_ContentControlOnExit` 级联入口），本科用 `src/dotm/vba/ThisDocument-undergraduate.txt`（无事件的最小文件）；`CountOfLines>0` 时先 `DeleteLines`，再 `AddFromString`，构建日志记录 `thisdocument_lines`；
5. 用 `src/dotm/forms/forms.json` 在 VBE 内创建选择器 `frmPicker` 和检查结果窗体 `frmReferenceCheck`，再注入各自的代码。
6. `doc.Save()` 并关闭 Word，解包 dotm，按变体写入功能区 XML（研究生 `src/dotm/customUI14-graduate.xml` 11 组、本科 `src/dotm/customUI14-undergraduate.xml` 11 组）与关系；
7. 清除脚注段落直接格式，补回 Word 保存时移除的样式属性，清理文档属性、VBA 库路径和编译缓存。
8. 以只读方式重新打开最终 DOTM，运行 `TestSmoke`，检查 VBA 库和窗体控件，记录 `vb_components`、`content_controls`、`sections`。
9. 将三份成品写入 `output/`，在终端打印构建摘要。

`--variant` 必选，选项来自 `configs/variants.yaml`。构建脚本通过命令参数把变体和输出路径传给子进程。测试内部使用 `--dotm-only`、`--keep-existing`、`--skip-test` 等选项。

## 内容控件 tag 契约

默认正文16个（专业学位无领域时移除中文学科控件）：`论文题目`、`学位类别`、`培养单位`、`学科（专业学位改为专业领域）`、`研究生（专业学位改为申请人）`、`指导教师`、`成文日期`、`英文论文题目`、`英文提交声明`、`英文degree行`、`英文学位名称`、`英文学科`、`英文专业领域`、`英文作者`、`指导教师（英文）`、`英文成文日期`。

可选（3）：`密级`、`副指导教师（可选）`、`联合指导教师（可选）`。

`学位类别` 为下拉内容控件，`学科（专业学位改为专业领域）` 为可输入组合框；其余为文本或 richText。`成文日期`、`英文提交声明`、`英文degree行`、`英文学位名称`、`英文成文日期` 构建时写真实内容，不带占位符。研究生 `密级` 控件在首节主页眉，本科在封面表格首行，默认空白。两版均用 `modCover.ToggleSecurity` 写入 `秘密★5年` 或清空，保留本科密级行。研究生默认总控件数为17（正文16、页眉1），本科为 7（均在正文）；删除密级后各少一个。

改 tag 必须同步说明页、`.agents/docs/80-dotm.md`、`src/dotm/vba/ThisDocument-graduate.txt` 和 VBA。

## 正文写作约定

示例文字按以下约定手工维护（具体论文正文不写入自动化测试）：

- Word 正文统一使用全角引号“”，引号字符按中文字体渲染（`w:rFonts w:hint="eastAsia"`）；不写 ASCII 直引号。
- 不在中文与西文之间手工加空格，Word 会按字体自动调整间距；`docs/` 与 `.agents/docs/` 中的 markdown 仍按中文排版习惯加空格。
- 中文标点用全角；数字、变量名与代码用半角，行内代码套“行内代码”样式。
- 正文和附录演示两版共用的写作流程、编号和样式，保留作者需要替换的内容提示。安装、更新模板和按钮细节写在用户文档。DOCX 与 DOTM 由同一套内容模块生成。

## 调试

- 生成失败先看 Python traceback；样式名 `KeyError` 多为旧字面量未改语义键。
- Word 卡住：保存工作后 `taskkill /F /IM WINWORD.EXE`，不要盲目重跑。
- PDF 未生成：检查 Word 是否弹窗/文件被占用；改用 ExportAsFixedFormat。
