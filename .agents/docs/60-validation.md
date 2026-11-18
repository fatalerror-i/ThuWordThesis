# 验证清单

每次重建后按下表检查产物。构建的 `TestSmoke` 检查模板能否打开及基本结构。公式、题注、引用等按钮的行为须通过 Word 回归验证。数字以现行产物为准：下文的节数、表格数与内容控件数是当前基线，改动示例正文会变。

## 1. 规则与构建

```powershell
python src/validate_rules.py --all
python src/build_dotm.py --variant graduate
python src/build_dotm.py --variant undergraduate
python -m unittest discover -s tests -v
```

期望：

- `OK: rules structure valid [graduate]` 与 `OK: rules structure valid [undergraduate]`；
- 两次构建 `ok=true`、冒烟 `OK|...`，在 `output/` 各生成一份 DOTM、DOCX 和 PDF；
- 两个变体的静态检查与 Word 回归均通过，共用断言按研究生、本科各跑一次。

研究生版为21节、17个内容控件、12个表格和11组功能区。冒烟结果含 `OK|cc=16|sections=21`，另有 1 个封面页眉控件。

本科版为17节、7个正文内容控件（含封面表格首行密级）、8个表格和11组功能区。冒烟结果含 `OK|cc=7|sections=17`，所有页眉为空，封面段落使用“页眉”样式。封面使用校徽与校名的组合图。

Word 的 `Document.ContentControls` 只统计正文，页眉控件须从 `Sections(1).Headers(1).Range.ContentControls` 单独统计。

新增或修改示例后核对样式、编号和页面显示。第4章的行间公式采用“公式”样式、居中排列，编号为 `(4.1)`。

## 2. PDF 版面

用 `output/template_preview-graduate.pdf`（研究生）与 `output/template_preview-undergraduate.pdf`（本科）检查：

- 第 1 页中文封面：密级在独立页眉左侧，公开论文默认留空。题目首字符从版心顶部（距页顶 6 cm）起排；作者信息首段约 442.5 pt，日期约 636.75 pt。位置用 Word Range 检查，不能将 PDF 字形边界当作段落起点。
- 第 2 页英文封面：题目 top≈182.1、statement≈307.2、导师≈590.7、日期≈679.3；无表格框线。
- 研究生版第 3 页为名单，第 4 页为授权说明，两页均无页眉页码。名单按无框线表格对齐，授权说明标题无章号。
- 研究生版自摘要页起采用0.75磅页眉下框线（y≈76.9）。
- 目录收录章标题、一级及二级节标题。研究生目录收录自身，目录首尾页页眉均为“目录”。核对四章、附录和图表清单的页码。
- 核对各章和附录标题的位置，正文从阿拉伯数字1开始编页，后续各章及附录延续正文页码。
- 表4.1至表4.4与表A.1、表A.2紧接前文排版。表格可跨页，单行保持完整，表题注与首行同页。
- 参考文献 `[1]` 顶格，正文首行缩进 2 字符。
- 列表样式示例：一级编号顶格，二级右移 2 字符，三级右移 4 字符；续行（含标记后的正文）较编号再缩进 2 字符，标记不压正文。

本科预览的差异：所有页眉为空，密级在版心顶部独立表格行；目录从正文开始（目录自身及其余前置部分均不收录）；封面页边距为上 3.8 / 下 3.2 / 左 3 / 右 3 厘米、装订线 0.2 厘米只作用于封面。

本科封面用 Word Range 检查：表格宽与版心一致（14.8 cm），顶部独立密级行高 18.2 pt；logo 顶部约 146.7 pt（版心顶部以下 39.2 pt），题目首段约 320.7 pt、作者信息约 473.2 pt、日期约 680.5 pt。标题排一行或两行、密级切换时，作者和日期位置保持不变，封面应仍为一页。

## 3. 图标校验（本机）

```powershell
python references/icons/check_icons.py
```

用本机 Word 的 `CommandBars.GetLabelMso` 验证 `src/dotm/icons.yaml` 与 `src/dotm/customUI14-graduate.xml` 中配置的 `imageMso` ID 是否有效。该脚本只读校验、不导出图片；候选池预览图库的导出脚本同在 `references/icons/`，产物不随仓库分发。

## 4. XML/结构抽查

- 所有章级标题前一段含 `w:sectPr`（`nextPage`），正文无 `w:br type="page"`，章级样式无 `w:pageBreakBefore`。
- 封面第 1–2 页各 1 个表格，`tblBorders` 六边为 `nil`。
- 名单页的表格是无框线列表（`tblBorders` 六边 `nil`，单元格文字小四居中，用直接格式 12pt）；授权说明页**不用表格**，只有 `授权说明文字` 样式的段落。
- 所有内容控件 tag 唯一，无重复。
- `图注`/`表注` 的 `basedOn` 为 `Caption`；`子图注` 的 `basedOn` 为 `图注`；`图表附注` 无题注父类。
- 六个功能区列表样式的 `w:ind` 含 `w:hangingChars="200"`、`w:hanging="480"`；二级另含 `w:leftChars="200"`/`w:left="480"`，三级另含 `w:leftChars="400"`/`w:left="960"`，不含 `firstLine`。
- 页眉 `STYLEREF` 与章标题之间有一个全角空格；`document.xml`/`styles.xml` 无 `<w:vanish`。
- 页眉边框 XML 只在 Header 样式/正文页眉部件出现。
- 两版 DOCX 和 DOTM 的图片样式均有 `w:keepNext`、`w:widowControl`。跨页效果由稿件内容与 Word 排版决定。
- 两版目录标题共用“章标题（无编号）”。只有本科目录标题段有 `w:outlineLvl=9` 的直接格式。
- 附录说明段无“错误!未定义书签”。`document.xml` 无 CR/退格控制字符。
- 符号表的E、m、c是带“行内公式”样式的OMML对象。附录三个一级节依次展示文档部件、内容检查清单、格式检查清单，后两节各有一张表，表A.1和表A.2的引用指向对应表题。

## 5. Word 内人工确认

- 全选 F9 两次后，目录、页眉、页码、题注仍正确。
- 单击内容控件可输入；Tab 可切换；占位提示不打印。
- 另存/重新打开后，样式继承、页眉线、列表缩进不丢失。
- 另存 PDF 后再检查页数、目录页眉与目录页码。打印设备上的单双面、右页起排和装订空白页按 [打印与装订](../../docs/01-快速入门.md#10-打印与装订) 检查。
- 若 Word 测试弹出 VBA 编译对话框，记录报错模块和行，关闭对话框及测试用编辑器，再改源码、重建并复测。参见 [Word COM](70-pitfalls.md#10-word-com)。

## 6. 变更后回归

任何样式、规则或路径改动后，重新检查页数、节数、表格数、内容控件数、封面位置、目录页码、页眉线与列表缩进。其中任一项偏离基线，通常意味着 postprocess 或域更新被破坏。

## 7. 自动化测试

```powershell
python -m unittest discover -s tests -v
```

只跑某一层：

```powershell
python -m unittest tests.test_build -v          # 静态格式/样式契约，不启动 Word
python -m unittest tests.test_vba_units -v      # 宏单元 + 导出 + 本科 Word 有效值，需要 Word
```

`tests/support.py` 按变体找产物，每个 Word 用例操作临时副本。现有 `ensure_dotm()` 在 DOTM 缺失或早于源码时会调用构建，采用默认输出路径。因此应先重建两版交付产物，再跑 Word 测试，避免测试准备阶段隐式覆盖旧模板。静态测试读取既有的 DOCX、DOTM 和 PDF，模板内容改动后需一起重建。

| 测试文件 | 验证范围 |
| --- | --- |
| `test_build.py` | 规则和包结构、样式、编号、分节、目录和清单、交付物元数据。共用格式由 `_SharedStaticMixin` 对两版分别验证。 |
| `test_official_content.py` | 学术／专业学位英文名与层次、学校固定文本及续表入口停用。 |
| `test_vba_units.py` | Word 中的宏行为、另存重开和提交版导出。研究生专属封面和共用宏主要在 `GraduateMacroTest`，本科有效格式在 `UndergraduateWordTest`，两者共用 `_AppendixSequenceChecks` 与 `_CoverSecurityChecks`。 |

改动宏时重点重跑以下用例：

| 用例 | 预期结果 |
| --- | --- |
| `test_style_presentation_after_reopening_new_document` | 两版新建、另存、重开后，标题、目录、列表和明显强调的 15 个中文替换名均可查找，且仍为内置样式。 |
| `test_security_header_does_not_move_cover` | 两版密级写入、清空、重建时封面位置与页数不变；标题一/两行时作者与日期固定，研究生标题从版心顶部起排。 |
| `test_security_survives_export` | 两版 DOCX/PDF 导出保留密级，且后续页面不含密级。 |
| `test_formula_edit_preserves_math_and_existing_ref_target` | 在原公式前插入 `z=f(1)+2` 并编号，保留括号。更新和重新编号后，旧引用仍指向原公式并显示新编号。 |
| `test_native_heading_styles_supply_separator_and_appendix_references` | 直接套用章、节及附录各级样式均自动留出全角间隔，附录小节可引用。 |
| `test_document_and_style_hyphenation` | DOCX和DOTM均保存全文断字默认值及代码、公式、参考文献样式例外。 |
| `test_hyphenation_defaults_and_style_exceptions_survive_reopening` | 从模板新建、保存重开后断字设置保持；正文含URL时仍沿用正文样式。 |
| `test_western_words_move_to_next_line_without_midword_wrapping` | 在 Word 窄行排版中核对临界单词整体移到下一行，防止误用 `w:wordWrap=0`。 |
| `test_export_preserves_hyphenation` | 提交版保留文档断字设置和参考文献例外。 |
| `test_long_url_wraps_without_changing_text` | 两类文献条目中的长URL可换行，文字完整并保持在版心内。 |
| `test_check_table_header_stays_with_first_data_row` | 两版检查表的表头与首条检查项同页。 |
| `test_cover_type_transitions_preserve_names_and_english_fields` | 学术、有领域、无领域类型切换后，中文控件及英文信息正确保留与恢复。 |
| `test_numbered_heading_preserves_bookmark` | 给带书签的现有段落套编号标题，保留书签和标题文字。 |
| `test_second_appendix_restarts_formula_figure_and_table_numbers` | 两版新增附录 B，其首条公式为 `(B.1)`，图、表分别按 1、2 排序。 |

测试数以终端报告为准，最近一次验证见 [当前交接](90-handoff.md)。

改格式时只需改共用断言：`GraduateStaticTest` 与 `UndergraduateStaticTest` 会同时验证两个变体，不用两处重复写。基线数字放在 `tests/support.py` 的 `BASE_COUNTS`，改示例正文后同步更新。

约定：

- 每个用例复制一份 dotm 到 `TemporaryDirectory`，结束关闭并删除。
- VBA 测试入口统一返回字符串：`ok|...` 或 `OK|...`，失败返回 `FAIL|错误号|描述`。
- 纯格式断言用正则接受不同章号。验证重排、复位和引用指向时，断言该场景的具体编号与目标。
- 导出目标写入用例自己的临时目录（dotm 副本同目录），用例结束随临时目录删除；不会写入 `output/` 交付目录。
