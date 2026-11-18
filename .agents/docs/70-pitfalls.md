# Word 行为坑与已知问题

## 1. 内容控件段落会被 Word 改写

Word 更新域并保存后，会删除内容控件所在段落的直接 `w:jc`/`w:ind`，也可能删除 `w:richText` 类型标记。因此：

- 标题居中、行距必须放在样式层；
- `postprocess_docx.py` 的 `BASE_MAP` 负责恢复被清掉的 `basedOn`；
- 不要依赖段落直接格式做封面定位。

## 2. 制表位对齐的是制表符后面的整段文字

中文封面要让冒号居中：先 `\t`，再标签+冒号，再 `\t`，再值。右制表位要对齐“标签+冒号”的右边缘（页面 267.3pt，单元格内 189.65pt），不是标签右边缘；否则冒号中心会偏左。左制表位 195.65pt，值起点约 273.3pt。

## 3. 无编号标题必须显式 numId=0

`授权说明标题`、`名单页小节标题` 若只 `basedOn Heading 1/2`，会继承章编号，导致正文从“第 3 章”开始。用：

```xml
<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
```

并保留 `outlineLvl=9`。`ensure_no_numbering()` 已处理。

## 4. 分节后的页码

新节会继承上一节 `w:pgNumType w:start="1"`，导致页码重新从 1 开始。`set_pgnum(sec, "decimal")` 不传 start 才是续前节。第 2 章后的新节必须调用。

## 5. 章标题的段前距

`chapterTitle`、`chapterTitleUnnumbered`、`chapterTitleNoOutline`、`appendixTitle` 的 `pageBreakBefore` 都必须为 `false`；每个章级部分在标题前由 `chapter_section()` 插入“分节符（下一页）”。只依赖章标题样式的“段前分页”会失去节级页眉页脚和页码控制；分节符直接插入，不额外设置自带空段落的字体或段落格式。

## 6. 图表目录样式（table of figures）

插图/附表清单用 `TOC \h \z \c "图"`/`\c "表"`，Word 用**内置 `table of figures`（中文显示名“图表目录”）**实例化条目；`toc 4`/`TOC4` 在成品里 0 引用（规则键 `listOfFiguresEntry` 保留作兼容，条目格式由 `tableOfFigures` 承载）。`doc.styles.add_style("table of figures")` 必须 `builtin=True`，否则生成的名字对不上；`toc 4` 同样需要 `builtin=True` + `styleId="TOC4"`。构建阶段不要访问 `doc.styles["toc 1"]`：TOC1 是 Word 更新目录后才生成的。

Word 每次更新域会重新盖直接缩进与 1.15 倍行距，所以 `postprocess_docx.py` 必须每次域更新后按 `tableOfFigures` 规则重写该样式的字体、行距、缩进与制表位。

## 7. 页眉下框线

研究生规则的 `styles.header.paragraphFormat.borderBottom` 要求 `single 0.75pt spacing 1pt`，只能加在 `Header`/`页眉` 样式的页眉段落上；本科规则没有 `borderBottom`，页眉为“无”。两变体的页脚都无框线，页眉页脚都不设制表位（`tabs: []`）。实现上 `postprocess_docx.py` 现在两变体都先清框线，只在 `VARIANT.decorate_header and 规则写了 borderBottom` 时给研究生页眉加下框线。

## 8. 列表样式缩进被 postprocess 清零

postprocess 的样式清零循环必须跳过 `List Number*`/`List Bullet*`，并在末尾按规则恢复缩进：六个样式都写 `w:hangingChars="200" w:hanging="480"`，二级另写 `w:leftChars="200" w:left="480"`，三级另写 `w:leftChars="400" w:left="960"`。编号层只保留悬挂缩进，不写 `w:left`；`w:firstLine` 一律不写。

## 9. 封面表格被页眉距离推下

研究生密级使用独立页眉，header distance 为 2.8 cm，必须使用“封面密级”样式，段前后为 0，后续节断开链接。本科页眉始终为空，密级在封面表格首行，行高固定 18.2 pt；填写或清空时保留该行，其他部件位置不变。

Word 在空单元格重建内容控件后，可能将 `w:sdt` 包在 `w:tc` 外，而非段落内。Word 能正常保存、导出和继续编辑；检查导出文件时应按 `密级` tag 查找控件，不能依赖 python-docx 的 `cell(0, 0)`，该 API 会跳过被控件包裹的单元格。

## 10. Word COM

- 用 `DispatchEx`，不要复用用户已打开的 Word。
- 结束 `doc.Close(False)`、`word.Quit()`；异常时也退出。
- Word 卡住：保存工作后 `taskkill /F /IM WINWORD.EXE`，不要盲目重跑。
- `SaveAs(PDF, FileFormat=17)` 偶尔不生成文件，改用 `ExportAsFixedFormat`。
- 测试里不要对 dotm 副本调用 `Document.Save()` 当“保存后导出”的前置步骤：同一 Word 实例先打开/导出过别的文档后，再 Save 模板副本会阻塞在隐藏的模态保存（`faulthandler` 栈停在 `_Document.Save`，可挂起数分钟）。冻结/更新域这类逻辑改用 VBA 测试入口在内存里断言（如 `TestCitationFreeze`），导出产物直接解包断言 `word/document.xml`。

VBA 编译错误可能在隐藏的 Word COM 测试中弹出 VBE 对话框，调用会等到对话框关闭。先读取错误和模块，关闭测试用对话框并结束测试用调试会话，再改 `src/dotm/vba/` 并重建。VBA 标识符不区分大小写，同一过程里的常量名和变量名不能仅以大小写区分。构建冒烟通过后仍需跑实际宏用例。

## 11. python-docx 不能直接打开 .dotx

报 `ValueError: ... content type ...template.main+xml`。用 Word COM 另存为 `.docx`，或解压读 `word/document.xml`。

## 12. 路径约定

所有新增代码从 `ROOT` 派生，不要写死 `.work`、盘符路径或用户目录；改动后扫描旧路径关键字与根目录旧规则名。

## 13. 附录 SEQ 复位域与书签错误

附录复位域必须写成 raw string `SEQ 表 \r 0 \h`。若 `\r` 被解释成回车，字段变成 `SEQ 表 0 \h`，Word 把 `0` 当成书签名，附录说明段显示“错误!未定义书签”。

构建用 raw string 写 `SEQ \r 0 \h`；`modMain.RepairAppendixResetFields` 在 `AutoOpen`、`EnsureReady`、`UpdateAllFieldsNow` 中扫描并重写这类域，也可手动点“引用 → 更新域”。正常字段应含 `SEQ 表 \r 0 \h`，不是 `SEQ 表 0 \h`。

运行时跨附录编号不能只依靠 `\s 1`，它按 Word 的“标题 1”重置。模板的附录使用自定义标题样式。`RefreshAppendixSequences()` 为每个附录的首条图、表、公式编号写 `\r 1`。对 OMath 内的 `Field.Code.Text` 赋值时，Word 可能把 `*` 变成数学乘号 `∗`。公式域重写时省去默认的 `\* ARABIC`，普通图表题注仍显式指定。

## 14. Ribbon customUI 静默失效，整个“论文”选项卡消失

Word 对 customUI 是**整包加载**的：XML 中只要有一个控件结构不合法，Word 可能不弹错误，而是丢弃整个 customUI。表现是 VBA 工程正常、宏也能运行，但“论文”选项卡完全不出现。不要先怀疑宏被禁用。

### 已确认的两个原因

1. **`splitButton` 结构不合法**
   - 旧图片组写成：

   ```xml
   <splitButton id="btnSubfigureCaption" onAction="InsertSubfigureCaption">
       <button .../>
       <menu .../>
   </splitButton>
   ```

   - 当前 Word 不接受这种 `splitButton` 同时带 `button` 和 `menu` 的写法，导致整个 customUI 加载失败。
    - 当前实现：图片组只保留普通按钮 `btnSubfigureCaption`；不使用 `splitButton`，也不另设“子图编号”菜单。

2. **包内 XML 使用了 `ns0:` 前缀**
    - `src/build_dotm.py` 用 ElementTree 重写 `_rels/.rels`、`[Content_Types].xml` 时，若输出带 `ns0:` 前缀（`ns0:Relationship`、`ns0:Override`），Word 的 Ribbon 加载器不识别。
    - 处理：序列化前设置默认命名空间：

   ```python
   ET.register_namespace("", REL_NS)
   ET.register_namespace("", CT_NS)
   ```

   - 另有一个可观察现象：带 `ns0:` 的 dotm 用 Word 另存为 `.docm` 后，`[Content_Types].xml` 中的 customUI override 会直接消失，但 customUI part 和关系仍在。

3. **空元素写成成对标签**
   - 写成 `<button ...></button>`（即使两个标签之间只有空白）会被 Word 的 customUI 解析器拒绝，整包 customUI 被丢弃，表现为“论文”选项卡消失、宏仍能运行。
   - 修复：空元素必须自闭合，写成 `<button .../>`。
    - 行间公式菜单的子按钮必须写成自闭合标签；成对写法会让整个“论文”选项卡消失。用最小对照（只替换 customUI14-graduate.xml）可快速确认。

### 正确诊断方法

- **不要用 `GetLabelMso("THUThesisTab")` 或 `CommandBars("Ribbon").Controls.Count` 判断自定义选项卡是否加载**：前者对自定义 tab id 会报 `-2147352567`，后者不统计 customUI 的 tab。以 Word 可见窗口的截图/选项卡文字为准。
- 用同一个 dotm 包做对照实验：
  1. 只放一个最小 customUI（一个 tab、一个 group、一个 button）；若能显示，说明包关系、内容类型和 Word 环境正常。
  2. 在当前完整 customUI 中删掉可疑控件，再重建；若选项卡出现，即可锁定该控件。
  3. 把可疑控件替换为更简单的 `button`/`menu`，确认修复。
- 在 Word“文件 → 选项 → 高级 → 常规”勾选“显示加载项用户界面错误”，部分 customUI 错误会给出更具体的提示。
- 确认 dotm 从受信任位置打开，避免受保护视图干扰；但“宏已启用、VBA 能运行”不代表 Ribbon XML 一定合法。

### 回归检查

- `tests/test_build.py::test_ribbon_interaction_layout` 检查 customUI 中不得出现 `splitButton`；并检查所有 `<button>` 元素都是自闭合（无文本、无子元素）。

## 15. 本科图表目录（Table of Figures）的左缩进会自己长回来

两变体插图/附表清单统一用 `TOC \h \z \c "图"` / `\c "表"`（按题注标签收集，与手工"插入表目录→选题注标签"一致）：

- Word 更新 `\c` 域时会实例化内置 `table of figures` 样式，默认带 2 字符左缩进（`w:leftChars=200 w:left=200 w:hangingChars=200 w:hanging=200`），并给每条结果段落盖直接 `w:left=960 w:hanging=480`。表现就是图表清单整体右缩、用户报告的"不该出现的左边距"。
- 官方本科样本里 `table of figures` 是 `basedOn=TOC 1`，样式层与段落层都没有 `w:ind`；`postprocess_docx.py` 必须同时规范化 `toc 1`–`toc 4` 与这个内置样式。
- **每次更新域都会重新加回缩进**，所以修复必须放在 `postprocess_docx.py`（构建链在每轮 `word_update_fields()` 后都会重跑它），不能只改 `front_matter.py` 的域指令。
- **必须按 `styleId` 匹配结果段落，不要按 `p.style.name`**：中文 Word 把内置样式名本地化成"图表目录"，英文名匹配会漏掉条目段落。`d.styles` 里仍能用 `"table of figures"` 找到样式对象。
- Word 在 `SaveAs2` 成 dotm 时会剥掉该样式的 `w:basedOn`，因此别把 `basedOn=TOC 1` 写成断言；真正要保证的是样式层无 `w:ind`、每条结果段落无直接 `w:ind`。
- **同一内置样式还带段后 10 磅 + 1.15 倍行距**（指南 2.3.7 要求清单条目宋体小四、固定 20 磅、段前段后 0 磅）。这个默认值是 Word 内建 `docDefaults` 提供的，`styles.xml` 里并不显式写 `w:spacing`，所以只查 XML 会以为"没设"而漏掉。必须像 `toc 1`–`toc 4` 一样在 `postprocess_docx.py` 里显式写 `space_before/after=Pt(0)` + `line_spacing=Pt(20)` + `EXACTLY`，并清掉结果段落的直接 `w:spacing`。规则文件 `listOfFiguresEntry` 必须写全 `fontSize`/`lineSpacing`/`spaceBefore`/`spaceAfter`。规则是唯一事实来源，规则、实现与测试三者要覆盖同一组属性。
- 回归：`tests/test_build.py::test_both_variants_tof_unified`（两变体、含 `w:spacing` 断言）。每次域更新后 `postprocess_docx.py` 都会重写，缩进不会回潮。
- 改 Ribbon 后至少验证：`customUI14-graduate.xml` 可解析、`_rels/.rels` 关系存在、`[Content_Types].xml` override 存在且无 `ns0:` 前缀、Word 可见窗口出现“论文”选项卡。

## 16. wordprocessingShape 占位图形不在 InlineShapes 里

示例插图改由 `src/build_docx.py` 的 `add_placeholder_shape()` 插入 Word 原生图形（`wps:wsp`，XML 里写的是 `<wp:inline>`），不产生图片部件、也不需要成像库。但 Word 对象模型对它的归类与 XML 不一致：

- 打开产物后 `ActiveDocument.InlineShapes.Count = 0`、`Shapes.Count = 1`，图形 Type 为 1（矩形），`WrapFormat.Type = 7`。
- 锚点段落的 `Range.Text` 是 `"/"`（图形在纯文本里落成斜杠），`Range.InlineShapes.Count = 0`，但 **`Range.ShapeRange.Count = 1`**；普通段落两者都是 0。
- 用 COM 新建 `Shapes.AddShape(...)` 再 `ConvertToInlineShape()` 存盘，XML 仍是 `<wp:inline>` + `wps:wsp`，重开又回到 `InlineShapes = 0`。**`wps` 图形永远进不了 InlineShapes。**

所以判断"光标是否在图上"不能只看 `InlineShapes`。`src/dotm/vba/modCaption.bas` 用两个私有辅助函数统一处理：

- `Private Function HasFigureAt(ByVal p As Paragraph) As Boolean`：`p.Range.InlineShapes.Count > 0` 或 `p.Range.ShapeRange.Count > 0` 即为真。`InsertFigureCaption`、`InsertSubfigureCaption`、`CaptionTitleFromContext` 三处都改走它。
- `Private Function FindFigureParagraph() As Paragraph`：先取 `ActiveDocument.InlineShapes(1).Range.Paragraphs(1)`，取不到再遍历 `ActiveDocument.Paragraphs` 找 `Range.ShapeRange.Count > 0` 的段落。`TestFigureCaption`、`TestFigureCaptionClick` 两个测试入口用它，不再硬取 `InlineShapes(1)`。

`modMain.bas` 的 `InsertPicture` 仍判断 `Selection.InlineShapes.Count`：那里处理的是用户经对话框真实插入的图片，嵌入型图片会正常计入，不受影响。

回归：`tests/test_vba_units.py::test_figure_caption_click`。

## 17. 图表编号是域，正文引用要自己写域

题注编号一直走域：`add_caption()` 依次写 `STYLEREF` 域、`.`、`SEQ` 域，所以增删图表后按 F9 就能重排。但正文里的引用原本是手打字符（“表1.1按几个常见维度做了对比。”），既不属于任何域，也不跟着重排，改一张表就得手改一遍正文。

原因是构建期只造编号、不造引用：`add_caption()` 从不写 `w:bookmarkStart`，文档里没有一个 `THU_FIG_` / `THU_TAB_` 书签，`REF` 域没有可指向的目标；而写 `REF` 的代码只存在于 VBA 宏（`modFormula.bas` 的 `InsertRefField`），构建链根本不走宏。

现在分两半补上：

- `src/build_docx.py` 的 `add_caption()` 把编号文本（`STYLEREF` + `.` + `SEQ`）包进书签，名字按 `THU_FIG_0001` / `THU_TAB_0001` 递增；`ref_caption()` 写 `REF THU_FIG_0001 \h` 并带上编号文本作为域结果。
- 公式编号在 `finalize_docx.py` 里才由 `#()` 生成，构建期算不出来，所以 `THU_EQ_0001` 书签也在 `finalize_docx.py` 建（取 `STYLEREF` 域结果起点到 `SEQ` 域结果终点）；`ref_formula()` 只写域，结果文本靠更新纠正。

踩过的两个坑：

- `finalize_docx.py` 的更新循环只处理 `SEQ`/`STYLEREF`，**`REF` 域不在其中**，所以构建期必须给 `REF` 写正确的缓存结果，否则首次打开就是空白或错误提示。
- 公式的 `REF` 域结果会被 Word 放进 `m:oMath`（数学区），`m:t` 里是 `3.1` 而不是 `w:t`。域值正确、显示正常，但用纯文本方式提取段落会取不到，写检查脚本时要按 XML 找。
- 附录公式的章节号不能沿用 `STYLEREF 1`，否则会编成正文的章号（如 `3.1`）。指南允许 `式(C.3)` 与 `式(C-3)` 两种风格，本模板统一用点号。`finalize_docx.py` 用样式名状态机切换：进 `附录标题` 后改用 `STYLEREF "附录标题" \s`，得到 `A.1`；构建期的 `peek_formula()` 同步按附录字母算缓存值。
- `ref_caption()` 的 `offset` 是前瞻语义，`0` 表示“下一个即将创建的题注”。引用刚建好的上一条题注要传 `-1`，否则引用会跳到下一条，或者因书签尚不存在而显示“错误!未找到引用源。”。

回归：`tests/test_build.py::test_no_broken_field_results`（正文不得含“错误!未定义书签”）、`tests/test_vba_units.py::test_subfigure_caption`（宏按顺序数子图注，新增示例子图注后会挤偏序号，需先记基数，避免模板自带示例子图注挤偏序号）。


## 18. Word 更新公式时可能使 COM 对象失效

`OMath.BuildUp` 和域更新可能重建公式段或嵌套域。转换后从 `Selection.Paragraphs(1)` 重新取公式段。分配书签时逐个读域代码，局部跳过失效的旧域对象。

缺少文档变量时，`Variables(name)` 返回的对象也可能在读取 `Value` 时才报 5825（对象已被删除）。在局部错误处理内直接读取字符串值，避免保存 `Variable` 对象后又在更新域时引用它。公式书签按公式段保留同名绑定，不能用全文排序复用旧的名字。见 [公式实现](80-dotm.md#公式)。

## 19. 在内容控件内新建另一控件可能忽略目标范围

保持“学位类别”下拉框选中时调用旧版 `TestCoverSelectDegree("工学博士学位")`，可复现作者区四个控件全部消失；光标位于学位或作者控件内时点击副导师／联合导师，可能只插入标签、未创建控件。从普通正文处调用则正常，因此只测宏最终结果会漏掉问题。

原因是 Word 的 `ContentControls.Add` 在当前选区仍位于另一控件时，可能受旧选区影响，即使提供了目标 Range。先删除作者区再创建控件，会将创建失败放大为数据丢失。修复必须同时处理：

- 在重入锁保护下先选择目标文字，再从该段落的集合创建控件；完成后恢复原选区。
- 行结构不变的硕博切换保留原控件；确需重建时先读取所有值、检查锁定状态，再删除。
- 重建必须放在独立撤销记录中；即使第一行成功、后续控件创建失败，也要恢复整个作者区。故障注入测试操作临时模板源码，不向交付宏添加测试开关。
- 不能只依赖退出控件事件：确认下拉选项后光标仍在控件内，英文会保持旧值。ThisDocument 订阅选区变化与保存事件，按最后已应用的学位去重，并限制在自身及附加自身的文档中。

回归见 `test_vba_units.py::GraduateMacroTest` 的封面交互用例，包含控件选中、正文光标、DOTM 新建及 DOCX 重开。

交付时必须将通过测试的文件同步回正式 `output/` 并核对哈希；只更新 `output/checks/` 中的构建副本，用户打开原路径仍会运行旧宏。测试应包含“选择学位 → 不离开控件直接刷新 → 再连续刷新”的组合路径。
