# `.dotm` 宏模板设计与维护

本册讲宏模板本身：VBA 模块分工、功能区、封面内容控件与级联、题注与文献引用的实现要点。构建步骤见 `50-build.md`，产物结构见 `30-docx-structure.md`。

两变体均可构建 DOTM，包含 9 个标准模块、ThisDocument 和两个窗体。`frmPicker` 用于选择引用目标，`frmReferenceCheck` 显示按图注、表注、引用和未引用提示分组的检查结果。两版均有 11 组功能区，本科封面组仅有密级选填菜单。

- 快速入门：[`docs/01-快速入门.md`](../../docs/01-快速入门.md)
- 格式定义：[`docs/02-格式定义.md`](../../docs/02-格式定义.md)
- 功能区定义：[`src/dotm/customUI14-graduate.xml`](../../src/dotm/customUI14-graduate.xml) / [`src/dotm/customUI14-undergraduate.xml`](../../src/dotm/customUI14-undergraduate.xml)

> 封面为封面页原生内容控件 + ThisDocument 级联；文献与图/表/公式引用仍用 `frmPicker` 复选列表。段落、字体、列表、图片、表格、公式、代码、引用均为独立功能组；“帮助”组只保留“关于”。

## 目标

封面在文档页内直接填写，不弹窗体：

- 学位类别：下拉列表，候选按 `configs/degree-options.yaml` 生成（只展开配置中存在英文名称的层次，专业学位按指南附录F列示）；
- 学科：可输入组合框，候选随学位类别与层次刷新，也允许键入目录外学科；
- 密级：研究生在中文封面独立页眉左侧，本科在封面表格首行，默认空白。`ToggleSecurity` 在无控件或空白时写入 `秘密★5年`，已有文字时删除控件并清空段落。`FindCC` 按 `THU_UNDERGRADUATE` 选择正文或首节页眉；`SecurityParagraph` 在本科只查首张表内的“封面密级”样式，避免旧稿误写到 logo。切换不增删行、不移动其他部件。
- 成文日期：研究生填 `2026-11` 之类年月，离开控件后转中文并同步英文；本科直接填写中文年月。
- 作者信息：固定 6 行，基本信息按学位类型占3或4行，余下行用“选填信息”菜单的“副导师行 / 联合导师行”命令增删副导师、联合导师：控件不存在时插入空行和控件，控件已存在时再点即删除该行（即使已填姓名），删除后保留空行；副导师、联合导师分别同步英文封面的 `Associate Supervisor:`、`Co-supervisor:` 行；
- 中英文题目、英文作者、英文导师姓名：直接填写；
- 检查宏：`modCheck.bas` 只做题注与交叉引用四类检查：每张图片是否有对应的图注/子图注且题注以域格式存在（含 `SEQ`/`REF` 域与 `THU_*` 书签）、每个表格是否有对应的表注且题注以域格式存在、图表题注是否被正文引用（未被引用者单独列为“未引用”提示分组，不计入硬错误；检查按显式引用统计，仅写“表4.2至表4.4”时，表4.3仍会显示为未引用）、交叉引用和文献引用的目标是否仍然存在（含书签、自定义属性和文献条目）；结果用一个弹窗列出。宏只读、不更新域、不写文档；图表配对检查只覆盖正文与附录，封面、目录、名单页等前置部分不参与；按钮为“引用”组的“引用检查”（两变体都有），建议先单击“引用 → 更新域”再检查。不检查的内容：封面内容控件、公式编号、题注文字格式（标签/全角空格）、编号重复、必备样式、标题/题目字数等内容项，以及字号、行距、段距、页眉页码。

学位类别变更后级联派生（只改派生字段，已填的英文题目/姓名/导师姓名不动）：

- 中文学位句：`(申请清华大学` + 学位类别内容 + `论文)`；
- 学位类别显示串：学术 `{门类}{层次}学位`（如工学硕士学位），专业 `{类别}{层次}专业学位`（如金融硕士专业学位）；
- 英文 degree 行：学术 `for the degree of`，专业 `for the professional degree of`；
- 英文提交声明：硕士 `Thesis submitted to`，博士 `Dissertation submitted to`；
- 博士导师标题：`Thesis Supervisor` → `Dissertation Supervisor`；英文学位名称由 VBA 函数 `CoverDegreeEnglish` 生成（如 `Master of Science`）；
- 中文日期转汉字（如 `二〇二六年十一月`），英文日期 `November, 2026`（加粗）。

## VBA 模块

| 模块 | 主要职责 |
| --- | --- |
| `modMain.bas` | 功能区入口、标题、书签、更新域及附录编号复位 |
| `modCover.bas` | 封面内容控件读写、学位与日期级联、作者信息与可选导师行重建、英文导师制表位、密级与可选导师行按内容控件存在性开关、英文 Associate Supervisor / Co-supervisor 行同步 |
| `modFormula.bas` | 行内公式、行间公式（有编号/无编号）、`#()` 编号、公式书签、公式引用、编号刷新 |
| `modCaption.bas` | 图注/表注/子图注、题注书签、图/表引用；`HasFigureAt()`/`FindFigureParagraph()` 同时识别嵌入型图片与 `wps` 原生图形 |
| `modCitation.bas` | 参考文献条目扫描、文献选择器、多篇文献合并上标、更新域整理、导出前冻结引用域 |
| `modPicker.bas` | 可滚动复选列表窗体逻辑 |
| `modExport.bas` | “导出”菜单回调，在副本更新并生成 DOCX 或 PDF |
| `modCheck.bas` | 题注与交叉引用检查 |
| `modDegreeOptions.bas`（构建生成） | 学位候选、专业领域开关、变体标记及旧稿续表识别常量 |
| `ThisDocument`（由 `ThisDocument-graduate.txt` 生成） | 文档打开／新建时订阅 Word 选区与保存事件；学位确认后同步，退出控件时补充级联 |

功能区只有一个“论文”选项卡。两版均为 11 组，包括封面、段落、字体、列表、图片、表格、公式、代码、引用、导出、帮助。本科封面组仅有“选填信息 → 密级信息”。两变体的“引用”组均有“引用检查”按钮（调用 `modCheck.bas` 的只读题注与交叉引用检查）；“帮助”组仍只有“关于”，该按钮不在帮助组。“附录章节”为小菜单，列表组的“调整编号”为小菜单并放在组末尾。研究生封面组为 2 个直接控件：“刷新封面”大按钮与“选填信息”大菜单（`imageMso=FilePrepareMenu`，Word 内置标签“准备”）；菜单内为“副导师行”“联合导师行”“密级信息”3 项，菜单条目无图标。所有按钮都可用，不使用 `getEnabled` 灰显。

## 技术路线

1. `src/build_dotm.py` 默认在临时目录调用 `src/build/assemble.py` 生成 docx，再由 Word COM `SaveAs2(..., FileFormat=15)` 生成 `.dotm`；不再固定读取 `output/` 的 docx。
2. 功能区：注入变体 `customUI14-graduate.xml` + `_rels/.rels` 关系，VBA 回调。本科变体保留封面组，仅提供“选填信息 → 密级信息”；没有学位/日期级联。
3. 构建时以 `src/dotm/forms/forms.json` 创建两个窗体，从对应的 `.code.txt` 注入代码。选择器支持拖动边缘调整大小，并在选中条目后启用“插入”。检查结果按选项卡显示，详单可滚动和复制。
4. 内容控件：按 `Tag` 定位，不依赖控件顺序。研究生默认正文中文7个、英文9个，另有封面页眉密级控件，tag 契约见下节。
5. 级联：研究生用 `src/dotm/vba/ThisDocument-graduate.txt` 在构建时写入 ThisDocument。`Document_Open`、`Document_New` 和控件进入事件初始化 `WithEvents Word.Application`；选项确认后的 `WindowSelectionChange` 与 `DocumentBeforeSave` 调用 `RefreshDegreeIfChanged`，按文档变量 `THU_DEGREE_APPLIED` 去重。事件只处理本模板及附加本模板的当前文档。`Document_ContentControlOnExit` 保留学位、密级和日期分发。`mUpdatingCover` 阻止程序修改选区或重建控件时再次进入级联。本科 ThisDocument 不订阅这些事件。
6. 可选导师行：“选填信息”菜单的“副导师行 / 联合导师行”命令在作者信息单元格插入 `\t{标签}：\t` + 空文本内容控件（tag 为 `副指导教师（可选）` / `联合指导教师（可选）`）；删除时删内容控件并清空段落、保留空行；每次将可选导师紧接基本信息，按“副在前、联合在后”重排并重算标签宽度。
7. 英文同步：`ToggleCoSupervisor` 在增删联合导师行后调用 `SyncEnglishCoSupervisor`，开启时在英文导师控件第一行后追加 `Co-supervisor:` 行，关闭时移除，保留已填的英文导师姓名与 `Thesis/Dissertation Supervisor` 标题。
8. 校验：“引用”组的“引用检查”按钮（两变体都有）调用 `modCheck.bas` 的只读宏，一次检查四类题注与引用问题：图片图注/子图注与表格表注是否存在且以域格式存在（`SEQ`/`REF` 域与 `THU_*` 书签）、图表题注是否被正文引用（未被引用者单独列为“未引用”提示，不计入硬错误）、`REF`/`DOCPROPERTY` 交叉引用的被引用项是否存在（书签/自定义属性缺失或结果显示“错误”即报错）；结果用一个弹窗列出。宏不更新域、不写文档；配对检查以首个「标题 1」章标题为正文起点（`bodyIdx`/`bodyStart`），前置版式表与封面徽标因此被排除；用章标题而非题注自身当分界，是为避免用户改写或删除早期题注后分界后移、令前面的表格与图片漏检（找不到章标题时回退到首个题注/图片段，仍找不到才从全文第一段开始）；建议先单击“引用 → 更新域”再检查；封面内容控件、公式编号、题注文字格式、编号重复、必备样式、字数等内容项与字号、行距、段距、页眉页码均不检查。

## 源码布局

```text
src/
  build_dotm.py              # 临时 docx → dotm → 导入 VBA 与 ThisDocument → 创建窗体 → 注入 customUI → 冒烟
  build/                     # assemble / variants / degree_catalog
  build/content/graduate/    # 研究生 cover / front_matter / back_matter
  build/content/undergraduate/ # 本科 cover / front_matter / back_matter
  build/content/examples.py  # 共用正文、摘要布局及符号说明
  build/content/appendix.py  # 共用附录
  dotm/
    customUI14-graduate.xml           # 研究生 11 组功能区
    customUI14-undergraduate.xml  # 本科 11 组功能区（封面组仅有密级菜单）
    icons.yaml               # 直接控件 → imageMso
    forms/forms.json         # 两个窗体的控件布局
    forms/*.code.txt         # 窗体事件代码（UTF-8，构建时注入）
    vba/*.bas                # 共享标准模块，构建时另生成 modDegreeOptions
    vba/ThisDocument-graduate.txt     # 研究生文档事件代码，构建时写入
    vba/ThisDocument-undergraduate.txt  # 本科无事件最小文件
tests/                       # 静态格式/样式契约 + Word 宏与导出回归
```

## 内容控件 tag 契约

默认正文16个（专业学位无领域时移除中文学科控件）：`论文题目`、`学位类别`、`培养单位`、`学科（专业学位改为专业领域）`、`研究生（专业学位改为申请人）`、`指导教师`、`成文日期`、`英文论文题目`、`英文提交声明`、`英文degree行`、`英文学位名称`、`英文学科`、`英文专业领域`、`英文作者`、`指导教师（英文）`、`英文成文日期`。

可选 3 个：`密级`、`副指导教师（可选）`、`联合指导教师（可选）`。其中 `密级` 位于研究生首节独立页眉或本科封面表格首行、默认空白，由“选填信息 → 密级信息”写入 `秘密★5年` 或清空；后两个控件由菜单中的导师行命令按存在性开关增删。

`学位类别` 为下拉内容控件，`学科（专业学位改为专业领域）` 为可输入组合框；`成文日期`、`英文提交声明`、`英文degree行`、`英文学位名称`、`英文成文日期` 构建时写真实内容，不带占位符。研究生 `密级` 控件在首节主页眉，本科在封面表格首行，默认空白。两版均用 `modCover.ToggleSecurity` 写入 `秘密★5年` 或清空，保留本科密级行。研究生默认总控件数为17（正文16、页眉1），本科为 7（均在正文）；删除密级后各少一个。

改 tag 必须同步 VBA、`src/dotm/vba/ThisDocument-graduate.txt`、模板说明页、本文件。

## 关键实现点

### 公式

- 用 `doc.OMaths.Add` 插入公式后，必须从段落 `p.Range.OMaths(1)` 取回 `OMath` 对象；`OMaths.Add` 不返回对象。
- 编号区写 `#(` + `STYLEREF`/`SEQ` 域 + `)`；书签加在编号数字上。Word 的隐藏书签用下划线前缀表示，`Bookmarks.ShowHidden=False` 时不会遍历到，但 `Bookmarks.Exists("_...")` 仍返回 True；宏为了扫描简单使用显式 `THU_EQ_` 书签。
- 公式引用使用 `REF ... \h \* MERGEFORMAT` 域，域结果字体设为 XITS Math。
- `StripTrailingNumber()` 只处理带 `#` 标记的 Word 公式编号。普通括号属于公式本体，例如 `f(1)`，不能据此推断手写编号。
- `FormulaBookmarkName()` 从公式所在段取已有的非空 `THU_EQ_` 书签。重新编号时，`BindFormulaBookmark()` 用同名书签覆盖章号结果起点到序号结果终点。构建时生成的 `THU_EQ_0001` 等名称继续有效。
- 新书签使用 `THU_EQ_N<n>`，由文档变量 `THU_EQ_LAST_ID` 记录已用序号。分配时还避开已有书签和正文 `REF` 使用的名字，防止原目标删除后旧引用误指新公式。引用选择器从当前段读取书签和域结果。
- `RefreshFormulaNumbers()` 按“公式”样式查找有编号的 `oMathPara`，更新域并恢复书签。`UpdateAllFieldsNow()` 先整理附录序号，更新正文、页眉页脚和清单，再刷新公式书签并重算 `REF THU_EQ_…`。无编号公式不进入引用列表。
- 行间公式下拉菜单分“有编号”“无编号”（`mnuDisplayFormula` → `btnDisplayFormulaNumbered`/`btnDisplayFormulaPlain`）。有编号应用“公式”样式，`#()` + `STYLEREF`/`SEQ` 生成 `(2.1)`/`(A.1)`；无编号只套样式并居中。行内公式仍为独立按钮；公式引用重复点击幂等。

### 题注

- 图注在图片段之后，表注在整张表之上（表格前一段），不写进表格单元格。
- 「段落里有没有图」用 `HasFigureAt(p)` 判断（`InlineShapes` 或 `ShapeRange` 任一非空）。示例插图的 `wps:wsp` 占位图形只出现在 `ShapeRange`，详见 `70-pitfalls.md` 第 16 条。
- 题注编号由 `STYLEREF` + `SEQ` 域生成；每个题注的编号文字建 `THU_FIG_`/`THU_TAB_` 书签。
- 标签与编号无空格，编号与题名一个全角空格（`captions.defaults.titleGap`）。
- 续表生成入口及测试入口已注释，功能区无按钮；`IsContinuedTableCaption()` 保留，只读识别旧稿，不新增题注。
- 子图注用独立“子图注”样式，编号固定 `SEQ 子图 \* alphabetic`；需要重新从 (a) 开始时，选中该子图注后点“列表→调整编号→重新开始于1”，`RestartSubfigureAtSelection` 把该域改成 `SEQ 子图 \r 1 \* alphabetic`，不使用隐藏复位域。

### 交叉引用与书签

- 本宏的图、表、公式引用采用**显式书签 + `REF` 域**：题注编号处建立 `THU_FIG_`/`THU_TAB_`/`THU_EQ_` 书签，引用处插入 `REF 书签 \h`。这与 Word 原生“交叉引用”不同：后者内部建立 `_Ref...` 隐藏书签。两者底层都是 `REF`，但宏不依赖 Word 生成的 `_Ref` 书签，因此复制、合并或重建后仍可由宏补建。
- 章节引用指向标题段落上的 `THU_CH_` 书签，插入 `REF 书签 \n \h`（`\n` 取该段落编号，`\h` 生成超链接）。必须指向具体标题书签；用 `STYLEREF "标题 1" \n` 只能取插入点之前最近的同样式标题，会错误地引用“当前章节”。
- 当前书签为显式书签，会出现在 Word“书签”对话框中；用户不需要手工管理。
- 可以建立隐藏书签：Word 中以 `_` 开头的书签默认隐藏；`Bookmarks.ShowHidden=False` 时不会遍历到，需临时设为 `True`（或直接用 `Bookmarks.Exists`）。当前 Word 对象模型没有 `Bookmark.Hidden` 属性，隐藏状态由名称前缀控制。

### 参考文献引用

- 文献条目用“参考文献条目”段落样式；宏按顺序扫描并给每条建 `THU_REF_ENTRY_0001` 书签，编号写入自定义文档属性 `_THU_REF_NUM_0001`。
- 正文引用是一个 `QUOTE "THU_CITE:0001;0003"` 域，域代码保存条目 id；`NormalizeCitations` 遍历 `ActiveDocument.Fields`，按属性把 id 翻译成当前编号，升序、去重、合并连续区间，写回域结果为 `[1]`、`[2-3]`、`[7, 9, 11-13]`，并套“文献引用”字符样式。
- 多选插入时把所选条目 id 用分号拼接，只插入一个域；不再逐条插入。
- “更新域”`UpdateAllFieldsNow` 在 `Fields.Update` 之后调用 `NormalizeCitations`，核验参考文献书签，并刷新和锁定文献域的显示结果。普通 F9 保留最近一次显示的文献编号。
- 导出 DOCX 或 PDF 时 `FreezeCitationFields` 把 `THU_CITE:` 域与旧 `_THU_REF_NUM_` 域 `Unlink` 成静态文字，固定导出时的编号。导出在写作稿副本进行。
- `DOCPROPERTY` 引用域仍按属性显示编号，不参与合并（兼容既有文档）。
- 测试入口：`TestCitationCompress`（压缩规则）、`TestCitationMerge`（合并/去重/重复整理）、`TestCitationUpdateRoundTrip`（F9 前后编号保持正确）、`TestCitationFreeze`（冻结前后域数）；导出冻结由 `tests/test_vba_units.py::GraduateCitationExportTest`（本科为 `UndergraduateCitationExportTest`）向 dotm 副本注入真实 `QUOTE` 域后走 `ExportDocxFile` 验证（不调用 `Document.Save()`，见 `.agents/docs/70-pitfalls.md` 第 10 条）。

### 标题与列表

- 内置样式通过 `w:aliases` 设置中文替换名，原名和内置身份保留。`PrimaryStyleName()` 去掉 `NameLocal` 的别名部分，标题、列表与引用检查均据此判断；`StyleExists()` 支持原名和替换名。功能区标题入口使用替换名。列表规则只定义 1–3 级。

- 章内无编号小节用独立样式 `章内无编号小节`：无编号、`outlineLevel=bodyText`、居中；附录“参考文献”和个人简历小标题共用。
- `ApplyHeadingStyleByName()` 只套标题样式；全角间隔已写入自动编号，保留原段落文字、书签、域及局部格式。
- “其他章节 → 章标题”由 `InsertOtherSection()` 用 `THU_UNDERGRADUATE` 区分候选。本科追加在学期间参加课题的研究成果，研究生追加个人简历及学术成果、指导教师（小组）评语和答辩委员会决议书。该入口插入下一页分节符。本科目录标题的段级大纲设为正文，其他前置标题直接使用不进目录样式。章内无编号小节入口只套样式。
- 列表组“调整编号”是下拉菜单：“继续编号”调用 `ApplyListTemplateWithLevel(..., ContinuePreviousList:=True)`，“重新开始于1”调用 `ApplyListTemplateWithLevel(..., ContinuePreviousList:=False)`；两者都先保存、后恢复段落缩进，修复 Word 重启编号后的格式。
- “参考文献”是下拉菜单，按 `control.Tag` 分别套“参考文献条目”（顺序编码制，带 `[1]` 编号）或“参考文献条目（著者出版年）”（无编号，保留悬挂缩进）。
- “参考文献”菜单之后是“缩略语”按钮（调用 `InsertSymbolEntry`）：空行写入“缩略语<Tab>说明”，有文字无制表符时补“<Tab>说明”，已有制表符时只套“符号说明条目”样式。样式本身为：符号在版心左缘，说明列 12 字符（144 pt ≈ 5.08 厘米），续行悬挂 12 字符。
- 列表：编号 1–3、项目符号 1–3，基于正文段落格式，设 2 字符悬挂缩进；二级、三级另设 2、4 字符左缩进以表示层级；选中多个列表段落时逐个套用。

### 附录编号

- 插图/附表清单两变体统一用 `TOC \h \z \c "图"` / `\c "表"`（按题注标签收集）。
- 附录内图注/表注改用 `附录图注`/`附录表注`，视觉继承“图注/表注”；题注宏按 `IsAppendixAt` 自动把 SEQ 标识符切到 `附录图`/`附录表`，`\c` 清单只收正文的 `SEQ 图`/`SEQ 表`，故附录图表不进清单。
- `RefreshAppendixSequences()` 按最近的“附录标题”与对象类别分组，首个编号域用 `\r 1`，后续沿用 `\s 1`。附录公式通过同段的 `STYLEREF "附录标题"` 识别，正文的 `SEQ 式` 保持原逻辑。
- 该过程由“更新域”、附录题注和附录公式插入路径调用。若首个编号域出现在旧的导语复位域之前，将后面的 `\r 0 \h` 改为 `\c \h`，避免同一附录重复从 1 开始。

### 封面

- 封面为原生内容控件：研究生默认共 17 个（正文中文 7、英文 9，另有页眉密级 1 个）；本科默认共 7 个（全部在正文，含密级）。Word 的 `Document.ContentControls` 不含页眉，研究生密级须单独从首节主页眉的 `Range.ContentControls` 访问。学位类别是 `dropDownList`（候选按 `configs/degree-options.yaml` 生成，默认“工学硕士学位”）；学科是 `comboBox`（默认“计算机科学与技术”，可输入）。成文日期、英文提交声明、英文 degree 行、英文学位名称、英文成文日期默认写真实内容，不使用占位符。
- 中文作者信息块固定 6 行高度，默认构建写入4个基本项和2个空段；专业学位无领域时重排为3个基本项，其后紧接可选导师；标签用 `w:fitText` 等宽（4/5/6 字分别为 1280/1600/1920 twips），制表位 159/164 pt。
- 英文导师行先清除样式自带的 183.7/195.6 pt 制表位，再设 195.5/200.5 pt；否则首个 tab 会落到旧位置，冒号后没有间隙。英文学位名称、英文作者、英文日期 Arial 加粗。
- 学位选项确认引起的选区变化会同步英文派生字段；退出控件与保存前补充刷新，避免光标还停留在下拉框时一直显示旧值。`ApplyDegreeCascade` 更新英文学位名称、提交声明、degree 行、导师标题、中文标签和学科候选；只有学科／专业领域行的有无改变时才重建作者区，通常的硕博切换保留原控件 ID 和文字。日期仍在退出控件时处理，把 `2026-11` 转成 `二〇二六年十一月`，填写内容不含“年”时同步英文日期。
- 新建作者控件前必须在重入锁内执行 `rngCC.Select`，再调用目标段落的 `Range.ContentControls.Add(..., rngCC)`。仅传 Range 参数仍可能受旧选区影响，造成控件加到光标处或创建失败；在删除作者区后失败就会表现为全块清空。重建前先读取全部值并检查锁定状态，完成后按原 Tag 恢复选区；无控件的正文选区通过跟随编辑的 Range 恢复。
- 作者区重建使用独立的 `UndoRecord`，删除、创建及格式调整合为一次操作。中途失败时结束记录并撤销该次重建，恢复原有作者文字与控件，再向外报告错误。回归通过向临时模板注入“第一行已创建、第二行前失败”的故障验证恢复；不能仅验证正常完成的路径。连续刷新测试直接选择下拉条目后调用功能区入口 `RefreshCover` 三次，覆盖带两位可选导师的学术硕博、无领域专业学位与有领域专业学位。
- 副/联合导师行由“选填信息”菜单的“副导师行 / 联合导师行”命令增删：在作者信息块后 2 行插入或删除带内容控件的行（tag 分别为 `副指导教师（可选）`、`联合指导教师（可选）`），删除后保留空行，按“副在前、联合在后”重排并重新等宽标签。副导师与联合导师行命令均调用 `SyncEnglishCoSupervisor`。
- 两版“密级信息”均调用 `modCover.ToggleSecurity`：无控件时在本科封面密级行或研究生首节页眉插入控件，写入 `秘密★5年`；控件为空时直接填写；已有文字时删除控件并清空段落。公开与涉密切换保留本科密级行高度。副导师行、联合导师行、密级信息三个动作失败时只写状态栏，不弹消息窗口。
- “刷新封面”按当前学位类别与成文日期重算派生内容。`modCheck.bas` 的“引用检查”只覆盖题注与交叉引用四类，封面必填内容控件、占位符与题目字数均不校验；封面需作者自行核对。
- 封面不使用窗体。

封面回归在 `GraduateMacroTest`：覆盖学位控件选中时反复切换、控件内确认即同步、各作者控件及正文光标位置增删导师行、跨学位类型保留已填写导师、锁定预检、DOTM 新建与 DOCX 保存重开，以及不改写无关文档。不能只调用级联函数后检查最终文字；必须同时检查控件数量、所在单元格、行顺序和固定六段。

### SEQ 复位域重写

- 若附录复位域的 `\r` 被写成回车，域代码变成 `SEQ 表 0 \h`，Word 把 `0` 当书签名，显示“错误!未定义书签”。
- 构建用 raw string 写 `SEQ \r 0 \h`；`modMain.RepairAppendixResetFields` 在打开模板、首次调用宏和“更新域”时扫描并重写这类域。

### 帮助

- “帮助”组只保留“关于”，功能区不提供内置帮助窗口。
- “关于”的名称取自 `configs/variants.yaml` 中对应变体的 `title`，版本取自 `configs/release-version.txt`。发布时先填写拟用的 Git 标签，再构建两版 DOTM。
- 面向用户的说明文档以 `docs/01`、`docs/02`、`docs/03`、`docs/04` 独立维护，不随 dotm 嵌入，也不参与构建。

## 功能范围

- 公式：行间公式（有编号/无编号）、行内公式、公式引用、编号刷新。
- 题注：正文图注/表注、附录表注、子图注、图/表/公式/章节引用。
- 文献引用：条目扫描、复选列表多选合并、更新域整理、导出冻结。
- 封面：原生内容控件 + 级联，学位/学科下拉，派生字段动态刷新，可选导师行与密级按控件存在性开关。
- 段落：正文章节（章 + 1–3 级节）、附录章节、其他章节（共用前后置章节及各变体专属条目）按菜单套样式。字体组只放“强调”菜单。
- 列表：编号 1–3、项目符号 1–3、“调整编号”菜单、“参考文献”菜单、“缩略语”按钮。
- 图片、表格、代码、引用组可用；帮助组只保留“关于”。
- 更新全部域；“导出”组支持生成 DOCX 和 PDF。
- 功能区图标：直接控件全部配置 `imageMso`，导出菜单的两个选项也按文件格式配图标；映射在 `src/dotm/icons.yaml`，对照表在 `85-ribbon-icons.md`；候选池 `references/icons/imageMso-pool.yaml`（1,494 个有效 ID，只读参考数据，不参与 dotm 构建），图库 `references/icons/imageMso-pool/gallery.html`；校验与导出脚本同在 `references/icons/`（`check_icons.py` / `export_control_icons.py` / `export_icon_pool.py`），不参与 dotm 构建。`.icon` 不是标准格式，Office 不保存 imageMso 图片；自定义图片用 PNG。

## 后续计划

- 中英文对照题注（见 `TODO.md`）；“引用检查”已覆盖题注与交叉引用四类，字号、行距、段距、页眉页码等版式数值按设计不自动检查，不再列为待办。
- 封面采用封面页原生内容控件 + `ThisDocument` 级联，学位/学科下拉，派生字段动态刷新；英文题目、英文姓名、英文导师姓名仍手填。
- 学位选项配置在 `configs/degree-options.yaml`（仅生成已配置层次），构建生成 `modDegreeOptions.bas`；不要再把选项硬编码回 VBA 数组。
- 功能区不提供“功能说明”，构建不生成 `modHelp`/`frmHelp`。

## 不做的事情

- 不生成续表题注。两个功能区按钮与 `InsertContinuedTableCaption`、`TestContinuedTableCaption` 已注释；保留旧稿识别逻辑和规则中的续表写法作为兼容数据。
- 参考文献条目关闭自动断字和西文词内换行，URL 保持原文。
- 不自动安排单双面、右页起排与装订空白页。

## 限制

- 宏默认禁用，用户需“启用内容”，或放受信任位置；组策略可能完全禁止。
- 研究生封面级联由文档事件触发。“刷新封面”可按当前文档的控件值手动重算。写作稿从 `.dotm` 新建，功能区和宏由 Word 从关联模板加载。
- WPS 免费版通常无 VBA；Mac Word 对 UserForm 支持有限。必须保留“无宏也能填内容控件”的降级路径。
- 最终提交 `.docx` 会去宏，但封面文字保留；不要依赖宏在打印时动态生成内容。
- 纯 VBA 不能自由定制功能区；用 `customUI14-graduate.xml` 注入按钮，VBA 只做回调。
- 宏模板单独版本管理，不覆盖 docx 母版。

## 本次封面修正

- 学术博士统一 `Doctor of Philosophy`，学术硕士使用 `Master of Arts` 或 `Master of Science`。专业学位英文名及允许层次由 `degree-options.yaml` 列明，缺少层次时不生成选项。
- `RebuildCoverAuthorInfo()` 按有无专业领域紧凑重排六行作者信息，保留各标签的用户文本。无领域时移除学科控件，恢复时从文档变量取回。
- `英文学科` 只含学科名称，`in` 在前一独立段落中，不属于内容控件；不得用软换行把两行合并。切为专业学位时清空 `in` 的普通文字并保留段落高度，恢复学术学位时重新写入。`英文专业领域` 承载作者名下的括号领域。非当前类型已填内容保存在 `THU_ACADEMIC_EN` / `THU_FIELD_EN` 文档变量中；英文内容由作者填写，不自动翻译。
- 学术学位下拉使用指南附录C列明的11个授予门类，专业学位的33个名称／层次组合与附录F一致，总计55个选项。学科名称完整保留附录D的59项；`cross_disciplinary` 单列集成电路科学与工程、设计学，`discipline_awards` 按[国家目录](https://www.moe.gov.cn/srcsite/A22/moe_833/202209/t20220914_660828.html)补充跨门类授予关系。例如统计学可选理学／经济学，集成电路可选理学／工学，不生成“交叉学科硕士／博士学位”。附录E的跨学科项目不自动等同于交叉学科门类。
- 指南没有列出现行专业领域清单，专业领域组合框候选保持空，允许手填；附录G只适用于2019级及之前，不映射到现行类别。候选为空时不回退成学位类别名称。`degree_catalog.discipline_options()` 同时供 DOCX 和 VBA 生成器使用，避免两套候选不一致。
- `SyncPublicDeclaration()` 随研究生密级按钮或密级控件退出事件切换学校声明中的“不包含涉及国家秘密的内容”，保留其他声明文字。本科声明不联动。
- 章节引用范围包含正文章节与附录各级节标题，附录小节是否进入目录仍按现行规则保持不变。
