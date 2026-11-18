# 功能区按钮图标对照表

本表列出功能区按钮、菜单入口和带格式图标的导出选项。

图标来源：本机 Word 已验证的 `imageMso`（候选池见 `references/icons/imageMso-pool.yaml`，共 1,494 个有效 ID）。导出图库：`references/icons/imageMso-pool/`（仅供本地预览，Office 图标版权归 Microsoft）。图标 ID 校验与图库导出脚本同在 `references/icons/`。

| 功能组 | 控件 | 按钮文字 | 大小 | imageMso | Word 图标标签 |
| --- | --- | --- | --- | --- | --- |
| 封面 | `btnRefreshCover` | 刷新封面 | large | `GroupAddins` | 加载项 |
| 封面 | `mnuOptionalInfo` | 选填信息 | large | `FilePrepareMenu` | 准备 |
| 段落 | `mnuChapter` | 正文章节 | large | `TextFromFileInsert` | 文件中的文字... |
| 段落 | `mnuAppendix` | 附录章节 | normal | `BodyTextHide` | 隐藏正文 |
| 段落 | `mnuOtherSection` | 其他章节 | normal | `TextFromFileInsert` | 文件中的文字... |
| 段落 | `btnBodyText` | 正文 | normal | `FontDialog` | 字体... |
| 字体 | `mnuEmphasis` | 强调 | large | `StylesPane` | 样式... |
| 列表 | `mnuOrderedList` | 有序列表 | large | `NumberingGalleryWord` | 编号 |
| 列表 | `mnuBulletList` | 无序列表 | large | `BulletsGalleryWord` | 项目符号 |
| 列表 | `mnuReferenceFormat` | 参考文献 | normal | `LabelInsert` | 标签 |
| 列表 | `btnSymbolEntry` | 缩略语 | normal | `SymbolInsertGallery` | 符号 |
| 列表 | `mnuAdjustNumbering` | 调整编号 | normal | `NumberingRestart` | 重新开始编号 |
| 图片 | `btnInsertPicture` | 插入图片 | large | `PictureInsertFromFile` | 图片... |
| 图片 | `btnFigureCaption` | 图注 | normal | `CaptionInsert` | 插入题注... |
| 图片 | `btnSubfigureCaption` | 子图注 | normal | `CaptionInsert` | 插入题注... |
| 表格 | `btnInsertTable` | 插入表格 | large | `TableInsertGallery` | 表格 |
| 表格 | `btnTableCaption` | 表注 | normal | `CaptionInsert` | 插入题注... |
| 公式 | `btnInlineFormula` | 行内公式 | large | `EquationScriptGallery` | 上下标 |
| 公式 | `mnuDisplayFormula` | 行间公式 | large | `TableFormulaDialog` | 公式... |
| 代码 | `btnInlineCode` | 行内代码 | normal | `FieldCodes` | 查看域代码 |
| 代码 | `btnBlockCode` | 行间代码 | normal | `GroupCode` | 代码 |
| 引用 | `mnuCrossRef` | 引用内容 | large | `CrossReferenceInsert` | 交叉引用... |
| 引用 | `btnCitation` | 引用文献 | large | `CitationInsert` | 插入引文 |
| 引用 | `btnFootnote` | 脚注 | large | `FootnoteInsert` | 插入脚注 |
| 引用 | `btnUpdateFields` | 更新域 | large | `FileStartWorkflow` | 工作流 |
| 引用 | `btnTemplateCheck` | 引用检查 | large | `FieldsUpdate` | 更新 |
| 导出 | `mnuExport` | 导出 | large | `FileSaveAs` | 另存为 |
| 导出 | `btnExportDocx` | 导出为 DOCX | 菜单项 | `FileSaveAsWordDocx` | 文档 |
| 导出 | `btnExportPdf` | 导出为 PDF | 菜单项 | `FileSaveAsPdfOrXps` | 发布为 PDF 或 XPS |
| 帮助 | `btnAbout` | 关于 | large | `About` | 关于 |

## 说明

- 29 个直接控件全部已配置图标；导出菜单的两项使用格式图标，其余 32 个菜单项沿用纯文字。
- 研究生封面组 2 个直接控件：刷新封面用 `GroupAddins`，选填信息菜单用 `FilePrepareMenu`（菜单内为副导师行、联合导师行、密级信息）；公式组 2 个直接控件：行内公式与行间公式菜单（菜单内为有编号、无编号）；列表组“参考文献”菜单之后是“缩略语”按钮（`SymbolInsertGallery`），最后的“调整编号”小菜单内为继续编号、重新开始于1；正文章节与其他章节共用 `TextFromFileInsert`；附录章节使用 `BodyTextHide`；正文使用 `FontDialog`；强调使用 `StylesPane`。
- `.icon` 不是 Office 支持格式；`imageMso` 不保存图片，自定义图标请用 PNG（16/32/64 px），Windows 文件图标才用 `.ico`。

本科封面组仅复用 `mnuOptionalInfo` 菜单及 `btnSecurity` 条目，菜单图标同为 `FilePrepareMenu`。
