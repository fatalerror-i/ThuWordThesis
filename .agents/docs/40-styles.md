# 样式体系：内置样式、继承、题注与列表

## 规模

当前规则 **60 个样式**，总样式数还包括 Word 内置及潜在样式，以产物为准。只保留实际需要的自定义样式；不添加未使用样式。

## 内置样式复用

生成器优先复用 Word 内置样式：

| 规则键 | Word 内置原名（实现查找） |
| --- | --- |
| `body` | Normal |
| `chapterTitle` | Heading 1 |
| `sectionTitle1/2/3` | Heading 2/3/4 |
| `caption` | Caption |
| `tocLevel1/2/3` | `toc 1/2/3` |
| `listOfFiguresEntry` | `toc 4` |
| `tableOfFigures` | table of figures（图表目录） |
| `header` | Header |
| `footer` | Footer |
| `footnoteText` | Footnote Text |
| `subtleEmphasis` | Subtle Emphasis |
| `emphasis` | Emphasis |
| `strong` | Strong |
| `hyperlink` | Hyperlink |
| `footnoteReference` | footnote reference |

自定义样式只用于内置样式无法表达的部分：无编号章标题（2 种）、名单页小节标题、附录标题/一级节标题、符号说明条目、公式、行内公式、图片、参考文献条目、封面 8 样式、授权说明文字/标题、图注/子图注/表注/图表附注。六个列表样式复用内置样式。

## 命名

- 图注、表注是当前使用的样式名。
- **图表附注**（`figureTableNote`）用于图、表下方的资料来源与附注，单倍行距。
- 内置样式替换名在 `configs/implementation-map.yaml` 的 `styleMap.*.displayName` 定义，由 `style_presentation.py` 写入 `w:aliases`。`name` 保留 Word 内置原名，便于 Python、目录域和编号继续识别；不另建同名自定义样式。
- 标题替换名为“章标题（有编号）”“一级节标题”“二级节标题”“三级节标题”；`toc 1–3` 为“目录一级条目”“目录二级条目”“目录三级条目”，`toc 4` 为“清单条目”；`Strong` 为“明显强调”。
- Word 的 `NameLocal` 会返回“原名,替换名”。宏用 `PrimaryStyleName()` 取原名判断层级，`StyleExists()` 直接按原名或替换名查找；功能区标题入口使用中文替换名。
- `图注`、`表注` **basedOn Caption**；`子图注` basedOn `图注`（用独立 `SEQ 子图`，避免进入插图清单）；`图表附注` basedOn None（不能继承题注）。
- 规则文件里的中文注释必须与最终显示名一致。

## 继承与字体

- `docDefaults`：中文宋体、西文 Times New Roman、12pt；段前段后 0，最小值行距 20pt（`pageSetup.documentDefaults`，写入 `w:rPrDefault` 与 `w:pPrDefault`）。
- Normal 不声明字体，只负责正文缩进/行距；**其他样式不继承 Normal**，避免把首行缩进带进标题/题注。
- 语义继承链只在实现层：
  - Heading 2→Heading 1、Heading 3→Heading 2、Heading 4→Heading 3；
  - 图注/表注→Caption；子图注→图注；
  - `toc 2`→`toc 1`、`toc 3`→`toc 2`、`toc 4`→`toc 1`；
  - 英文封面题目→中文封面题目；
  - 授权说明标题→章标题（无编号），名单页小节标题→Heading 2，但均显式 `numId=0`；
  - 附录标题→Heading 1、附录一级节标题→Heading 2、附录二级节标题→Heading 3、
    附录三级节标题→Heading 4（与 `configs/implementation-map.yaml` 的 `fallback` 一致，
    字号因此是 16 / 14 / 13 / 12 pt）。
- `src/postprocess_docx.py` 的 `BASE_MAP` 负责在 Word 保存后恢复这些链。
  **注意**：`BASE_MAP` 与 `implementation-map.yaml` 的 `fallback` 必须一致；附录二、三级节标题分别继承 Heading 3 / Heading 4（13 / 12 pt）。

## 标题编号与长URL

正文章、节及附录各级的 `numbering.lists.*.levels.*.text` 末尾带一个全角空格，`suffix` 为 `none`。间隔属于自动编号，生成器和宏不再向标题文字插入前导空格；直接在Word中套用标题样式也得到相同效果。引用和题注仍取纯编号，不能把间隔带入 `4.1`、`A.1` 等结果。

文档断字由两版规则的 `pageSetup.hyphenation` 指定：`automatic: true`、`hyphenateCaps: false`、`consecutiveLineLimit: 2`。`documentDefaults.paragraphFormat` 设置 `hyphenation: true`、`wordWrap: false`，使普通样式默认按英文单词断字，不任意拆开西文单词。

`formula`、`blockCode`、`referenceEntry`、`referenceEntryAuthorYear` 的段落格式显式设置 `hyphenation: false`、`wordWrap: false`，关闭自动断字，同时禁止西文在单词中间换行。行内代码、行内公式为字符样式，不控制整段断字。

规则中的 `wordWrap: false` 表示关闭“允许西文在单词中间换行”。Word 的文件语义相反，`w:wordWrap w:val="1"` 按词换行，`0` 会在字符间拆词。`apply_document_text_flow()` 将 `1` 写入包内每个样式定义，包括未使用的内置样式和表格条件格式。参见 [Open XML 的 wordWrap 示例](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.wordwrap)。

Word 另存为 DOTM 时可能省略重复的样式属性。`finalize_style_presentation()` 在 Word 关闭后对 `styles.xml` 和存在的 `stylesWithEffects.xml` 写回显式值。

正文、表格、页眉页脚、脚注尾注和文本框等部件的现有段落也统一写入 `w:wordWrap=1`，用于覆盖直接格式。操作只读取段落结构。回归先检查 XML 和 Word 实际样式，再用窄行测试验证英文单词整体移到下一行；用 `GetPressedMso("HyphenationAutomatic")` 检查功能区的“自动”是否选中。

`src/build/hyphenation.py` 由生成器和后处理共用。`hyphenation` 反向映射为 `w:suppressAutoHyphens`，`wordWrap` 反向映射为 `w:wordWrap`。文档自动断字使用独立的 `w:autoHyphenation`、`w:doNotHyphenateCaps` 和 `w:consecutiveHyphenLimit`。

## 列表样式

模板只定义编号和项目符号各 1–3 级，替换名为“列表编号1–3”“列表项目符号1–3”，与功能区入口一致。这六个样式都写 2 字符悬挂缩进；一级不写左缩进，二级、三级分别写 2、4 字符左缩进：

```xml
<!-- 一级 -->
<w:ind w:hangingChars="200" w:hanging="480"/>
<!-- 二级 -->
<w:ind w:leftChars="200" w:left="480" w:hangingChars="200" w:hanging="480"/>
<!-- 三级 -->
<w:ind w:leftChars="400" w:left="960" w:hangingChars="200" w:hanging="480"/>
```

由此一级、二级、三级的编号依次位于 0、2、4 字符处，文字位于 2、4、6 字符处，可表示多级列表。不写 `w:firstLine`；编号层的 `w:ind` 只保留悬挂缩进，不写 `w:left`，避免与样式层叠加。4/5 级不再作为模板规则或推荐样式提供。`postprocess_docx.py` 的清零循环必须跳过 `List Number*`/`List Bullet*`，并在末尾按规则恢复六个样式的左缩进与悬挂缩进，否则 Word 更新域后会丢失。

`styleId` 无需指定，Word 保存时会重新编号。实现映射用 `name` 查找内置样式，用 `displayName` 设置中文替换名，不改样式标识。

## 表格样式

`三线表` 是表格样式，不是段落样式；正文 3 个表格实际使用，不要删除。边框由规则键 `styles.threeLineTable.tableFormat` 生成（顶/底 1.5pt，表头行下 1pt，其余 none；单元格内边距 0），`add_table_style()` 单独建，**不参与段落/字符样式循环（`TABLE_STYLE_KEYS` 过滤）**。表格内容文字用 11pt，段前段后 3/3，唯一定义在 `styles.tableText`；`tables.cellText` 只写 `style: tableText` 表示引用，不再重复字符与段落格式。

## 修改样式

1. 改 `configs/thesis-format-rules-graduate.yaml`；
2. 同步 `configs/implementation-map.yaml` 的 `styleMap`/`fallback`；
3. 同步 `src/build_docx.py` 的 `STYLE_NAMES`；
4. 需要恢复继承时改 `src/postprocess_docx.py` 的 `BASE_MAP`；
5. 重建后检查样式是否存在、是否被使用、显示名是否正确，删除未使用样式。

规则中绝不写 `inherits`/`basedOn`/`parent`/`extends`。

## 推荐窗格和样式库

`configs/implementation-map.yaml` 的 `stylePresentation` 管理推荐分组和快速样式库。各变体用格式规则中的样式作为推荐窗格的显示名单。规则外的显式样式和潜在内置样式按原优先级分档，排列在论文样式之后，关闭“使用后显示”。

Word 的有效界面优先级为 1–100，对应 styles.xml 的 0–99。构建为规则外档位保留末尾空间，必要时收紧论文样式靠后的优先级。分组顺序和同组并列关系由实现映射决定。

