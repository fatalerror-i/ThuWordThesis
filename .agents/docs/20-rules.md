# 格式规则模型与修改流程

## 唯一事实来源

研究生用 `configs/thesis-format-rules-graduate.yaml`，本科用 `configs/thesis-format-rules-undergraduate.yaml`；两份都描述“最终长什么样”，不描述 Word 的样式 ID、父类或脚本实现。改规则时先确认改哪个变体，两边必须保持样式键与实现映射一致。

- 键、值用英文；中文只出现在注释、编号文本、题注字段名等白名单。
- 数值与单位分开：`{value: 2, unit: character}`。
- 行距：`{rule: single|oneAndHalf|double|multiple|minimum|fixed, amount: {...}}`。
- 只写显式设置；没有写的字段由实现层 fallback，不在规则里写 `inherits`。
- 拓扑顺序：`metadata`、`conventions`、`pageSetup`、`styles`、`numbering`、`captions`、`tables`、`footnotes`、`headersAndFooters`、`cover`、`documentParts`。
- `tables` 段只写表格层面的约定，不重复 `styles` 里的格式：表格单元格文字的字符与段落格式只在 `styles.tableText` 定义，`tables.cellText` 只写 `style: tableText` 表示引用。`validate_rules.py` 会检查 `tables.default`、`captionStyle`、`noteStyle` 以及各表格项的 `style` 都指向真实存在的样式键。

## 样式键命名

- lowerCamelCase、英文、表意明确，不用 Word XML 缩写。
- 每个样式注释至少包含：中文名、对应 Word 样式名/用途、指南条款（有则写）。
- 不要新增未使用的样式；删除样式时同步删除规则索引、`implementation-map.yaml`、`STYLE_NAMES`、`BASE_MAP` 和脚本引用。

## 改一个格式值

1. 在 YAML 找到对应样式/页面设置；
2. 保持 `{value, unit}` 和行距结构；
3. `python src/validate_rules.py --all`（校验两变体）；
4. `python src/build_dotm.py --variant <改动所属变体>`；
5. 检查对应变体的 dotm 冒烟结果与 PDF 对应页、页数、目录、页眉页码没有意外变化。

## 新增或调整样式

1. 在 `styles` 下新增语义键；
2. 写中文注释和指南依据；
3. 在 `configs/implementation-map.yaml` 的 `styleMap` 增加样式名与 `fallback`（不写 `styleId`，理由见 `50-build.md` 的「规则到 Word 的映射」）；
4. 在 `src/build_docx.py` 的 `STYLE_NAMES` 增加映射；
5. 涉及编号时更新 `numberingApply` 与 `build_numbering()`；
6. 重建并检查样式是否实际存在、是否被使用。

## 规则边界

- 官方指南与规则冲突：规则优先；若认为规则有误，先向用户汇报，不要自行改规则。
- 规则没写的格式：不要凭感觉补；先在回复中列出待定项。
- 用户明确说“按规则”时，逐条核对 YAML；例如页眉下框线在 `styles.header.paragraphFormat.borderBottom`（研究生 `single 0.75pt spacing 1pt`；本科无此键，即页眉不带框线）。`headersAndFooters.header` 只描述页眉内容与所用样式，不写边框。
- 规则文件不写继承；`basedOn`、fallback、Word 内置样式 ID 只在实现层。