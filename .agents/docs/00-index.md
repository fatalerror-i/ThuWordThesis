# 渐进式指令索引

本目录保存拆分后的维护指南。**先读本文件，再按任务只读相关分册**，不要把全部文件一次读入。

## 必读顺序

1. `.agents/AGENTS.md`：硬约束、目录边界与路由。
2. 本文件：任务路由与关键事实。
3. 只读与当前任务相关的 1–3 个分册。

## 文档分工

| 入口 | 放在这里的内容 |
| --- | --- |
| [README](../../README.md) | 项目简介、模板选择、文件链接和阅读入口 |
| 模板示例（`src/build/content/`） | 新建、排版和定稿的简短提示，以及可在 Word 中查看的编号、样式和引用示例 |
| [用户文档](../../docs/01-快速入门.md) | 按钮操作、排版规则、常见故障和打印方式 |
| 本目录 | 源码入口、构建步骤、已知的 Word 行为和开发验证 |

## 任务路由

| 当前任务 | 读取 |
| --- | --- |
| 第一次接触仓库、目录、命令、维护流程 | `10-workflow.md` |
| 改格式值、增删样式、改页面设置 | `20-rules.md` |
| 改封面、分节、目录、页眉页脚、页码、域 | `30-docx-structure.md` |
| 改样式名、继承关系、内置样式、列表样式 | `40-styles.md` |
| 改生成器/后处理/构建流程 | `50-build.md` |
| 重建后检查 | `60-validation.md` |
| 遇到 Word 行为异常 | `70-pitfalls.md` |
| 开发/维护宏模板 `.dotm`、VBA 模块、关键实现点 | `80-dotm.md` |
| 查功能区按钮图标与 `imageMso` 对照表 | `85-ribbon-icons.md` |
| **下次会话交接、当前状态** | `90-handoff.md` |
| 规则覆盖情况 | `95-rule-coverage.md` |
| 查与官方写作指南的差异与处理方法 | [格式定义](../../docs/02-格式定义.md#与写作指南的差异) |

## 优先级

冲突时按以下顺序：

1. 用户在当前会话中的最新明确指令；
2. 当前变体的 `configs/thesis-format-rules-graduate.yaml` 或 `configs/thesis-format-rules-undergraduate.yaml` 中明确写出的值；
3. 当前变体对应的研究生写作指南或本科综合论文训练写作规范，见 `references/guide-graduate/` 或 `references/guide-undergraduate/`；
4. `src/` 中的既有实现；
5. `references/` 下的旧模板、官方示例、LaTeX 对照材料。

规则未规定、但实现必须决定的事项：**先向用户汇报并确认，不要默默设计**。

## 关键事实

具体数字（样式数、测试数、页数、文件大小、功能区组数）以现行 `configs/*.yaml`、`src/dotm/customUI14*.xml` 和 `output/` 产物为准，本文件不写死，避免过期。

- 同仓双变体：`configs/variants.yaml` 注册 `graduate`（研究生学位论文）与 `undergraduate`（本科综合论文训练）；两份规则文件分别是 `configs/thesis-format-rules-graduate.yaml` 与 `configs/thesis-format-rules-undergraduate.yaml`，二者共享 `configs/implementation-map.yaml` 与学位配置。
- 主交付是两份 dotm：`output/清华大学研究生学位论文模板.dotm` 与 `output/清华大学综合论文训练论文模板.dotm`。docx/PDF 只用于构建和预览。构建摘要打印在终端，不落盘。
- `build_dotm.py` 每次生成 DOTM、DOCX 示例和 PDF 预览。用 `--variant` 选择 `graduate` 或 `undergraduate`。
- 示例正文改 `src/build/content/examples.py`，附录改 `src/build/content/appendix.py`。摘要、封面及专属前后置章节分别在变体目录下维护。
- 校验入口：`python src/validate_rules.py --all`（分别校验两变体，无参数时默认研究生）、`python -m unittest discover -s tests -v`。
- 图标入口：映射 `src/dotm/icons.yaml`、对照表 `85-ribbon-icons.md`、候选池 `references/icons/imageMso-pool.yaml`、图库 `references/icons/imageMso-pool/gallery.html`；校验与导出脚本 `references/icons/check_icons.py` 等。
- 图标校验与预览导出脚本在 `references/icons/`，与候选池、图库同目录；不被构建链导入，命令见 `10-workflow.md`。
- 规则文件不能出现 `inherits`/`basedOn`/`parent`/`extends`；继承只在 `configs/implementation-map.yaml` 和脚本里实现。
- 后处理脚本 `src/postprocess_docx.py` 的 `BASE_MAP` 与 `configs/implementation-map.yaml` 的 `fallback` 必须一致，两处都描述样式继承。
- 六个功能区列表样式（编号/项目符号 1–3）写 2 字符悬挂缩进；一级不设左缩进，二级、三级分别设 2、4 字符左缩进，以表示多级列表。
- 当前交接与未完成事项见 `90-handoff.md`。

## 目录速查

```text
.agents/     规则、方案、提示词、维护指南（本目录）
docs/        快速入门、格式定义（含指南差异）、常见问题
src/         构建 Python 与 dotm 源码
tests/       回归测试
configs/      规则与实现映射
output/      交付 docx、dotm、预览 PDF、检查产物
assets/      封面校徽/校名矢量图
references/  参考文档（官方指南、旧模板、官方示例、thuthesis）
```

详细职责见 `10-workflow.md`。
