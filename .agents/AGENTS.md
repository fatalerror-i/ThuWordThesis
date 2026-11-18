# AGENTS.md

本仓库维护清华大学学位论文 Word 模板（`.docx`）与带宏模板（`.dotm`），同仓双变体：`graduate`（研究生学位论文）与 `undergraduate`（本科综合论文训练）。面向写作者与使用者的说明在 `README.md` 和 `docs/`；本文件与 `.agents/docs/` 面向维护者和 AI agent。

## 先做什么

1. 读本文件，确认硬约束与目录边界。
2. 读 `.agents/docs/00-index.md` 的任务路由。
3. 只读与当前任务相关的 1–3 个分册，不要一次加载全部。
4. 动手前确认要改的是规则、实现、测试，还是输出。

## 六条硬约束

1. **规则是唯一事实来源**：`configs/thesis-format-rules-graduate.yaml`（研究生）与 `configs/thesis-format-rules-undergraduate.yaml`（本科）。规则明确写出的值必须照做；规则未规定、但实现必须决定的事项，先向用户汇报，不要默默设计。
2. **不要手改生成物**：`output/清华大学研究生学位论文模板.dotm` 与 `output/清华大学综合论文训练论文模板.dotm` 是构建产物。格式、宏、功能区改动都回到规则、`src/` 源码和构建脚本。
3. **最新用户指令优先**：与历史指南、脚本或文档冲突时，以用户当前明确要求为准，并同步更新规则、实现、测试和文档。
4. **规则不写继承**：`inherits` / `basedOn` / `parent` / `extends` 只存在于 `configs/implementation-map.yaml` 和脚本。
5. **规则键/值用英文，注释用中文**；改完必须跑规则校验、重建和测试（命令见 `10-workflow.md`）。
6. **不写入本地路径、凭据和真实个人信息**：仓库中的源码、文档、测试、构建输出和交付物都不得包含本机绝对路径、本地账户名、密码、API key、token、私钥，也不得包含真实人物的姓名、学号、专业/院系、单位、联系方式、身份证号、照片或签名。封面采用指定的 ThuThesis 公开作者示例，其他示例和测试数据使用虚构占位信息。`references/` 中第三方文献、许可证和版权声明里的公开署名按原许可证保留。检查细则见 `10-workflow.md`。

## 目录边界

```text
.agents/       规则、方案、提示词、维护指南（本文件所在目录）
docs/          快速入门、格式定义（含指南差异）、常见问题
src/           构建 Python 与 dotm 源码（build/、build/content/、dotm/）
tests/         静态格式/样式契约与 VBA 宏单元测试
configs/        variants.yaml + 两份格式规则 + 实现映射 + 学位配置
output/        交付 docx、dotm、预览 PDF、检查产物
assets/        封面校徽/校名矢量图
references/    参考文档与参考数据（官方指南、官方示例、thuthesis、imageMso 候选池、图标图库与图标工具脚本；只读，不作为当前规则依据）
```

目录规则：

- 读者文档只放 `docs/` 和 `README.md`；agent 规则、方案、提示词只放 `.agents/`。
- 构建链 Python 与 VBA 源码只放 `src/`；图标校验与预览导出脚本放 `references/icons/`；不要把宏写进 `output/` 的 dotm 再当作源码。
- 测试只放 `tests/`；测试不得修改 `output/` 交付物。
- 生成物只放 `output/`；检查临时产物放 `output/checks/`；图标候选池、本机预览图库与图标工具脚本放 `references/icons/`（Office 图标版权归 Microsoft，不随仓库分发）。
- 图片放在 `assets/`，参考文档放在 `references/`。第三方素材保留原权利归属和许可，详见 [README](../README.md#许可与第三方素材)。
- 所有脚本从自身位置推导仓库根 `ROOT`，不要写死盘符或绝对路径。
- 在 Git Bash/MSYS 下禁止用 `>nul` 重定向：`nul` 是 Windows 保留设备名，MSYS 会把它写成仓库里的普通文件；清空输出请用 `>/dev/null` 或 `$null`。不要创建 `nul`、`con`、`aux`、`prn`、`com1`–`com9`、`lpt1`–`lpt9` 等保留名文件。

## 路由

| 任务 | 文件 |
| --- | --- |
| 目录、环境、标准命令、维护流程、改一件事要动哪些文件 | `.agents/docs/10-workflow.md` |
| 格式规则模型、改规则 | `.agents/docs/20-rules.md` |
| 节、封面、目录、页眉页脚、域 | `.agents/docs/30-docx-structure.md` |
| 内置样式、继承、题注、列表 | `.agents/docs/40-styles.md` |
| 生成器、后处理、构建管线 | `.agents/docs/50-build.md` |
| 重建后的检查 | `.agents/docs/60-validation.md` |
| Word 行为坑 | `.agents/docs/70-pitfalls.md` |
| `.dotm` 宏模板、VBA 模块、关键实现点 | `.agents/docs/80-dotm.md` |
| 功能区按钮图标 / `imageMso` 对照表 | `.agents/docs/85-ribbon-icons.md` |
| **下次会话交接** | `.agents/docs/90-handoff.md` |
| 规则覆盖情况 | `.agents/docs/95-rule-coverage.md` |
| 与官方写作指南的差异及处理方法 | [格式定义](../docs/02-格式定义.md#与写作指南的差异) |

命令、构建关系、隐私检查细则、Git 与许可说明都在 `10-workflow.md`。
