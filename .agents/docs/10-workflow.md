# 目录职责、环境与标准流程

## 目录职责

```text
ThuWordThesis/
├─ .agents/
│  ├─ AGENTS.md                  入口：硬约束、目录边界、命令
│  └─ docs/                      渐进式维护指南（本目录）
├─ README.md                     给读者看的项目说明
├─ TODO.md                       维护待办
├─ configs/
│  ├─ thesis-format-rules-graduate.yaml   格式唯一事实来源
│  └─ implementation-map.yaml    规则键↔Word 样式/父类/编号
├─ src/
│  ├─ build_docx.py            从规则生成 docx（样式/页面/节骨架）
│  ├─ finalize_docx.py         生成→更新域→后处理→导出 PDF（预览用）
│  ├─ build_dotm.py            构建 DOTM、DOCX 和 PDF 的统一入口
│  ├─ postprocess_docx.py      Word 更新域后的规范化
│  ├─ validate_rules.py        规则结构校验
│  ├─ build/                   assemble/degree_catalog/variants 等
│  ├─ build/content/           cover/front_matter/examples/back_matter
│  └─ dotm/
│     ├─ customUI14-graduate.xml          11 组功能区
│     ├─ icons.yaml              28 个直接控件图标和 2 个导出选项图标
│     ├─ forms/                  forms.json 与窗体代码片段
│     └─ vba/                    9 个 VBA 标准模块（含生成模块）
├─ tests/                        静态格式/样式契约与 VBA 宏单元测试
├─ output/
│  ├─ 清华大学研究生学位论文模板.dotm  主交付
│  ├─ 清华大学研究生学位论文模板.docx  构建/预览中间物
│  ├─ template_preview-graduate.pdf
│  └─ checks/                    检查临时产物
├─ docs/                        面向使用者的文档（编号连续）
│  ├─ 01-快速入门.md             安装、功能区、各类对象怎么插入
│  ├─ 02-格式定义.md             两版排版设置、与指南的差异及核对方法
│  └─ 03-常见问题.md             报错、异常现象与解决办法
├─ assets/                       封面校徽/校名矢量图
├─ references/
│  ├─ guide-graduate/            研究生院官方指南 Markdown/PDF
│  ├─ guide-undergraduate/       本科写作规范
│  ├─ official-samples-graduate/ 研究生官方示例
│  ├─ official-samples-undergraduate/ 本科官方示例
│  ├─ icons/                      imageMso 候选池、预览图库与图标工具脚本
│  │  ├─ imageMso-pool.yaml       1,494 个有效 imageMso 候选（参考数据）
│  │  ├─ imageMso-pool/           候选池预览图库（本机生成，不分发）
│  │  ├─ check_icons.py           校验 imageMso ID 是否有效
│  │  ├─ export_control_icons.py  导出已选控件图标预览
│  │  └─ export_icon_pool.py      导出候选池图库
│  └─ thuthesis/                 LaTeX 参考
```

## 环境

- Windows 10/11，安装 Microsoft Word（COM 自动化；本机无 WPS）。
- Python 3.11+，依赖见 `requirements.txt`。`pyOpenVBA` 和 `pypdf` 分别用于清理宏项目和预览 PDF 的元数据。示例插图用 Word 原生图形占位，图标导出工具只用标准库。
- PDF 转图片可用本机已安装的 `pdftoppm`（按环境配置路径，不写入仓库）。

## 标准命令

```powershell
# 1) 校验两变体规则
python src/validate_rules.py --all
# 期望：OK: rules structure valid [graduate] / [undergraduate]

# 2) 每个变体一次生成 DOTM、DOCX 和 PDF，并检查 DOTM
python src/build_dotm.py --variant graduate
python src/build_dotm.py --variant undergraduate
# 期望：两次 TestSmoke 均为 OK，每个变体各写入三份文件

# 3) 全量测试
python -m unittest discover -s tests -v
```

```powershell
# 图标辅助命令（脚本在 references/icons/）
python references/icons/check_icons.py
python references/icons/export_icon_pool.py
```

`build_dotm.py` 每次只生成一份 DOCX，随后从该文件导出 PDF、转换 DOTM。`--variant` 必填，选项来自 `configs/variants.yaml`。`validate_rules.py` 无参数时校验研究生规则。

构建产物：

- `output/清华大学研究生学位论文模板.docx` / `output/清华大学综合论文训练论文模板.docx`
- `output/清华大学研究生学位论文模板.dotm` / `output/清华大学综合论文训练论文模板.dotm`（主交付）
- `output/template_preview-graduate.pdf` / `output/template_preview-undergraduate.pdf`（由 `finalize_docx.py` 导出）

## 构建关系

1. `src/build_dotm.py` 按 `--variant` 读取配置。`src/build/assemble.py` 将变体名和临时 DOCX 路径传给 `src/finalize_docx.py`，由它调用 `src/build_docx.py` 生成 DOCX。
2. `src/finalize_docx.py` 用 Word 更新域并完成后处理，从该 DOCX 导出 PDF。
3. `src/build_dotm.py` 对同一份 DOCX 执行 `SaveAs2(FileFormat=15)`，导入宏、创建窗体、写入功能区和文档事件代码。
4. DOTM 冒烟测试通过后，将三份文件写入 `output/`。测试工具仅需重建 DOTM 时使用内部参数 `--dotm-only`。

## 为什么两轮更新域

`finalize_docx.py` 的顺序：

1. `build_docx.py` 生成 docx；
2. `word_update_fields()` 第一轮更新 `SEQ`/`STYLEREF`、目录、图表清单、页眉页脚域；
3. `postprocess_docx.py` 修编号、样式、页眉边框、目录字体；
4. `word_update_fields()` 第二轮；
5. 再跑一次 `postprocess_docx.py`；
6. 清理 DOCX 文档属性，再由 `word_export_pdf()` 导出预览并清理 PDF 元数据。

后处理可能改变编号或样式，第一次算出的目录会缓存旧章号；必须再让 Word 重算一次。

## 改一件事要动哪些文件

| 改动 | 需要同步的位置 |
| --- | --- |
| 改某个格式数值 | 该变体的规则 YAML → `python src/validate_rules.py --all` → 重建 dotm |
| 新增/删除样式键 | 规则 YAML 的 `styles` 块 + 文件内 `# 样式索引` 注释块 + `configs/implementation-map.yaml` 的 `styleMap`（三者必须一一对应，校验器会报错） |
| 改样式继承 | 只改 `configs/implementation-map.yaml` 的 `basedOn`/`fallback`；规则文件禁止写 `inherits`/`basedOn`/`parent`/`extends` |
| 改后处理行为 | `src/postprocess_docx.py`；注意它有两处样式继承来源（`BASE_MAP` 与 `implementation-map.yaml` 的 `fallback`），两者必须一致 |
| 改宏或功能区 | `src/dotm/vba/*.bas`、`src/dotm/vba/ThisDocument*.txt`、`src/dotm/forms/*`、`src/dotm/customUI14*.xml`、`src/dotm/icons.yaml`。续表生成入口已注释，规则前后缀只保留为旧稿识别兼容数据 |
| 改内容控件 tag | VBA、`src/dotm/vba/ThisDocument-graduate.txt`、模板说明页、`80-dotm.md` |
| 改示例文字 | 共用正文改 `src/build/content/examples.py`，共用附录改 `appendix.py`，封面、摘要及专属前后置章节改相应变体目录。同步重建 DOCX、DOTM 和 PDF |
| 改读者文档 | `README.md` 与 `docs/`；改维护文档同步 `.agents/docs/` 与 `.agents/AGENTS.md` 路由 |

## 维护流程

1. 改格式：先改对应变体的规则文件（`configs/thesis-format-rules-graduate.yaml` / `configs/thesis-format-rules-undergraduate.yaml`）或 `implementation-map.yaml`，再改 `src/build_docx.py` 或 `src/postprocess_docx.py`。
2. 改交互：改 `src/dotm/vba/*.bas`、`src/dotm/vba/ThisDocument*.txt`、`src/dotm/forms/*`、`src/dotm/customUI14*.xml`。
3. 运行 `python src/validate_rules.py --all`，再按上面的标准命令构建 DOTM 及配套 DOCX、PDF。
4. 验证：跑全量测试；检查两个变体的 docx 页数/节数、PDF、dotm 冒烟。
5. 不直接编辑 `output/` 中的 docx/dotm；不把本机路径、账户名或真实个人信息写入源码、文档、测试或构建日志。

## 变更纪律

- 改格式：先改 `configs/thesis-format-rules-graduate.yaml`，再校验，再重建；不要手改输出 docx。
- 改样式/编号：同步 `configs/implementation-map.yaml`、`src/build_docx.py`、`src/postprocess_docx.py`。
- 改路径：所有脚本从自身位置推导 `ROOT`，不要新增绝对路径。
- 每次交付保留：输出 docx、dotm、预览 PDF、规则、映射、源码；试验产物放 `output/checks/`（已在 `.gitignore` 中忽略）。
- 规则未覆盖的决定先问用户，不要用“看起来更好”替代明确规则。

## 敏感信息与隐私

- **路径**：源码、文档、测试和构建输出中不写盘符路径、用户目录、临时目录或具体 TeX/Office 安装路径。文档用“仓库根目录”或相对路径；代码用 `ROOT` 或 `Path(__file__)` 派生。
- **凭据**：不写入密码、API key、token、cookie、私钥、连接串或 `.env` 内容。需要配置的密钥只从环境变量读取，不提交样例真值。
- **真实个人信息**：不写入真实姓名、学号、专业/院系、单位、电话、邮箱、身份证号、地址、照片、签名等。封面预填以 `references/thuthesis/thusetup.tex` 的公开作者示例为限，其他示例及测试数据使用虚构占位。
- **生成物元数据**：构建时清除 DOCX/DOTM 的作者、最后修改者、软件信息和构建时间。DOTM 同时清除 VBA 库的本机路径和编译缓存。预览 PDF 清除文档信息和 XMP。
- **例外**：`references/guide-graduate/` 的官方指南、`references/` 的第三方模板/文献属于外部材料，其中公开的引用作者、版权署名和链接按原许可证保留；不要把这些材料扩展成个人资料，也不要把它们当作本项目收集的个人信息。
- **提交前检查**：扫描文本文件中的盘符路径、用户目录、邮箱、电话、证件号、`api_key`/`password`/`token` 等关键字。运行交付物元数据测试，检查两种变体的 DOCX、DOTM 和 PDF。

## Git 与许可

- `.gitignore` 忽略 Python 缓存、Word 锁文件、临时检查产物。
- `.gitattributes` 将 docx、dotm、doc、pdf 和图片标记为二进制，避免错误换行转换。
- [LICENSE](../../LICENSE) 是 CC BY-NC 4.0，适用于本项目的原创源码、配置、文档、模板和预览。
- 校徽、校名、官方参考材料及 ThuThesis 示例的来源见 [README](../../README.md#许可与第三方素材)。分发第三方素材时保留原权利声明和适用许可。
