#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""构建清华论文宏模板、DOCX 示例与 PDF 预览。

流程：
  1. 生成 DOCX 示例、导出 PDF 预览，再打开同一份 DOCX；
  2. SaveAs2(FileFormat=15) 生成 dotm；
  3. 导入 src/dotm/vba/*.bas（UTF-8 转 GBK）；
  4. 保存并关闭；
  5. 把变体功能区 XML 注入 dotm 包（OPC 部件名固定为 customUI/customUI14.xml）；
  6. 重新打开 dotm，运行 TestSmoke，检查 VBA 与内容控件；
  7. 清洗 docx/dotm 的个人元数据；
  8. 将对应变体的 DOTM、DOCX 和 PDF 写入 output/（摘要打印到终端）。

前置：
  - Windows + Word 2016+；
  - 临时启用“信任对 VBA 工程对象模型的访问”（脚本会保存并在结束后恢复）。
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import yaml
from xml.etree import ElementTree as ET

from lxml import etree
from office_metadata import sanitize_office_metadata

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from build.assemble import assemble_docx
from build.degree_catalog import write_vba_module
from build.style_presentation import finalize_style_presentation
from build.variants import Variant, choices, resolve

# --variant 是必选参数；各变体的产物路径、Ribbon、ThisDocument 由注册表给出。
_VARIANT = None  # 在 main() 中按 --variant 解析后赋值


def _cfg():
    """返回当前变体；未设置时回退到研究生变体（构建入口总会显式设置）。"""
    return _VARIANT or resolve("graduate")


VBA_DIR = ROOT / "src" / "dotm" / "vba"
FORMS_DIR = ROOT / "src" / "dotm" / "forms"
FORMS_SPEC = FORMS_DIR / "forms.json"

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def log(msg: str) -> None:
    """输出带构建脚本前缀的提示信息。"""
    print(f"[build_dotm] {msg}")


def _repo_rel(path: Path) -> str:
    """构建摘要只写仓库相对路径；外部路径只保留文件名。"""
    try:
        return path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        return path.name


class Vbom:
    """临时打开 Word“信任对 VBA 工程对象模型的访问”。"""

    KEY = r"Software\Microsoft\Office\16.0\Word\Security"

    def __init__(self) -> None:
        """记录 VBA 工程访问设置的初始值。"""
        self.original = None
        self.changed = False

    def __enter__(self):
        """临时开启 VBA 工程对象模型访问。"""
        import winreg

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.KEY, 0, winreg.KEY_READ)
            try:
                self.original, _ = winreg.QueryValueEx(key, "AccessVBOM")
            finally:
                winreg.CloseKey(key)
        except FileNotFoundError:
            self.original = None
        if self.original == 1:
            return self
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, self.KEY, 0, winreg.KEY_SET_VALUE)
        try:
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            self.changed = True
        finally:
            winreg.CloseKey(key)
        log("已临时启用 AccessVBOM（结束后恢复）")
        return self

    def __exit__(self, *exc):
        """退出时恢复 VBA 工程访问设置。"""
        if not self.changed:
            return False
        import winreg

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.KEY, 0, winreg.KEY_SET_VALUE)
        try:
            if self.original is None:
                try:
                    winreg.DeleteValue(key, "AccessVBOM")
                except FileNotFoundError:
                    pass
            else:
                winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, int(self.original))
        finally:
            winreg.CloseKey(key)
        log("已恢复 AccessVBOM 原值")
        return False


def gbk_copy(src: Path, dst_dir: Path) -> Path:
    """把 VBA 源文件转成 GBK 编码的 .bas 文件。"""
    text = src.read_text(encoding="utf-8")
    # Word 的 VBA 编辑器用系统 ANSI 代码页；中文 Windows 为 GBK。
    text.encode("gbk")
    dst = dst_dir / (src.stem + ".bas")
    dst.write_text(text, encoding="gbk", newline="\r\n")
    return dst


def _set_form_props(control, props: dict) -> None:
    """逐项设置窗体控件属性，失败时记录提示。"""
    for key, value in props.items():
        try:
            setattr(control, key, value)
        except Exception as exc:
            log(f"  窗体属性未设置 {key}={value!r}：{exc}")


def create_form(vbp, name: str, spec: dict) -> None:
    """在 Word VBE 内直接创建窗体，避免 Excel .frx 控件 blob 的运行时兼容问题。"""
    for comp in list(vbp.VBComponents):
        if comp.Name == name:
            vbp.VBComponents.Remove(comp)
    comp = vbp.VBComponents.Add(3)  # 添加 MSForm 组件（vbext_ct_MSForm）
    comp.Name = name
    try:
        comp.Properties("Caption").Value = spec.get("caption", name)
    except Exception as exc:
        log(f"  窗体标题未设置：{exc}")
    try:
        comp.Properties("Width").Value = spec.get("width", 620)
        comp.Properties("Height").Value = spec.get("height", 438)
    except Exception as exc:
        log(f"  窗体尺寸未设置：{exc}")
    for key, value in (spec.get("props") or {}).items():
        try:
            comp.Properties(key).Value = value
        except Exception as exc:
            log(f"  窗体属性未设置 {key}={value!r}：{exc}")
    designer = comp.Designer
    for item in spec.get("controls", []):
        try:
            control = designer.Controls.Add(item["progid"])
            control.Name = item["name"]
            control.Left = item.get("left", 0)
            control.Top = item.get("top", 0)
            control.Width = item.get("width", 100)
            control.Height = item.get("height", 20)
            _set_form_props(control, item.get("props") or {})
        except Exception as exc:
            log(f"  窗体控件创建失败 {name}.{item.get('name')}：{exc}")
    code_path = FORMS_DIR / f"{name}.code.txt"
    if code_path.exists():
        cm = comp.CodeModule
        cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(code_path.read_text(encoding="utf-8"))


def clean_footnote_paragraphs(path: Path) -> None:
    """清除脚注段落被 Word 带入的直接缩进/段距，使其继承“脚注文本”样式。"""
    w_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(path) as zf:
        items = [(info, zf.read(info.filename)) for info in zf.infolist()]
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for info, data in items:
            if info.filename == "word/footnotes.xml":
                root = etree.fromstring(data)
                for paragraph in root.iter(w_ns + "p"):
                    pPr = paragraph.find(w_ns + "pPr")
                    if pPr is None:
                        continue
                    for tag in ("spacing", "ind"):
                        for element in pPr.findall(w_ns + tag):
                            pPr.remove(element)
                data = etree.tostring(
                    root, encoding="UTF-8", xml_declaration=True, standalone=True
                )
            zf.writestr(info, data)
    tmp.replace(path)


def restore_style_pagination(path: Path) -> None:
    """恢复 Word 保存时从样式中省略的分页控制。"""
    w_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    cfg = _cfg()
    with open(cfg.rules, encoding="utf-8") as fh:
        rules = yaml.safe_load(fh)
    with open(cfg.implementation_map, encoding="utf-8") as fh:
        style_map = yaml.safe_load(fh).get("styleMap", {})

    wanted = {}
    for key, spec in rules["styles"].items():
        pf = spec.get("paragraphFormat") or {}
        name = (style_map.get(key) or {}).get("name")
        flags = {tag: pf[rule] for rule, tag in (
            ("keepWithNext", "keepNext"), ("widowControl", "widowControl"),
        ) if rule in pf}
        if name and flags:
            wanted[name.casefold()] = flags
    if not wanted:
        return

    with zipfile.ZipFile(path) as zf:
        items = [(info, zf.read(info.filename)) for info in zf.infolist()]
    tmp = path.with_suffix(path.suffix + ".tmp")
    changed = 0
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for info, data in items:
            if info.filename == "word/styles.xml":
                root = etree.fromstring(data)
                for style in root.iter(w_ns + "style"):
                    name_el = style.find(w_ns + "name")
                    if name_el is None:
                        continue
                    name = name_el.get(w_ns + "val") or ""
                    if name.casefold() not in wanted:
                        continue
                    pPr = style.find(w_ns + "pPr")
                    if pPr is None:
                        pPr = etree.Element(w_ns + "pPr")
                        anchor_el = None
                        for tag in ("rPr", "tblPr", "trPr", "tcPr", "tblStylePr"):
                            anchor_el = style.find(w_ns + tag)
                            if anchor_el is not None:
                                break
                        if anchor_el is not None:
                            style.insert(list(style).index(anchor_el), pPr)
                        else:
                            style.append(pPr)
                    order = ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl")
                    for tag, enabled in wanted[name.casefold()].items():
                        node = pPr.find(w_ns + tag)
                        if node is None:
                            node = etree.Element(w_ns + tag)
                            preceding = {w_ns + item for item in order[:order.index(tag)]}
                            index = next((i for i, child in enumerate(pPr) if child.tag not in preceding), len(pPr))
                            pPr.insert(index, node)
                            changed += 1
                        if enabled:
                            node.attrib.pop(w_ns + "val", None)
                        else:
                            node.set(w_ns + "val", "0")
                data = etree.tostring(
                    root, encoding="UTF-8", xml_declaration=True, standalone=True
                )
            zf.writestr(info, data)
    tmp.replace(path)
    log(f"按规则恢复样式分页控制：{changed} 处（{', '.join(sorted(wanted))}）")


def inject_custom_ui(dotm_path: Path, xml_path: Path) -> None:
    """把变体功能区 XML 写入 OPC 包，并补关系与内容类型。"""
    xml_data = xml_path.read_bytes()
    with zipfile.ZipFile(dotm_path, "r") as zin:
        names = zin.namelist()
        data = {name: zin.read(name) for name in names}

    data["customUI/customUI14.xml"] = xml_data

    rels = ET.fromstring(data["_rels/.rels"])
    for rel in list(rels):
        if rel.get("Type", "").endswith("/ui/extensibility"):
            rels.remove(rel)
    ET.SubElement(
        rels,
        f"{{{REL_NS}}}Relationship",
        {
            "Id": "rIdCustomUI",
            "Type": "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
            "Target": "customUI/customUI14.xml",
        },
    )
    # Word 对带 ns0: 前缀的包关系/内容类型不识别 customUI，必须用默认命名空间序列化。
    ET.register_namespace("", REL_NS)
    data["_rels/.rels"] = ET.tostring(rels, xml_declaration=True, encoding="UTF-8")

    cts = ET.fromstring(data["[Content_Types].xml"])
    for ov in list(cts):
        if ov.get("PartName") == "/customUI/customUI14.xml":
            cts.remove(ov)
    ET.SubElement(
        cts,
        f"{{{CT_NS}}}Override",
        {
            "PartName": "/customUI/customUI14.xml",
            "ContentType": "application/xml",
        },
    )
    ET.register_namespace("", CT_NS)
    data["[Content_Types].xml"] = ET.tostring(cts, xml_declaration=True, encoding="UTF-8")

    tmp = dotm_path.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, data[name])
        if "customUI/customUI14.xml" not in names:
            zout.writestr("customUI/customUI14.xml", xml_data)
    tmp.replace(dotm_path)


def validate_vba_sources(files: list[Path]) -> bool:
    """导入前做最小语法检查：行首中文必须是注释，避免整模块编译失败。"""
    import re

    errors = []
    for path in files:
        for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            stripped = line.lstrip()
            if stripped and re.match(r"[一-鿿]", stripped):
                errors.append(f"{path.name}:{lineno}: 中文未注释：{line!r}")
    if errors:
        log("VBA 源检查失败：")
        for error in errors:
            log("  " + error)
        return False
    return True


def main() -> int:
    """解析构建变体并执行 dotm 组装与冒烟测试。"""
    ap = argparse.ArgumentParser(description="一次生成 DOTM 模板、DOCX 示例和 PDF 预览。")
    ap.add_argument("--variant", required=True, choices=choices(),
                    help="构建变体：graduate（研究生学位论文）/ undergraduate（本科综合论文训练），必选且无默认值")
    ap.add_argument("--docx", default=None, help=argparse.SUPPRESS)
    ap.add_argument("--dotm-only", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--out", default=None,
                    help="输出 dotm 路径；默认取 configs/variants.yaml 中该变体的 dotm")
    ap.add_argument("--skip-test", action="store_true", help="跳过 Word 冒烟测试")
    ap.add_argument("--keep-existing", action="store_true", help="允许覆盖已有 dotm")
    args = ap.parse_args()

    variant = resolve(args.variant)
    out = (Path(args.out) if args.out else variant.dotm).resolve()
    include_preview = args.docx is None and not args.dotm_only
    if include_preview and out in (variant.docx.resolve(), variant.pdf.resolve()):
        ap.error("DOTM 输出路径不能与 DOCX 或 PDF 相同")
    with tempfile.TemporaryDirectory(prefix="thu_dotm_") as directory:
        return build(args, variant, out, Path(directory), include_preview)


def build(args: argparse.Namespace, variant: Variant, out: Path,
          tmpdir: Path, include_preview: bool) -> int:
    global _VARIANT
    _VARIANT = variant
    ribbon = variant.ribbon
    thisdocument = variant.this_document
    log(f"变体：{variant.name}（{variant.label}）")
    release_version = (ROOT / "configs" / "release-version.txt").read_text(encoding="utf-8").strip()
    if not release_version or len(release_version.splitlines()) != 1:
        log("configs/release-version.txt 应填写一行发布版本号")
        return 2

    docx = Path(args.docx).resolve() if args.docx else tmpdir / variant.docx.name
    pdf = tmpdir / variant.pdf.name if include_preview else None
    if args.docx is None:
        log("组装 DOCX 并导出预览" if include_preview else "组装临时 DOCX")
        assemble_docx(docx, variant.name, pdf)
    if not docx.exists():
        log(f"缺少母版 docx：{docx}")
        return 2

    import win32com.client as win32

    degree_bas = write_vba_module(tmpdir / "modDegreeOptions.bas")
    continued = yaml.safe_load(variant.rules.read_text(encoding="utf-8"))["captions"]["continuedTable"]
    prefix = continued["prefix"] + "表"
    suffix = continued.get("suffix", "")
    app_title = variant.title.replace('"', '""')
    app_version = release_version.replace('"', '""')
    constants = (
        f'Public Const THU_UNDERGRADUATE As Boolean = {variant.name == "undergraduate"}\n'
        f'Public Const APP_TITLE As String = "{app_title}"\n'
        f'Public Const APP_VERSION As String = "{app_version}"\n'
        f'Public Const THU_CONTINUED_PREFIX As String = "{prefix}"\n'
        f'Public Const THU_CONTINUED_SUFFIX As String = "{suffix}"\n'
    )
    source = degree_bas.read_text(encoding="utf-8")
    degree_bas.write_text(source.replace("Option Explicit\n", "Option Explicit\n" + constants, 1),
                          encoding="utf-8")
    bas_files = sorted(VBA_DIR.glob("*.bas")) + [degree_bas]
    if not bas_files:
        log(f"缺少 VBA 源：{VBA_DIR}")
        return 2
    if not validate_vba_sources(bas_files):
        return 2
    if not FORMS_SPEC.exists():
        log(f"缺少窗体定义：{FORMS_SPEC}")
        return 2
    form_specs = json.loads(FORMS_SPEC.read_text(encoding="utf-8"))
    if not ribbon.exists():
        log(f"缺少 Ribbon XML：{ribbon}")
        return 2
    if not thisdocument.exists():
        log(f"缺少 ThisDocument 事件源：{thisdocument}")
        return 2

    word = None
    doc = None
    report = {
        "dotm": _repo_rel(out),
        "docx": _repo_rel(docx),
        "modules": [],
        "forms": [],
        "smoke": None,
        "thisdocument_lines": 0,
    }
    try:
        dotm = tmpdir / "dotm" / out.name if include_preview else out
        if dotm.exists() and not args.keep_existing:
            dotm.unlink()
        dotm.parent.mkdir(parents=True, exist_ok=True)
        gbk_dir = tmpdir / "gbk"
        gbk_dir.mkdir()
        gbk_files = [gbk_copy(p, gbk_dir) for p in bas_files]

        with Vbom():
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(docx), ReadOnly=False, AddToRecentFiles=False)
            log(f"已打开 docx：{docx.name}")

            # 1) 另存为启用宏的模板
            doc.SaveAs2(FileName=str(dotm), FileFormat=15)
            log(f"已生成 dotm：{dotm.name}")

            # 2) 导入 VBA 模块（先移除同名标准模块，保留 ThisDocument）
            vbp = doc.VBProject
            existing = {c.Name for c in vbp.VBComponents}
            for path in gbk_files:
                name = path.stem
                if name in existing and name != "ThisDocument":
                    vbp.VBComponents.Remove(vbp.VBComponents(name))
                vbp.VBComponents.Import(str(path))
                report["modules"].append(name)
                log(f"已导入模块：{name}")

            # 2b) 写入 ThisDocument 的事件代码（内容控件级联入口）
            thisdoc = vbp.VBComponents("ThisDocument").CodeModule
            thisdoc_code = thisdocument.read_text(encoding="utf-8")
            if thisdoc.CountOfLines:
                thisdoc.DeleteLines(1, thisdoc.CountOfLines)
            thisdoc.AddFromString(thisdoc_code)
            report["thisdocument_lines"] = int(thisdoc.CountOfLines)
            log(f"已写入 ThisDocument 事件代码：{report['thisdocument_lines']} 行")

            for name, spec in form_specs.items():
                create_form(vbp, name, spec)
                report["forms"].append(name)
                log(f"已创建窗体：{name}")

            doc.Save()
            doc.Close(SaveChanges=0)
            doc = None
            word.Quit()
            word = None

        # 3) 注入 Ribbon（在 Word 关闭后做，避免 Word 保存时重写包）
        inject_custom_ui(dotm, ribbon)
        log("已注入变体功能区 XML")

        # 4) 恢复脚注和样式格式，清理交付包后再运行冒烟测试。
        clean_footnote_paragraphs(dotm)
        restore_style_pagination(dotm)
        finalize_style_presentation(dotm, variant)
        sanitize_office_metadata(dotm)
        log("已清理 dotm 的文档属性和本机 VBA 缓存")

        # 5) 用交付文件验证宏和窗体引用。
        if not args.skip_test:
            with Vbom():
                word = win32.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                doc = word.Documents.Open(str(dotm), ReadOnly=True, AddToRecentFiles=False)
                smoke = word.Run("TestSmoke")
                report["smoke"] = str(smoke)
                log(f"TestSmoke：{smoke}")
                missing = [ref.Name for ref in doc.VBProject.References if ref.IsBroken]
                if missing:
                    raise RuntimeError(f"VBA 引用无法加载：{', '.join(missing)}")
                for name in form_specs:
                    if doc.VBProject.VBComponents(name).Designer.Controls.Count == 0:
                        raise RuntimeError(f"VBA 窗体没有加载控件：{name}")
                report["vb_components"] = [c.Name for c in doc.VBProject.VBComponents]
                report["content_controls"] = int(doc.ContentControls.Count)
                report["sections"] = int(doc.Sections.Count)
                doc.Close(SaveChanges=0)
                doc = None
                word.Quit()
                word = None

        with zipfile.ZipFile(dotm) as z:
            names = set(z.namelist())
            report["has_vba"] = "word/vbaProject.bin" in names
            report["has_custom_ui"] = "customUI/customUI14.xml" in names
            report["size"] = dotm.stat().st_size
        if not report["has_vba"]:
            log("警告：dotm 中没有 vbaProject.bin")
        if not report["has_custom_ui"]:
            log("警告：dotm 中没有 customUI/customUI14.xml")
        report["ok"] = bool(
            report.get("has_vba")
            and report.get("has_custom_ui")
        )
        if report["smoke"] and not str(report["smoke"]).startswith("OK"):
            report["ok"] = False
        if report["ok"] and pdf is not None:
            for source, target in ((docx, variant.docx), (pdf, variant.pdf), (dotm, out)):
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(source, target)
                log(f"saved: {_repo_rel(target)}")
            report["docx"] = _repo_rel(variant.docx)
            report["pdf"] = _repo_rel(variant.pdf)
        log(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if report["ok"] else 1
    finally:
        for obj, attr in ((doc, "Close"), (word, "Quit")):
            if obj is not None:
                try:
                    getattr(obj, attr)()
                except Exception:
                    pass


if __name__ == "__main__":
    sys.exit(main())
