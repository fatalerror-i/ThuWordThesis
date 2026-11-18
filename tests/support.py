"""测试公共设施。

分层：
- 静态检查：不启动 Word，检查规则、母版 docx、dotm 包与 Ribbon 回调。
- 宏单元：启动隐藏 Word，打开 dotm 临时副本，调用 VBA 测试入口。
- 导出：在宏单元基础上验证 docx/pdf 产物。

约定：测试只改临时副本，不修改 output/ 下的交付文件；
产物比源码新时直接复用，避免每次跑测试都重建。
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import re
import unittest
import yaml
from pathlib import Path
from zipfile import ZipFile

from docx import Document

try:
    import win32com.client as win32
except ImportError:  # pragma: no cover - 无 Word 环境
    win32 = None

ROOT = Path(__file__).resolve().parents[1]
BUILD_SCRIPT = ROOT / "src" / "build_dotm.py"

def read_part(path: Path, part: str) -> str:
    """读取 docx/dotm 包内某个部件。"""
    with ZipFile(path) as zf:
        return zf.read(part).decode("utf-8")


def style_block(styles_xml: str, name: str) -> str:
    """按样式名取 w:style 片段；找不到返回空串。"""
    pattern = re.compile(
        r'<w:style [^>]*>(?:(?!</w:style>).)*?<w:name w:val="%s".*?</w:style>'
        % re.escape(name),
        re.S,
    )
    found = pattern.search(styles_xml)
    return found.group(0) if found else ""


def style_names(styles_xml: str) -> dict:
    """样式名（小写）到 styleId 的映射。"""
    names = {}
    for match in re.finditer(r"<w:style [^>]*w:styleId=\"([^\"]+)\"[^>]*>(.*?)</w:style>",
                             styles_xml, re.S):
        name = re.search(r'<w:name w:val="([^"]*)"', match.group(2))
        if name:
            names[name.group(1).casefold()] = match.group(1)
    return names


def spacing_attr(style_xml: str, attr: str):
    """取 w:spacing 上某个属性的值。"""
    found = re.search(r"<w:spacing[^/]*w:%s=\"([^\"]*)\"" % attr, style_xml)
    return found.group(1) if found else None


# 两个变体的交付产物；键名与 configs/variants.yaml 一致。
PRODUCTS = {
    "graduate": {
        "label": "研究生",
        "dotm": ROOT / "output" / "清华大学研究生学位论文模板.dotm",
        "docx": ROOT / "output" / "清华大学研究生学位论文模板.docx",
        "pdf": ROOT / "output" / "template_preview-graduate.pdf",
        "ribbon": ROOT / "src" / "dotm" / "customUI14-graduate.xml",
        "rules": ROOT / "configs" / "thesis-format-rules-graduate.yaml",
    },
    "undergraduate": {
        "label": "本科",
        "dotm": ROOT / "output" / "清华大学综合论文训练论文模板.dotm",
        "docx": ROOT / "output" / "清华大学综合论文训练论文模板.docx",
        "pdf": ROOT / "output" / "template_preview-undergraduate.pdf",
        "ribbon": ROOT / "src" / "dotm" / "customUI14-undergraduate.xml",
        "rules": ROOT / "configs" / "thesis-format-rules-undergraduate.yaml",
    },
}
DEFAULT_VARIANT = "graduate"
VARIANTS = tuple(PRODUCTS)

# 每个变体的基线数量，供两变体共用的用例断言。
BASE_COUNTS = {
    "graduate": {
        "sections": 21,
        "tables": 12,
        "content_controls": 17,
        "chapter_headings": 18,
    },
    "undergraduate": {
        "sections": 17,
        "tables": 8,
        "content_controls": 7,
        "chapter_headings": 15,
    },
}

_ready: dict[str, bool] = {}


def product_paths(variant: str = DEFAULT_VARIANT) -> dict:
    """返回变体的产物路径表。"""
    try:
        return PRODUCTS[variant]
    except KeyError:
        raise AssertionError(f"未知变体：{variant}（可选 {sorted(PRODUCTS)}）")


def newest_source_mtime() -> float:
    """src/ 与 configs/ 下最新的文件时间，用来判断产物是否过期。"""
    newest = 0.0
    for folder in ("src", "configs"):
        for path in (ROOT / folder).rglob("*"):
            if path.is_file() and path.suffix in (".py", ".yaml", ".yml", ".xml", ".bas", ".txt"):
                newest = max(newest, path.stat().st_mtime)
    return newest


def _is_fresh(path: Path) -> bool:
    """判断产物是否比最新源码更新。"""
    return path.exists() and path.stat().st_mtime >= newest_source_mtime()


def ensure_dotm(variant: str = DEFAULT_VARIANT) -> None:
    """按需生成 dotm；产物比源码新时直接复用，同一变体只构建一次。"""
    paths = product_paths(variant)
    dotm = paths["dotm"]
    if variant in _ready and dotm.exists():
        return
    if _is_fresh(dotm):
        _ready[variant] = True
        return
    if not paths["docx"].exists():
        raise unittest.SkipTest(f"缺少母版 docx：{paths['docx']}")
    build = subprocess.run(
        [sys.executable, str(BUILD_SCRIPT), "--variant", variant,
         "--dotm-only", "--skip-test", "--keep-existing"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        timeout=600,
    )
    if build.returncode != 0:
        raise AssertionError(
            f"build_dotm.py --variant {variant} 失败：\n"
            f"stdout:\n{build.stdout[-4000:]}\n"
            f"stderr:\n{build.stderr[-4000:]}"
        )
    if not dotm.exists():
        raise AssertionError(f"构建后缺少 dotm：{dotm}")
    _ready[variant] = True


class ProductCase(unittest.TestCase):
    """按变体取产物的测试基类。

    子类只需设 variant，共用断言写一份、两个变体各跑一次；
    variant 只在确有差异的用例里出现。
    """

    variant = DEFAULT_VARIANT

    @classmethod
    def setUpClass(cls) -> None:
        """加载变体产物路径、规则与母版文档。"""
        cls.paths = product_paths(cls.variant)
        cls.docx_path = cls.paths["docx"]
        cls.dotm_path = cls.paths["dotm"]
        cls.ribbon_path = cls.paths["ribbon"]
        cls.rules_path = cls.paths["rules"]
        cls.counts = BASE_COUNTS[cls.variant]
        if not cls.docx_path.exists():
            raise unittest.SkipTest(f"缺少 {cls.variant} 母版 docx：{cls.docx_path}")
        cls.doc = Document(cls.docx_path)
        cls.dotm_styles = (
            read_part(cls.dotm_path, "word/styles.xml") if cls.dotm_path.exists() else ""
        )
        cls.rules = yaml.safe_load(cls.rules_path.read_text(encoding="utf-8"))
        cls.style_map = yaml.safe_load(
            (ROOT / "configs" / "implementation-map.yaml").read_text(encoding="utf-8")
        )["styleMap"]


class WordDotmTest(unittest.TestCase):
    """打开 dotm 临时副本的测试基类；每个用例一份独立文档。

    子类把 variant 设为 "undergraduate" 即可在同一套宏用例上跑本科模板。
    """

    variant = DEFAULT_VARIANT

    @classmethod
    def setUpClass(cls) -> None:
        """启动隐藏 Word 并准备临时目录。"""
        if win32 is None:
            raise unittest.SkipTest("未安装 pywin32，无法运行 Word COM 测试")
        ensure_dotm(cls.variant)
        cls.tmpdir = Path(tempfile.mkdtemp(prefix=f"thu-{cls.variant}-"))
        cls.app = win32.DispatchEx("Word.Application")
        cls.app.Visible = False
        cls.app.DisplayAlerts = 0

    @classmethod
    def tearDownClass(cls) -> None:
        """关闭 Word 并清理临时目录。"""
        try:
            # 显式不保存：Normal.dotm 缺失时 Word 会弹出“是否保存共用模板更改”的模态框，
            # 导致 Quit 被挂起并残留隐藏 WINWORD 进程。
            cls.app.Quit(0)
        except Exception:
            pass
        shutil.rmtree(cls.tmpdir, ignore_errors=True)

    def setUp(self) -> None:
        """为每个用例打开一份 dotm 临时副本。"""
        self.doc_path = self.tmpdir / f"{self._testMethodName}.dotm"
        shutil.copyfile(product_paths(self.variant)["dotm"], self.doc_path)
        self.doc = self.app.Documents.Open(
            str(self.doc_path), ReadOnly=False, AddToRecentFiles=False
        )

    def tearDown(self) -> None:
        """关闭当前用例的文档且不保存。"""
        try:
            self.doc.Close(SaveChanges=False)
        except Exception:
            pass

    def run_macro(self, name: str, *args):
        """调用 VBA 测试入口并把异常转为断言失败。"""
        try:
            return self.app.Run(name, *args)
        except Exception as exc:  # pywintypes.com_error 等
            self.fail(f"运行宏 {name} 失败：{exc!r}")


