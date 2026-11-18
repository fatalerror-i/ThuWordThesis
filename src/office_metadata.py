# -*- coding: utf-8 -*-
"""清理构建产物中的账户、路径和构建环境信息。"""

from __future__ import annotations

import re
import struct
import sys
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from lxml import etree

PACKAGE_DATE = (1980, 1, 1, 0, 0, 0)
PUBLIC_PROPERTIES = {
    "docProps/core.xml": {"title", "subject"},
    "docProps/app.xml": {
        "Pages", "Words", "Characters", "CharactersWithSpaces", "Lines", "Paragraphs",
    },
}
VBA_PART = "word/vbaProject.bin"
# MS-OVBA 2.1.1.8 中以 GUID 标识的库引用。
LIBID = re.compile(
    rb"\*\\G\{[0-9a-f-]{36}\}#[0-9a-f]+\.[0-9a-f]+#[0-9a-f]+#([^#\x00]*)#",
    re.IGNORECASE,
)


def _remove_library_paths(references: bytes) -> bytes:
    """清空 LibidPath，保留类型库 GUID、版本、语言和显示名。"""
    data = bytearray(references)
    for match in reversed(list(LIBID.finditer(references))):
        start, end = match.span(1)
        if start == end:
            continue
        libid = match.start()
        # MS-OVBA 2.3.4.2.2.3–5：原始引用只记录 Libid 长度，注册库
        # 和控件引用还需更新外围记录长度。
        if references[libid - 6:libid - 4] == b"\x33\x00":
            lengths = (libid - 4,)
        elif references[libid - 10:libid - 8] in {b"\x0d\x00", b"\x2f\x00", b"\x30\x00"}:
            lengths = (libid - 4, libid - 8)
        else:
            raise ValueError("无法清理未知结构的 VBA 类型库引用")
        for offset in lengths:
            size, = struct.unpack_from("<I", data, offset)
            struct.pack_into("<I", data, offset, size - (end - start))
        del data[start:end]
    return bytes(data)


def _sanitize_vba(data: bytes) -> bytes:
    from pyopenvba.cfb import CFB
    from pyopenvba.vba import compress, parse_vba_project, serialize_dir_modules_section

    cfb = CFB.from_bytes(data)
    project = parse_vba_project(cfb)
    prefix = _remove_library_paths(project.dir_raw[:project.dir_modules_offset])
    # 丢弃 VBE 缓存，让 Word 从源代码编译，并用本机注册信息解析库。
    for module in project.modules:
        stream = cfb.get_stream_in_storage("VBA", module.stream_name)
        cfb.write_stream_in_storage("VBA", module.stream_name, stream[module.text_offset:])
        module.text_offset = 0
    cfb.write_stream_in_storage("VBA", "dir", compress(
        prefix + serialize_dir_modules_section(project),
    ))
    # MS-OVBA 2.3.4.1：标准 7 字节头部，版本置为 0xFFFF。
    cfb.write_stream_in_storage("VBA", "_VBA_PROJECT", b"\xcc\x61\xff\xff\x00\x01\x00")
    for name in cfb.list_streams_in_storage("VBA"):
        if name.startswith("__SRP_"):
            cfb.remove_stream_in_storage("VBA", name)
    # CFB 库保留目录项原始字节，[MS-CFB] 2.6.1 的两个时间戳位于 100–115。
    for entry in cfb._directory:
        if entry.raw:
            entry.raw = entry.raw[:100] + bytes(16) + entry.raw[116:]
    return cfb.to_bytes()


def _keep_public_properties(data: bytes, names: set[str]) -> bytes:
    root = etree.fromstring(data)
    removed = False
    for element in list(root):
        if etree.QName(element).localname not in names:
            root.remove(element)
            removed = True
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True) if removed else data


def sanitize_office_metadata(path: str | Path) -> bool:
    """清理生成的 Word 包，返回文件是否有改动。"""
    path = Path(path)
    if path.suffix.lower() not in {".docx", ".dotm", ".dotx", ".docm"}:
        return False

    with TemporaryDirectory(dir=path.parent) as directory:
        target = Path(directory) / path.name
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(target, "w") as zout:
            changed = bool(zin.comment)
            for info in zin.infolist():
                data = zin.read(info)
                cleaned = data
                if info.filename in PUBLIC_PROPERTIES:
                    cleaned = _keep_public_properties(data, PUBLIC_PROPERTIES[info.filename])
                elif info.filename == VBA_PART:
                    cleaned = _sanitize_vba(data)
                changed |= (cleaned != data or info.date_time != PACKAGE_DATE
                            or bool(info.extra or info.comment))
                info.date_time = PACKAGE_DATE
                info.extra = info.comment = b""
                zout.writestr(info, cleaned)
        if changed:
            target.replace(path)
    return changed


def sanitize_pdf_metadata(path: str | Path) -> None:
    """先断开元数据引用，再复制预览 PDF 的有效对象。"""
    from pypdf import PdfReader, PdfWriter

    path = Path(path)
    reader = PdfReader(path)
    reader.trailer.pop("/Info", None)
    reader.trailer.pop("/ID", None)
    reader.root_object.pop("/Metadata", None)
    writer = PdfWriter(clone_from=reader, keep_initial_header=True)
    with TemporaryDirectory(dir=path.parent) as directory:
        target = Path(directory) / path.name
        with target.open("wb") as stream:
            writer.write(stream)
        target.replace(path)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: python src/office_metadata.py <built-docx|built-dotm|preview-pdf> [...]")
        raise SystemExit(2)
    for value in sys.argv[1:]:
        target = Path(value)
        if target.suffix.lower() == ".pdf":
            sanitize_pdf_metadata(target)
            print("sanitized", target.name)
        else:
            print("sanitized" if sanitize_office_metadata(target) else "unchanged", target.name)
