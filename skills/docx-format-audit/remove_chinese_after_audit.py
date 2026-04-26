#!/usr/bin/env python3
from __future__ import annotations

import argparse
from copy import deepcopy
import importlib.util
import sys
from pathlib import Path
import shutil
import tempfile
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def qn(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def _load_audit_module():
    audit_path = Path(__file__).with_name("audit_docx_translation.py")
    spec = importlib.util.spec_from_file_location("audit_docx_translation", audit_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load audit_docx_translation.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _source_pair_paragraphs(root: etree._Element) -> list[etree._Element]:
    paragraphs: list[etree._Element] = []
    for bookmark in root.findall(".//w:bookmarkStart", NS):
        name = bookmark.get(qn(W, "name")) or ""
        if not (name.startswith("btx_") and name.endswith("_src")):
            continue
        paragraph = bookmark
        while paragraph is not None and paragraph.tag != qn(W, "p"):
            paragraph = paragraph.getparent()
        if paragraph is not None and paragraph not in paragraphs:
            paragraphs.append(paragraph)
    return paragraphs


def _remove_source_pair_paragraphs_in_place(root: etree._Element) -> int:
    removed = 0
    for paragraph in _source_pair_paragraphs(root):
        parent = paragraph.getparent()
        if parent is not None:
            parent.remove(paragraph)
            removed += 1
    return removed


def remove_chinese_source_paragraphs(root: etree._Element, audit_module=None) -> int:
    audit_module = audit_module or _load_audit_module()
    report = audit_module.audit_document_root(root)
    if report.blocking:
        messages = "; ".join(item.message for item in report.blocking)
        raise RuntimeError(f"cannot remove Chinese source paragraphs; audit blocking findings: {messages}")
    trial_root = deepcopy(root)
    removed = _remove_source_pair_paragraphs_in_place(trial_root)
    post_report = audit_module.audit_document_root(trial_root, mode="english-only")
    if post_report.blocking:
        messages = "; ".join(item.message for item in post_report.blocking)
        raise RuntimeError(f"cannot finalize English-only document; English-only audit blocking findings: {messages}")
    _remove_source_pair_paragraphs_in_place(root)
    return removed


def remove_chinese_from_docx(input_docx: str | Path, output_docx: str | Path) -> int:
    input_docx = Path(input_docx)
    output_docx = Path(output_docx)
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        with zipfile.ZipFile(input_docx) as archive:
            archive.extractall(tmpdir)
        document_xml = tmpdir / "word" / "document.xml"
        root = etree.fromstring(document_xml.read_bytes())
        removed = remove_chinese_source_paragraphs(root)
        document_xml.write_bytes(etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))
        if output_docx.exists():
            output_docx.unlink()
        shutil.make_archive(str(output_docx.with_suffix("")), "zip", tmpdir)
        output_docx.with_suffix(".zip").replace(output_docx)
    return removed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Remove Chinese source paragraphs from an audited bilingual DOCX.")
    parser.add_argument("input_docx")
    parser.add_argument("output_docx")
    args = parser.parse_args(argv)
    removed = remove_chinese_from_docx(args.input_docx, args.output_docx)
    print(f"Removed source paragraphs: {removed}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
