#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

ROOT = Path(__file__).resolve().parents[2]
AUDIT = ROOT / "skills" / "docx-format-audit" / "audit_docx_translation.py"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W, "m": M, "r": R}


def load_module():
    spec = importlib.util.spec_from_file_location("audit_docx_translation", AUDIT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def qn(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def make_doc(formula_in_english: bool = True, table: bool = False) -> etree._Element:
    english_formula = '<m:oMath><m:r><m:t>x+1=2</m:t></m:r></m:oMath>' if formula_in_english else ""
    body_pair = f"""
      <w:p>
        <w:bookmarkStart w:id="1" w:name="btx_p1_src"/><w:bookmarkEnd w:id="1"/>
        <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>
        <w:r><w:t>中文公式</w:t></w:r>
        <m:oMath><m:r><m:t>x+1=2</m:t></m:r></m:oMath>
      </w:p>
      <w:p>
        <w:bookmarkStart w:id="2" w:name="btx_p1_en"/><w:bookmarkEnd w:id="2"/>
        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>
        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/></w:rPr><w:t>English formula</w:t></w:r>
        {english_formula}
      </w:p>
    """
    table_pair = f"""
      <w:tbl><w:tr><w:tc>
        <w:p><w:bookmarkStart w:id="3" w:name="btx_t1_src"/><w:bookmarkEnd w:id="3"/><w:r><w:t>表格中文</w:t></w:r></w:p>
        <w:p><w:bookmarkStart w:id="4" w:name="btx_t1_en"/><w:bookmarkEnd w:id="4"/><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/></w:rPr><w:t>Table English</w:t></w:r></w:p>
      </w:tc></w:tr></w:tbl>
    """ if table else ""
    xml = f'<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:r="{R}"><w:body>{body_pair}{table_pair}</w:body></w:document>'
    return etree.fromstring(xml.encode("utf-8"))


def test_audit_passes_for_pair_marker_font_formula_superscript_and_table() -> None:
    module = load_module()
    report = module.audit_document_root(make_doc(table=True))
    assert report.blocking == []


def test_audit_fails_when_english_formula_signature_is_missing() -> None:
    module = load_module()
    report = module.audit_document_root(make_doc(formula_in_english=False))
    assert any("formula signature" in item.message for item in report.blocking)


def test_audit_fails_when_english_font_is_not_times_new_roman() -> None:
    module = load_module()
    root = make_doc()
    fonts = root.xpath(".//w:bookmarkStart[@w:name='btx_p1_en']/../w:r/w:rPr/w:rFonts", namespaces=NS)[0]
    fonts.set(qn(W, "ascii"), "Arial")
    report = module.audit_document_root(root)
    assert any("Times New Roman" in item.message for item in report.blocking)


def test_audit_accepts_times_new_roman_from_style_inheritance() -> None:
    module = load_module()
    root = make_doc()
    for fonts in root.xpath(".//w:bookmarkStart[@w:name='btx_p1_en']/../w:r/w:rPr/w:rFonts", namespaces=NS):
        fonts.getparent().remove(fonts)
    styles = etree.fromstring(f"""<w:styles xmlns:w="{W}">
      <w:style w:type="paragraph" w:styleId="Normal">
        <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/></w:rPr>
      </w:style>
    </w:styles>""".encode("utf-8"))
    report = module.audit_document_root(root, styles_root=styles)
    assert report.blocking == []


def test_audit_fails_when_superscript_marker_is_missing() -> None:
    module = load_module()
    root = make_doc()
    en_sup = root.xpath(".//w:bookmarkStart[@w:name='btx_p1_en']/../w:r/w:rPr/w:vertAlign", namespaces=NS)[0]
    en_sup.getparent().remove(en_sup)
    report = module.audit_document_root(root)
    assert any("superscript" in item.message for item in report.blocking)


def test_audit_fails_when_table_cell_pair_is_missing() -> None:
    module = load_module()
    root = make_doc(table=True)
    en_table_para = root.xpath(".//w:bookmarkStart[@w:name='btx_t1_en']/..", namespaces=NS)[0]
    en_table_para.getparent().remove(en_table_para)
    report = module.audit_document_root(root)
    assert any("table-cell" in item.message for item in report.blocking)


def test_audit_fails_for_unmarked_chinese_body_paragraph() -> None:
    module = load_module()
    root = make_doc()
    body = root.find(".//w:body", NS)
    assert body is not None
    unmarked = etree.fromstring(
        f'<w:p xmlns:w="{W}"><w:r><w:t>未标记中文段落</w:t></w:r></w:p>'.encode("utf-8")
    )
    body.append(unmarked)
    report = module.audit_document_root(root)
    assert any("unmarked Chinese body paragraph" in item.message for item in report.blocking)


def test_audit_fails_for_unmarked_chinese_table_cell_paragraph() -> None:
    module = load_module()
    root = make_doc(table=True)
    cell = root.find(".//w:tc", NS)
    assert cell is not None
    unmarked = etree.fromstring(
        f'<w:p xmlns:w="{W}"><w:r><w:t>未标记表格中文</w:t></w:r></w:p>'.encode("utf-8")
    )
    cell.append(unmarked)
    report = module.audit_document_root(root)
    assert any("unmarked Chinese table-cell paragraph" in item.message for item in report.blocking)


def test_english_only_audit_fails_on_residual_chinese() -> None:
    module = load_module()
    root = make_doc()
    report = module.audit_document_root(root, mode="english-only")
    assert any("Chinese text remains" in item.message for item in report.blocking)


def test_english_only_audit_allows_remaining_english_markers_after_source_removal() -> None:
    module = load_module()
    root = make_doc()
    source_para = root.xpath(".//w:bookmarkStart[@w:name='btx_p1_src']/..", namespaces=NS)[0]
    source_para.getparent().remove(source_para)
    report = module.audit_document_root(root, mode="english-only")
    assert report.blocking == []


def test_english_only_audit_checks_remaining_english_fonts() -> None:
    module = load_module()
    root = make_doc()
    source_para = root.xpath(".//w:bookmarkStart[@w:name='btx_p1_src']/..", namespaces=NS)[0]
    source_para.getparent().remove(source_para)
    fonts = root.xpath(".//w:bookmarkStart[@w:name='btx_p1_en']/../w:r/w:rPr/w:rFonts", namespaces=NS)[0]
    fonts.set(qn(W, "ascii"), "Arial")
    report = module.audit_document_root(root, mode="english-only")
    assert any("Times New Roman" in item.message for item in report.blocking)


def test_audit_docx_file_blocks_when_python_docx_cannot_open_package() -> None:
    module = load_module()
    with tempfile.TemporaryDirectory() as tmp:
        broken_docx = Path(tmp) / "broken.docx"
        with zipfile.ZipFile(broken_docx, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(
                "word/document.xml",
                f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>English</w:t></w:r></w:p></w:body></w:document>',
            )
        report = module.audit_docx_file(broken_docx)
        assert any("cannot be opened by python-docx" in item.message for item in report.blocking)


def test_audit_docx_file_uses_python_docx_openability_for_valid_package() -> None:
    module = load_module()
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = Path(tmp) / "valid.docx"
        document = Document()
        document.add_paragraph("English only")
        document.save(docx_path)
        report = module.audit_docx_file(docx_path, mode="english-only")
        assert not any("cannot be opened by python-docx" in item.message for item in report.blocking)
