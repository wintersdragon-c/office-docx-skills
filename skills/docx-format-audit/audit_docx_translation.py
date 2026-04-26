#!/usr/bin/env python3
from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from pathlib import Path
import zipfile

from docx import Document
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
NS = {"w": W, "m": M, "r": R, "v": V}


def qn(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


@dataclass
class Finding:
    severity: str
    pair_id: str
    message: str


@dataclass
class AuditReport:
    blocking: list[Finding] = field(default_factory=list)
    warnings: list[Finding] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.blocking


def _text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS))


def _has_chinese(text: str) -> bool:
    return any("\u4e00" <= char <= "\u9fff" for char in text)


def _pair_bookmarks(root: etree._Element) -> dict[str, dict[str, etree._Element]]:
    pairs: dict[str, dict[str, etree._Element]] = {}
    for bookmark in root.findall(".//w:bookmarkStart", NS):
        name = bookmark.get(qn(W, "name")) or ""
        if not name.startswith("btx_"):
            continue
        role = "src" if name.endswith("_src") else "en" if name.endswith("_en") else ""
        if not role:
            continue
        pair_id = name[4 : -4 if role == "src" else -3]
        paragraph = bookmark
        while paragraph is not None and paragraph.tag != qn(W, "p"):
            paragraph = paragraph.getparent()
        if paragraph is not None:
            pairs.setdefault(pair_id, {})[role] = paragraph
    return pairs


def _style_font_map(styles_root: etree._Element | None) -> dict[str, dict[str, str]]:
    if styles_root is None:
        return {}
    result: dict[str, dict[str, str]] = {}
    for style in styles_root.findall(".//w:style", NS):
        style_id = style.get(qn(W, "styleId"))
        fonts = style.find(".//w:rFonts", NS)
        if style_id and fonts is not None:
            result[style_id] = {name: fonts.get(qn(W, name), "") for name in ["ascii", "hAnsi", "cs"]}
    return result


def _paragraph_style_id(paragraph: etree._Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.get(qn(W, "val")) if style is not None else "Normal"


def _english_font_ok(paragraph: etree._Element, styles_root: etree._Element | None) -> bool:
    inherited = _style_font_map(styles_root).get(_paragraph_style_id(paragraph), {})
    text_runs = [
        run
        for run in paragraph.findall(".//w:r", NS)
        if "".join(run.xpath(".//w:t/text()", namespaces=NS)).strip()
    ]
    for run in text_runs:
        if _has_chinese("".join(run.xpath(".//w:t/text()", namespaces=NS))):
            continue
        effective = dict(inherited)
        fonts = run.find("./w:rPr/w:rFonts", NS)
        if fonts is not None:
            for name in ["ascii", "hAnsi", "cs"]:
                value = fonts.get(qn(W, name))
                if value:
                    effective[name] = value
        if any(effective.get(name) != "Times New Roman" for name in ["ascii", "hAnsi", "cs"]):
            return False
    return True


def _formula_signatures(paragraph: etree._Element, relationship_targets: dict[str, str] | None = None) -> list[str]:
    relationship_targets = relationship_targets or {}
    signatures: list[str] = []
    for node in paragraph.findall(".//m:oMath", NS):
        signatures.append("m:oMath:" + "".join(node.xpath(".//m:t/text()", namespaces=NS)))
    for xpath, kind in [(".//w:object", "w:object"), (".//w:drawing", "w:drawing"), (".//v:imagedata", "v:imagedata")]:
        for node in paragraph.findall(xpath, NS):
            rel_id = node.get(qn(R, "id")) or node.get(qn(R, "embed")) or ""
            if not rel_id:
                values = node.xpath(".//@r:embed | .//@r:id", namespaces=NS)
                rel_id = values[0] if values else ""
            target = relationship_targets.get(rel_id, rel_id)
            signatures.append(f"{kind}:{target}")
    return signatures


def _run_markers(paragraph: etree._Element, marker: str) -> list[str]:
    return [item.get(qn(W, "val"), "") for item in paragraph.findall(f".//w:{marker}", NS)]


def _same_table_cell(left: etree._Element, right: etree._Element) -> bool:
    def cell(node: etree._Element) -> etree._Element | None:
        while node is not None and node.tag != qn(W, "tc"):
            node = node.getparent()
        return node

    return cell(left) is cell(right)


def _inside_table_cell(paragraph: etree._Element) -> bool:
    node = paragraph.getparent()
    while node is not None:
        if node.tag == qn(W, "tc"):
            return True
        node = node.getparent()
    return False


def _marked_paragraph_ids(pairs: dict[str, dict[str, etree._Element]]) -> set[int]:
    result: set[int] = set()
    for pair in pairs.values():
        for role in ["src", "en"]:
            paragraph = pair.get(role)
            if paragraph is not None:
                result.add(id(paragraph))
    return result


def _check_unmarked_chinese_paragraphs(
    root: etree._Element,
    pairs: dict[str, dict[str, etree._Element]],
    report: AuditReport,
) -> None:
    marked_ids = _marked_paragraph_ids(pairs)
    for paragraph in root.findall(".//w:p", NS):
        if id(paragraph) in marked_ids:
            continue
        if not _has_chinese(_text(paragraph)):
            continue
        context = "table-cell" if _inside_table_cell(paragraph) else "body"
        report.blocking.append(
            Finding(
                "blocking",
                "unmarked",
                f"unmarked Chinese {context} paragraph has no pair-ID-linked English translation",
            )
        )


def _paragraph_label(paragraph: etree._Element) -> str:
    bookmark = paragraph.find("./w:bookmarkStart", NS)
    if bookmark is not None:
        return bookmark.get(qn(W, "name")) or "paragraph"
    return "paragraph"


def _english_only_paragraphs(root: etree._Element) -> list[etree._Element]:
    paragraphs: list[etree._Element] = []
    for paragraph in root.findall(".//w:p", NS):
        text = _text(paragraph).strip()
        if text and not _has_chinese(text):
            paragraphs.append(paragraph)
    return paragraphs


def _relationship_ids(paragraph: etree._Element) -> list[str]:
    result: list[str] = []
    for value in paragraph.xpath(".//@r:embed | .//@r:id | .//@r:link", namespaces=NS):
        if value and value not in result:
            result.append(value)
    return result


def _check_english_only_output(
    root: etree._Element,
    styles_root: etree._Element | None,
    relationship_targets: dict[str, str] | None,
    report: AuditReport,
) -> None:
    for paragraph in _english_only_paragraphs(root):
        label = _paragraph_label(paragraph)
        if not _english_font_ok(paragraph, styles_root):
            report.blocking.append(Finding("blocking", label, "English-only text does not resolve to Times New Roman"))
        if relationship_targets is not None:
            for rel_id in _relationship_ids(paragraph):
                if rel_id not in relationship_targets:
                    report.blocking.append(
                        Finding("blocking", label, f"missing relationship target for preserved object: {rel_id}")
                    )


def audit_document_root(
    root: etree._Element,
    styles_root: etree._Element | None = None,
    relationship_targets: dict[str, str] | None = None,
    mode: str = "bilingual",
) -> AuditReport:
    report = AuditReport()
    if mode == "english-only" and _has_chinese(_text(root)):
        report.blocking.append(Finding("blocking", "document", "Chinese text remains in English-only output"))
    if mode == "english-only":
        _check_english_only_output(root, styles_root, relationship_targets, report)
        return report
    pairs = _pair_bookmarks(root)
    if mode == "bilingual":
        _check_unmarked_chinese_paragraphs(root, pairs, report)
    for pair_id, pair in sorted(pairs.items()):
        src = pair.get("src")
        en = pair.get("en")
        if src is None or en is None:
            context = "table-cell " if src is not None and _inside_table_cell(src) else ""
            report.blocking.append(Finding("blocking", pair_id, f"missing {context}source or English pair marker"))
            continue
        if _inside_table_cell(src) and not _same_table_cell(src, en):
            report.blocking.append(Finding("blocking", pair_id, "table-cell pair is not in the same cell"))
        if _has_chinese(_text(src)) and not _text(en).strip():
            report.blocking.append(Finding("blocking", pair_id, "English counterpart is empty"))
        if not _english_font_ok(en, styles_root):
            report.blocking.append(Finding("blocking", pair_id, "English inserted text does not resolve to Times New Roman"))
        for signature in _formula_signatures(src, relationship_targets):
            if signature not in _formula_signatures(en, relationship_targets):
                report.blocking.append(Finding("blocking", pair_id, f"missing formula signature in English pair: {signature}"))
        marker_labels = {
            "vertAlign": "superscript/subscript",
            "footnoteReference": "footnoteReference",
            "endnoteReference": "endnoteReference",
        }
        for marker, label in marker_labels.items():
            if _run_markers(src, marker) and not _run_markers(en, marker):
                report.blocking.append(Finding("blocking", pair_id, f"missing {label} marker in English pair"))
    return report


def read_docx_root(path: str | Path) -> etree._Element:
    with zipfile.ZipFile(path) as archive:
        return etree.fromstring(archive.read("word/document.xml"))


def read_docx_styles(path: str | Path) -> etree._Element | None:
    with zipfile.ZipFile(path) as archive:
        try:
            return etree.fromstring(archive.read("word/styles.xml"))
        except KeyError:
            return None


def audit_docx_file(path: str | Path, mode: str = "bilingual") -> AuditReport:
    path = Path(path)
    try:
        Document(path)
    except Exception as exc:
        return AuditReport(
            blocking=[
                Finding(
                    "blocking",
                    "document",
                    f"DOCX cannot be opened by python-docx: {exc}",
                )
            ]
        )
    return audit_document_root(read_docx_root(path), read_docx_styles(path), mode=mode)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Audit a bilingual DOCX translation.")
    parser.add_argument("docx", help="Bilingual or English-only DOCX to audit")
    parser.add_argument("--mode", choices=["bilingual", "english-only"], default="bilingual")
    args = parser.parse_args(argv)
    report = audit_docx_file(args.docx, mode=args.mode)
    for finding in report.blocking:
        print(f"BLOCKING [{finding.pair_id}] {finding.message}")
    for finding in report.warnings:
        print(f"WARNING [{finding.pair_id}] {finding.message}")
    if report.ok:
        print("PASS: DOCX translation audit")
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
