from __future__ import annotations

import copy
import shutil
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NSMAP = {"w": W}


def qn(tag: str) -> str:
    return f"{{{W}}}{tag.split(':')[-1]}"


class TrackedChangeEditor:
    """Edit .docx body paragraphs so Word shows tracked revisions."""

    def __init__(self, docx_path: str | Path, author: str = "Codex") -> None:
        self.docx_path = Path(docx_path)
        self.author = author
        self.date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

        with zipfile.ZipFile(self.docx_path, "r") as archive:
            self.doc_xml = archive.read("word/document.xml")

        self.root = etree.fromstring(self.doc_xml)
        self.body = self.root.find("w:body", NSMAP)
        if self.body is None:
            raise ValueError("word/document.xml does not contain a body element")

        self._find_max_id()
        self._index_body_paragraphs()

    def _find_max_id(self) -> None:
        self.max_id = 0
        for tag in ["ins", "del", "rPrChange", "pPrChange", "sectPrChange"]:
            for element in self.root.findall(f".//w:{tag}", NSMAP):
                revision_id = element.get(qn("w:id"))
                if not revision_id:
                    continue
                try:
                    self.max_id = max(self.max_id, int(revision_id))
                except ValueError:
                    continue

    def _next_id(self) -> str:
        self.max_id += 1
        return str(self.max_id)

    def _index_body_paragraphs(self) -> None:
        """Index paragraphs after the TOC (w:sdt block), or all if none exists."""
        self.body_paras = []
        past_toc = False
        for child in self.body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "sdt":
                past_toc = True
                continue
            if past_toc and tag == "p":
                self.body_paras.append(child)
        if not self.body_paras:
            self.body_paras = self.body.findall("w:p", NSMAP)

    def _get_para_text(self, paragraph) -> str:
        return "".join(text.text for text in paragraph.findall(".//w:t", NSMAP) if text.text)

    def _get_default_rpr(self, paragraph):
        for run in paragraph.findall(".//w:r", NSMAP):
            run_properties = run.find("w:rPr", NSMAP)
            if run_properties is not None:
                return copy.deepcopy(run_properties)
        return None

    def _collect_direct_runs(self, paragraph):
        return [child for child in paragraph if child.tag == qn("w:r")]

    def replace_paragraph_text(self, para_index: int, new_text: str) -> None:
        """Replace all direct runs in a body paragraph using tracked changes."""
        paragraph = self.body_paras[para_index]
        direct_runs = self._collect_direct_runs(paragraph)
        if not direct_runs:
            return

        default_rpr = self._get_default_rpr(paragraph)
        first_run_pos = list(paragraph).index(direct_runs[0])

        deletion = etree.Element(qn("w:del"))
        deletion.set(qn("w:id"), self._next_id())
        deletion.set(qn("w:author"), self.author)
        deletion.set(qn("w:date"), self.date)
        for run in direct_runs:
            for text in run.findall("w:t", NSMAP):
                text.tag = qn("w:delText")
            paragraph.remove(run)
            deletion.append(run)

        insertion = etree.Element(qn("w:ins"))
        insertion.set(qn("w:id"), self._next_id())
        insertion.set(qn("w:author"), self.author)
        insertion.set(qn("w:date"), self.date)
        new_run = etree.SubElement(insertion, qn("w:r"))
        if default_rpr is not None:
            new_run.append(copy.deepcopy(default_rpr))
        new_text_node = etree.SubElement(new_run, qn("w:t"))
        new_text_node.set(XML_SPACE, "preserve")
        new_text_node.text = new_text

        paragraph.insert(first_run_pos, deletion)
        paragraph.insert(first_run_pos + 1, insertion)

    def insert_paragraph_after_with_tracked_change(self, para_index: int, new_text: str) -> None:
        """Insert a new body paragraph whose content appears as a tracked insertion."""
        anchor = self.body_paras[para_index]
        new_paragraph = etree.Element(qn("w:p"))
        insertion = etree.SubElement(new_paragraph, qn("w:ins"))
        insertion.set(qn("w:id"), self._next_id())
        insertion.set(qn("w:author"), self.author)
        insertion.set(qn("w:date"), self.date)
        new_run = etree.SubElement(insertion, qn("w:r"))
        default_rpr = self._get_default_rpr(anchor)
        if default_rpr is not None:
            new_run.append(copy.deepcopy(default_rpr))
        new_text_node = etree.SubElement(new_run, qn("w:t"))
        new_text_node.set(XML_SPACE, "preserve")
        new_text_node.text = new_text

        body_children = list(self.body)
        anchor_position = body_children.index(anchor)
        self.body.insert(anchor_position + 1, new_paragraph)
        self._index_body_paragraphs()

    def save(self, output_path: str | Path) -> None:
        output_path = Path(output_path)
        doc_xml_new = etree.tostring(
            self.root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        with NamedTemporaryFile(suffix=".docx", delete=False) as tmp_file:
            tmp_path = Path(tmp_file.name)

        try:
            with zipfile.ZipFile(self.docx_path, "r") as zin:
                with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename == "word/document.xml":
                            zout.writestr(item, doc_xml_new)
                        else:
                            zout.writestr(item, zin.read(item.filename))
            shutil.move(tmp_path, output_path)
        finally:
            if tmp_path.exists():
                tmp_path.unlink()
