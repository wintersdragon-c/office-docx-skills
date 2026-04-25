#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    from docx import Document
    from lxml import etree
except ImportError:
    print("Missing dependency: install python-docx and lxml")
    sys.exit(1)


ROOT = Path(__file__).resolve().parents[2]
HELPER = ROOT / "skills" / "docx-tracked-changes" / "tracked_change_editor.py"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def load_editor_class():
    spec = importlib.util.spec_from_file_location("tracked_change_editor", HELPER)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load tracked_change_editor.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.TrackedChangeEditor


def main() -> None:
    if not HELPER.exists():
        print(f"Missing helper: {HELPER}")
        sys.exit(1)
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        input_docx = tmpdir / "input.docx"
        output_docx = tmpdir / "output.docx"

        doc = Document()
        doc.add_paragraph("Original paragraph.")
        doc.save(input_docx)

        editor_cls = load_editor_class()
        editor = editor_cls(input_docx, author="Office DOCX Skills Test")
        editor.replace_paragraph_text(0, "Replacement paragraph.")
        editor.save(output_docx)

        if not output_docx.exists():
            print("Output DOCX was not created")
            sys.exit(1)
        if input_docx.read_bytes() == output_docx.read_bytes():
            print("Output DOCX should differ from input DOCX")
            sys.exit(1)

        with zipfile.ZipFile(output_docx) as archive:
            xml = archive.read("word/document.xml")
        root = etree.fromstring(xml)
        if not root.findall(".//w:ins", NS):
            print("Missing w:ins tracked insertion")
            sys.exit(1)
        if not root.findall(".//w:del", NS):
            print("Missing w:del tracked deletion")
            sys.exit(1)
        if not root.findall(".//w:delText", NS):
            print("Missing w:delText deleted text")
            sys.exit(1)

    print("PASS: tracked changes smoke test")


if __name__ == "__main__":
    main()
