#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import subprocess
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
VERIFY = ROOT / "skills" / "docx-tracked-changes" / "verify_tracked_changes.py"
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
    if not VERIFY.exists():
        print(f"Missing verifier: {VERIFY}")
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
        editor.insert_paragraph_after_with_tracked_change(0, "Inserted tracked paragraph.")
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
        inserted = [
            "".join(t.text for t in ins.findall(".//w:t", NS) if t.text)
            for ins in root.findall(".//w:ins", NS)
        ]
        if "Inserted tracked paragraph." not in inserted:
            print("Missing tracked inserted paragraph")
            sys.exit(1)

        verified = subprocess.run(
            [
                sys.executable,
                str(VERIFY),
                str(output_docx),
                "--author",
                "Office DOCX Skills Test",
            ],
            check=False,
            text=True,
            capture_output=True,
        )
        if verified.returncode != 0:
            print(verified.stdout)
            print(verified.stderr)
            sys.exit(1)
        if "insertions: 2" not in verified.stdout:
            print("Verifier did not report the expected insertion count")
            print(verified.stdout)
            sys.exit(1)

    print("PASS: tracked changes smoke test")


def test_tracked_changes_smoke() -> None:
    main()


if __name__ == "__main__":
    main()
