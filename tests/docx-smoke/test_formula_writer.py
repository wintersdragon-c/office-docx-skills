#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    from lxml import etree
except ImportError:
    print("Missing dependency: install lxml")
    sys.exit(1)


ROOT = Path(__file__).resolve().parents[2]
HELPER = ROOT / "skills" / "word-formula-writing" / "formula_writer.py"
NS = {
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def load_module():
    spec = importlib.util.spec_from_file_location("formula_writer", HELPER)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load formula_writer.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> None:
    if not HELPER.exists():
        print(f"Missing helper: {HELPER}")
        sys.exit(1)

    with tempfile.TemporaryDirectory() as tmp:
        output_docx = Path(tmp) / "formula.docx"
        module = load_module()
        module.create_docx_with_omml_formula(
            output_docx,
            "x+1=2",
            body_text="Formula smoke test",
        )

        if not output_docx.exists():
            print("Formula DOCX was not created")
            sys.exit(1)

        with zipfile.ZipFile(output_docx) as archive:
            xml = archive.read("word/document.xml")
        root = etree.fromstring(xml)
        equations = root.findall(".//m:oMath", NS)
        if len(equations) != 1:
            print(f"Expected 1 m:oMath element, found {len(equations)}")
            sys.exit(1)
        math_text = "".join(t.text for t in equations[0].findall(".//m:t", NS) if t.text)
        if math_text != "x+1=2":
            print(f"Unexpected formula text: {math_text!r}")
            sys.exit(1)

    print("PASS: formula writer smoke test")


def test_formula_writer_smoke() -> None:
    main()


if __name__ == "__main__":
    main()
