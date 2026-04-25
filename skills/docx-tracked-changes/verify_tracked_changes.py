#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W}


def qn(tag: str) -> str:
    return f"{{{W}}}{tag.split(':')[-1]}"


def element_text(element, text_tag: str) -> str:
    return "".join(text.text for text in element.findall(f".//w:{text_tag}", NSMAP) if text.text)


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify Word tracked changes in a DOCX file.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--author", default=None)
    args = parser.parse_args()

    with zipfile.ZipFile(args.docx) as archive:
        xml = archive.read("word/document.xml")
    root = etree.fromstring(xml)

    insertions = root.findall(".//w:ins", NSMAP)
    deletions = root.findall(".//w:del", NSMAP)
    if args.author:
        insertions = [item for item in insertions if item.get(qn("w:author")) == args.author]
        deletions = [item for item in deletions if item.get(qn("w:author")) == args.author]

    print(f"author: {args.author or '(any)'}")
    print(f"insertions: {len(insertions)}")
    for index, insertion in enumerate(insertions, start=1):
        print(f"  ins[{index}]: {element_text(insertion, 't')[:100]}")
    print(f"deletions: {len(deletions)}")
    for index, deletion in enumerate(deletions, start=1):
        print(f"  del[{index}]: {element_text(deletion, 'delText')[:100]}")

    if not insertions:
        print("No tracked insertions found", file=sys.stderr)
        return 1
    if not deletions:
        print("No tracked deletions found", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
