#!/usr/bin/env python3
from __future__ import annotations

from copy import deepcopy
from hashlib import sha256
from pathlib import Path
import re
import shutil

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W, "m": M, "r": R, "v": V, "a": A}


def qn(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def has_chinese(text: str) -> bool:
    return any("\u4e00" <= char <= "\u9fff" for char in text)


def set_english_rfonts(r_pr: etree._Element | None = None) -> etree._Element:
    if r_pr is None:
        r_pr = etree.Element(qn(W, "rPr"))
    r_fonts = r_pr.find("w:rFonts", NS)
    if r_fonts is None:
        r_fonts = etree.SubElement(r_pr, qn(W, "rFonts"))
    for attr_name in ["ascii", "hAnsi", "cs", "eastAsia"]:
        r_fonts.set(qn(W, attr_name), "Times New Roman")
    return r_pr


def _next_bookmark_id(root: etree._Element, requested: int) -> int:
    existing = []
    for item in root.findall(".//w:bookmarkStart", NS):
        value = item.get(qn(W, "id"))
        if value and value.isdigit():
            existing.append(int(value))
    return max([requested - 1, *existing]) + 1


def _add_bookmark(paragraph: etree._Element, name: str, bookmark_id: int) -> None:
    start = etree.Element(qn(W, "bookmarkStart"))
    start.set(qn(W, "id"), str(bookmark_id))
    start.set(qn(W, "name"), name)
    end = etree.Element(qn(W, "bookmarkEnd"))
    end.set(qn(W, "id"), str(bookmark_id))
    p_pr = paragraph.find("w:pPr", NS)
    insert_at = 1 if p_pr is not None else 0
    paragraph.insert(insert_at, start)
    paragraph.insert(insert_at + 1, end)


def ensure_pair_bookmarks(
    source_paragraph: etree._Element,
    english_paragraph: etree._Element,
    pair_id: str,
    bookmark_start_id: int = 1,
) -> tuple[int, int]:
    root = source_paragraph.getroottree().getroot()
    first_id = _next_bookmark_id(root, bookmark_start_id)
    _add_bookmark(source_paragraph, f"btx_{pair_id}_src", first_id)
    second_id = _next_bookmark_id(root, first_id + 1)
    _add_bookmark(english_paragraph, f"btx_{pair_id}_en", second_id)
    return first_id, second_id


def rewrite_relationship_references(node: etree._Element, old_rid: str, new_rid: str) -> int:
    changed = 0
    for element in node.iter():
        for attr_name in [qn(R, "id"), qn(R, "embed"), qn(R, "link")]:
            if element.get(attr_name) == old_rid:
                element.set(attr_name, new_rid)
                changed += 1
    return changed


def _run_has_formula_candidate(run: etree._Element) -> bool:
    for xpath in [
        ".//m:oMath",
        ".//m:oMathPara",
        ".//w:object",
        ".//w:drawing",
        ".//v:shape",
        ".//v:imagedata",
    ]:
        if run.find(xpath, NS) is not None:
            return True
    return False


def _formula_candidate_runs(paragraph: etree._Element) -> list[etree._Element]:
    return [run for run in paragraph.findall("./w:r", NS) if _run_has_formula_candidate(run)]


def _standalone_math_nodes(paragraph: etree._Element) -> list[etree._Element]:
    return paragraph.findall("./m:oMath", NS) + paragraph.findall("./m:oMathPara", NS)


def _run_has_preserved_marker(run: etree._Element) -> bool:
    for xpath in [
        ".//w:footnoteReference",
        ".//w:endnoteReference",
        "./w:rPr/w:vertAlign",
    ]:
        if run.find(xpath, NS) is not None:
            return True
    return False


def _preserved_marker_runs(paragraph: etree._Element) -> list[etree._Element]:
    return [
        run
        for run in paragraph.findall("./w:r", NS)
        if _run_has_preserved_marker(run) and not _run_has_formula_candidate(run)
    ]


def _remove_marker_formatting(r_pr: etree._Element) -> etree._Element:
    for marker in r_pr.findall("./w:vertAlign", NS):
        r_pr.remove(marker)
    return r_pr


def _body_text_r_pr(source_paragraph: etree._Element) -> etree._Element:
    for run in source_paragraph.findall("./w:r", NS):
        if _run_has_formula_candidate(run) or _run_has_preserved_marker(run):
            continue
        if not "".join(run.xpath(".//w:t/text()", namespaces=NS)).strip():
            continue
        r_pr = run.find("./w:rPr", NS)
        if r_pr is not None:
            return _remove_marker_formatting(deepcopy(r_pr))
    source_r_pr = source_paragraph.find(".//w:rPr", NS)
    if source_r_pr is not None:
        return _remove_marker_formatting(deepcopy(source_r_pr))
    return etree.Element(qn(W, "rPr"))


def make_english_paragraph_like(
    source_paragraph: etree._Element,
    english_text: str,
    relationship_id_map: dict[str, str] | None = None,
) -> etree._Element:
    relationship_id_map = relationship_id_map or {}
    paragraph = etree.Element(qn(W, "p"))
    p_pr = source_paragraph.find("w:pPr", NS)
    if p_pr is not None:
        paragraph.append(deepcopy(p_pr))
    r_pr = _body_text_r_pr(source_paragraph)
    set_english_rfonts(r_pr)
    run = etree.SubElement(paragraph, qn(W, "r"))
    run.append(r_pr)
    text = etree.SubElement(run, qn(W, "t"))
    text.text = english_text
    for formula_run in _formula_candidate_runs(source_paragraph):
        copied_run = deepcopy(formula_run)
        for old_rid, new_rid in relationship_id_map.items():
            rewrite_relationship_references(copied_run, old_rid, new_rid)
        paragraph.append(copied_run)
    for math_node in _standalone_math_nodes(source_paragraph):
        paragraph.append(deepcopy(math_node))
    for marker_run in _preserved_marker_runs(source_paragraph):
        paragraph.append(deepcopy(marker_run))
    return paragraph


def insert_translation_paragraph_after(
    source_paragraph: etree._Element,
    english_text: str,
    pair_id: str,
    bookmark_start_id: int = 1,
    relationship_id_map: dict[str, str] | None = None,
) -> etree._Element:
    parent = source_paragraph.getparent()
    if parent is None:
        raise ValueError("source paragraph has no parent")
    english_paragraph = make_english_paragraph_like(source_paragraph, english_text, relationship_id_map)
    parent.insert(parent.index(source_paragraph) + 1, english_paragraph)
    ensure_pair_bookmarks(source_paragraph, english_paragraph, pair_id, bookmark_start_id)
    return english_paragraph


def collect_formula_signatures(
    paragraph: etree._Element,
    relationship_targets: dict[str, str] | None = None,
    related_blobs: dict[str, bytes] | None = None,
) -> list[dict[str, str]]:
    relationship_targets = relationship_targets or {}
    related_blobs = related_blobs or {}
    signatures: list[dict[str, str]] = []
    for node in paragraph.findall(".//m:oMath", NS):
        signatures.append({"type": "m:oMath", "text": "".join(node.xpath(".//m:t/text()", namespaces=NS))})
    for xpath, kind in [(".//w:object", "w:object"), (".//w:drawing", "w:drawing"), (".//v:imagedata", "v:imagedata")]:
        for node in paragraph.findall(xpath, NS):
            rel_id = node.get(qn(R, "id")) or node.get(qn(R, "embed")) or ""
            if not rel_id:
                values = node.xpath(".//@r:embed | .//@r:id", namespaces=NS)
                rel_id = values[0] if values else ""
            target = relationship_targets.get(rel_id, "")
            signature = {"type": kind}
            if target:
                signature["target"] = target
                blob = related_blobs.get(target)
                if blob is not None:
                    signature["sha256"] = sha256(blob).hexdigest()
            elif rel_id:
                signature["rId"] = rel_id
            signatures.append(signature)
    return signatures


def next_relationship_id(existing_ids: set[str]) -> str:
    max_id = 0
    for rel_id in existing_ids:
        if rel_id.startswith("rId") and rel_id[3:].isdigit():
            max_id = max(max_id, int(rel_id[3:]))
    return f"rId{max_id + 1}"


def _rels_path(package_root: Path, part_name: str) -> Path:
    part = Path(part_name)
    return package_root / part.parent / "_rels" / f"{part.name}.rels"


def _target_path(package_root: Path, part_name: str, target: str) -> Path:
    part_parent = Path(part_name).parent
    target_path = Path(target.lstrip("/"))
    if target.startswith("/"):
        return package_root / target_path
    return package_root / part_parent / target_path


def _unique_relationship_target(dest_root: Path, part_name: str, target: str) -> tuple[Path, str]:
    part_parent = Path(part_name).parent
    target_path = Path(target.lstrip("/"))
    package_relative = target_path if target.startswith("/") else part_parent / target_path
    candidate = package_relative
    stem_match = re.match(r"^(.*?)(\d+)$", package_relative.stem)
    base_stem = stem_match.group(1) if stem_match else package_relative.stem
    counter = int(stem_match.group(2)) if stem_match else 1
    while (dest_root / candidate).exists():
        counter += 1
        candidate = package_relative.with_name(f"{base_stem}{counter}{package_relative.suffix}")
    if target.startswith("/"):
        relationship_target = "/" + candidate.as_posix()
    else:
        relationship_target = candidate.relative_to(part_parent).as_posix()
    return dest_root / candidate, relationship_target


def copy_relationship_target(source_root: Path, dest_root: Path, part_name: str, source_rid: str) -> tuple[str, str]:
    src_rels = _rels_path(source_root, part_name)
    dst_rels = _rels_path(dest_root, part_name)
    src_tree = etree.parse(str(src_rels))
    dst_tree = etree.parse(str(dst_rels))
    src_rel = src_tree.find(f".//{{{REL_NS}}}Relationship[@Id='{source_rid}']")
    if src_rel is None:
        raise FileNotFoundError(f"relationship {source_rid} not found in {src_rels}")
    target = src_rel.get("Target")
    rel_type = src_rel.get("Type")
    target_mode = src_rel.get("TargetMode")
    if target_mode == "External":
        raise ValueError(f"relationship {source_rid} points to an external target and cannot be copied")
    if not target or not rel_type:
        raise ValueError(f"relationship {source_rid} is missing Target or Type")
    source_target = _target_path(source_root, part_name, target).resolve()
    if not source_target.exists():
        raise FileNotFoundError(f"relationship target missing: {source_target}")
    dest_target, dest_relationship_target = _unique_relationship_target(dest_root, part_name, target)
    dest_target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source_target, dest_target)
    existing_ids = {item.get("Id", "") for item in dst_tree.findall(f".//{{{REL_NS}}}Relationship")}
    new_rid = next_relationship_id(existing_ids)
    etree.SubElement(dst_tree.getroot(), f"{{{REL_NS}}}Relationship", Id=new_rid, Type=rel_type, Target=dest_relationship_target)
    dst_tree.write(str(dst_rels), encoding="UTF-8", xml_declaration=True)
    _ensure_content_type(dest_root / "[Content_Types].xml", dest_target.suffix.lstrip("."))
    return new_rid, dest_relationship_target


def _ensure_content_type(content_types_path: Path, extension: str) -> None:
    tree = etree.parse(str(content_types_path))
    root = tree.getroot()
    if root.find(f".//{{{CT_NS}}}Default[@Extension='{extension}']") is None:
        content_types = {
            "bin": "application/vnd.openxmlformats-officedocument.oleObject",
            "emf": "image/x-emf",
            "gif": "image/gif",
            "jpeg": "image/jpeg",
            "jpg": "image/jpeg",
            "png": "image/png",
            "wmf": "image/x-wmf",
        }
        content_type = content_types.get(extension.lower(), "application/octet-stream")
        etree.SubElement(root, f"{{{CT_NS}}}Default", Extension=extension, ContentType=content_type)
        tree.write(str(content_types_path), encoding="UTF-8", xml_declaration=True)
