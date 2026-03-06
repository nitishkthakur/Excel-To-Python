#!/usr/bin/env python3
"""Hardcode level-1 predecessor sheets in an XLSX workbook.

Given a target worksheet, this script:
1) Finds direct predecessor sheets used by target formulas.
2) Replaces formulas in predecessor sheets with cached values.
3) Removes every other sheet from the workbook.
4) Preserves formulas in the target sheet.
5) Keeps workbook artifacts (styles, defined names, formatting) intact where possible.
"""

from __future__ import annotations

import argparse
import posixpath
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from lxml import etree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

NSMAP = {
    "x": NS_MAIN,
    "r": NS_REL,
    "pr": NS_PKG_REL,
}

WORKSHEET_REL_TYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet",
}

CALCCHAIN_REL_TYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain",
}

TARGET_ERROR_VALUES = {"#VALUE!", "#NAME?"}

STRING_LITERAL_RE = re.compile(r'"(?:[^"]|"")*"')
QUOTED_SHEET_REF_RE = re.compile(r"'((?:[^']|'')+)'!")
UNQUOTED_SHEET_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.\\])([A-Za-z0-9_.]+(?::[A-Za-z0-9_.]+)?)!"
)
TOKEN_RE = re.compile(r"\b[A-Za-z_\\][A-Za-z0-9_.\\]*\b")


@dataclass(frozen=True)
class SheetRecord:
    name: str
    old_index: int
    rel_id: Optional[str]
    rel_type: Optional[str]
    part_path: Optional[str]
    is_worksheet: bool


@dataclass(frozen=True)
class DefinedNameRecord:
    name_lower: str
    formula: str
    local_sheet_id: Optional[int]


def parse_xml(data: bytes) -> etree._Element:
    parser = etree.XMLParser(remove_blank_text=False)
    return etree.fromstring(data, parser=parser)


def xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=None)


def normalize_rel_target(base_part: str, target: str) -> str:
    """Resolve a relationship target path to an xl/... package path."""
    base_dir = posixpath.dirname(base_part)
    resolved = posixpath.normpath(posixpath.join(base_dir, target))
    if resolved.startswith("/"):
        resolved = resolved[1:]
    return resolved


def is_external_ref(token: str) -> bool:
    token = token.strip()
    return token.startswith("[") or "]" in token


def strip_string_literals(formula: str) -> str:
    return STRING_LITERAL_RE.sub("", formula)


def expand_sheet_range(
    start: str,
    end: str,
    sheet_order: List[str],
    sheet_positions: Dict[str, int],
    sheet_set: Set[str],
) -> Set[str]:
    if start not in sheet_set or end not in sheet_set:
        return set()
    i = sheet_positions[start]
    j = sheet_positions[end]
    if i <= j:
        return set(sheet_order[i : j + 1])
    return set(sheet_order[j : i + 1])


def extract_sheet_refs(
    formula: str,
    sheet_order: List[str],
    sheet_positions: Dict[str, int],
    sheet_set: Set[str],
) -> Set[str]:
    refs: Set[str] = set()
    no_strings = strip_string_literals(formula)

    for match in QUOTED_SHEET_REF_RE.finditer(no_strings):
        token = match.group(1).replace("''", "'")
        if is_external_ref(token):
            continue
        if ":" in token:
            start, end = token.split(":", 1)
            refs.update(expand_sheet_range(start, end, sheet_order, sheet_positions, sheet_set))
        elif token in sheet_set:
            refs.add(token)

    without_quoted = QUOTED_SHEET_REF_RE.sub("", no_strings)

    for match in UNQUOTED_SHEET_REF_RE.finditer(without_quoted):
        token = match.group(1)
        if is_external_ref(token):
            continue
        if ":" in token:
            start, end = token.split(":", 1)
            refs.update(expand_sheet_range(start, end, sheet_order, sheet_positions, sheet_set))
        elif token in sheet_set:
            refs.add(token)

    return refs


def extract_name_tokens(formula: str, candidate_names: Set[str]) -> Set[str]:
    no_strings = strip_string_literals(formula)
    no_sheet_refs = QUOTED_SHEET_REF_RE.sub("", no_strings)
    no_sheet_refs = UNQUOTED_SHEET_REF_RE.sub("", no_sheet_refs)

    out: Set[str] = set()
    for tok in TOKEN_RE.findall(no_sheet_refs):
        low = tok.lower()
        if low in candidate_names:
            out.add(low)
    return out


def parse_defined_names(workbook_root: etree._Element) -> Tuple[List[DefinedNameRecord], Dict[Tuple[int, str], DefinedNameRecord], Dict[str, DefinedNameRecord]]:
    all_records: List[DefinedNameRecord] = []
    local_map: Dict[Tuple[int, str], DefinedNameRecord] = {}
    global_map: Dict[str, DefinedNameRecord] = {}

    for dn in workbook_root.xpath("./x:definedNames/x:definedName", namespaces=NSMAP):
        name = dn.get("name")
        if not name:
            continue

        formula = (dn.text or "").strip()
        local_attr = dn.get("localSheetId")
        local_id = int(local_attr) if local_attr is not None and local_attr.isdigit() else None

        rec = DefinedNameRecord(name_lower=name.lower(), formula=formula, local_sheet_id=local_id)
        all_records.append(rec)

        if local_id is None:
            global_map.setdefault(rec.name_lower, rec)
        else:
            local_map.setdefault((local_id, rec.name_lower), rec)

    return all_records, local_map, global_map


def resolve_defined_name(
    name_lower: str,
    context_sheet_id: int,
    local_map: Dict[Tuple[int, str], DefinedNameRecord],
    global_map: Dict[str, DefinedNameRecord],
) -> Optional[DefinedNameRecord]:
    return local_map.get((context_sheet_id, name_lower)) or global_map.get(name_lower)


def gather_formulas(sheet_xml: bytes) -> List[str]:
    root = parse_xml(sheet_xml)
    formulas: List[str] = []
    for f_node in root.xpath(".//x:f", namespaces=NSMAP):
        text = (f_node.text or "").strip()
        if text:
            formulas.append(text)
    return formulas


def strip_formulas_keep_cached_values(sheet_xml: bytes) -> Tuple[bytes, int]:
    root = parse_xml(sheet_xml)
    formula_nodes = root.xpath(".//x:c/x:f", namespaces=NSMAP)
    changed = 0
    for f_node in formula_nodes:
        parent = f_node.getparent()
        if parent is None:
            continue
        parent.remove(f_node)
        changed += 1
    return xml_bytes(root), changed


def compute_level1_predecessors(
    target_formulas: Iterable[str],
    target_sheet_index: int,
    sheet_order: List[str],
    worksheet_names: Set[str],
    local_names: Dict[Tuple[int, str], DefinedNameRecord],
    global_names: Dict[str, DefinedNameRecord],
) -> Set[str]:
    predecessors: Set[str] = set()
    sheet_positions = {name: idx for idx, name in enumerate(sheet_order)}

    candidate_names = set(global_names.keys()) | {name for (_, name) in local_names.keys()}

    # Direct sheet references from target formulas.
    for formula in target_formulas:
        refs = extract_sheet_refs(formula, sheet_order, sheet_positions, worksheet_names)
        predecessors.update(refs)

    # Resolve defined names referenced by target formulas (recursive).
    queue: List[Tuple[str, int]] = []
    visited: Set[Tuple[str, int]] = set()

    for formula in target_formulas:
        for nm in extract_name_tokens(formula, candidate_names):
            queue.append((nm, target_sheet_index))

    while queue:
        nm, context_idx = queue.pop()
        key = (nm, context_idx)
        if key in visited:
            continue
        visited.add(key)

        rec = resolve_defined_name(nm, context_idx, local_names, global_names)
        if rec is None or not rec.formula:
            continue

        refs = extract_sheet_refs(rec.formula, sheet_order, sheet_positions, worksheet_names)
        predecessors.update(refs)

        next_context = rec.local_sheet_id if rec.local_sheet_id is not None else context_idx
        for nested in extract_name_tokens(rec.formula, candidate_names):
            queue.append((nested, next_context))

    return predecessors


def remap_workbook_sheets_and_defined_names(
    workbook_root: etree._Element,
    keep_sheet_names: Set[str],
    old_index_to_new_index: Dict[int, int],
) -> None:
    sheets_parent = workbook_root.find(f"{{{NS_MAIN}}}sheets")
    if sheets_parent is None:
        raise ValueError("workbook.xml is missing <sheets>.")

    for sheet_elem in list(sheets_parent):
        name = sheet_elem.get("name")
        if name not in keep_sheet_names:
            sheets_parent.remove(sheet_elem)

    defined_names_parent = workbook_root.find(f"{{{NS_MAIN}}}definedNames")
    if defined_names_parent is None:
        return

    for dn in list(defined_names_parent):
        local_attr = dn.get("localSheetId")
        if local_attr is None or not local_attr.isdigit():
            continue
        old_local_idx = int(local_attr)
        if old_local_idx not in old_index_to_new_index:
            defined_names_parent.remove(dn)
            continue
        dn.set("localSheetId", str(old_index_to_new_index[old_local_idx]))

    if len(defined_names_parent) == 0:
        workbook_root.remove(defined_names_parent)


def prune_workbook_relationships(
    rels_root: etree._Element,
    keep_rel_ids: Set[str],
    removed_sheet_rel_ids: Set[str],
    removed_calcchain_part_paths: Set[str],
    workbook_part: str,
) -> None:
    for rel in list(rels_root):
        rid = rel.get("Id")
        rel_type = rel.get("Type")

        if rel_type in CALCCHAIN_REL_TYPES:
            target = rel.get("Target")
            if target:
                removed_calcchain_part_paths.add(normalize_rel_target(workbook_part, target))
            rels_root.remove(rel)
            continue

        if rid in removed_sheet_rel_ids:
            rels_root.remove(rel)
            continue

        if rel_type in WORKSHEET_REL_TYPES and rid not in keep_rel_ids:
            rels_root.remove(rel)


def prune_content_types(content_types_root: etree._Element, removed_part_paths: Set[str]) -> None:
    for override in list(content_types_root):
        if not override.tag.endswith("Override"):
            continue
        part_name = override.get("PartName")
        if not part_name:
            continue
        normalized = part_name[1:] if part_name.startswith("/") else part_name
        if normalized in removed_part_paths:
            content_types_root.remove(override)


def prune_sheet_relationship_file(
    rels_xml: bytes,
    base_sheet_part: str,
    removed_part_paths: Set[str],
) -> bytes:
    root = parse_xml(rels_xml)
    for rel in list(root):
        target_mode = rel.get("TargetMode")
        if target_mode and target_mode.lower() == "external":
            continue
        target = rel.get("Target")
        if not target:
            continue
        target_path = normalize_rel_target(base_sheet_part, target)
        if target_path in removed_part_paths:
            root.remove(rel)
    return xml_bytes(root)


def collect_sheet_records(workbook_root: etree._Element, workbook_rels_root: etree._Element) -> List[SheetRecord]:
    rel_by_id: Dict[str, etree._Element] = {
        rel.get("Id"): rel
        for rel in workbook_rels_root.findall(f"{{{NS_PKG_REL}}}Relationship")
        if rel.get("Id")
    }

    sheets_parent = workbook_root.find(f"{{{NS_MAIN}}}sheets")
    if sheets_parent is None:
        raise ValueError("workbook.xml has no <sheets> element.")

    sheet_records: List[SheetRecord] = []
    for idx, sheet_elem in enumerate(list(sheets_parent)):
        name = sheet_elem.get("name")
        rel_id = sheet_elem.get(f"{{{NS_REL}}}id")

        rel_type = None
        part_path = None
        is_worksheet = False

        if rel_id and rel_id in rel_by_id:
            rel = rel_by_id[rel_id]
            rel_type = rel.get("Type")
            target = rel.get("Target")
            if target:
                part_path = normalize_rel_target("xl/workbook.xml", target)
            is_worksheet = rel_type in WORKSHEET_REL_TYPES

        if not name:
            name = f"__unnamed_{idx}"

        sheet_records.append(
            SheetRecord(
                name=name,
                old_index=idx,
                rel_id=rel_id,
                rel_type=rel_type,
                part_path=part_path,
                is_worksheet=is_worksheet,
            )
        )

    return sheet_records


def build_old_to_new_index_map(sheet_records: List[SheetRecord], keep_names: Set[str]) -> Dict[int, int]:
    mapping: Dict[int, int] = {}
    next_idx = 0
    for rec in sheet_records:
        if rec.name in keep_names:
            mapping[rec.old_index] = next_idx
            next_idx += 1
    return mapping


def transform_workbook(input_xlsx: Path, output_xlsx: Path, target_sheet: str, fail_on_target_errors: bool) -> Dict[str, object]:
    workbook_part = "xl/workbook.xml"
    workbook_rels_part = "xl/_rels/workbook.xml.rels"
    content_types_part = "[Content_Types].xml"

    replacements: Dict[str, bytes] = {}
    skipped_paths: Set[str] = set()

    with zipfile.ZipFile(input_xlsx, "r") as zin:
        names = set(zin.namelist())

        missing = [p for p in [workbook_part, workbook_rels_part, content_types_part] if p not in names]
        if missing:
            raise ValueError(f"Input is missing required XLSX parts: {missing}")

        workbook_root = parse_xml(zin.read(workbook_part))
        workbook_rels_root = parse_xml(zin.read(workbook_rels_part))
        content_types_root = parse_xml(zin.read(content_types_part))

        sheet_records = collect_sheet_records(workbook_root, workbook_rels_root)
        sheet_by_name = {s.name: s for s in sheet_records}

        if target_sheet not in sheet_by_name:
            known = ", ".join(sorted(sheet_by_name.keys()))
            raise ValueError(f"Target sheet '{target_sheet}' not found. Available sheets: {known}")

        target_rec = sheet_by_name[target_sheet]
        if not target_rec.is_worksheet or not target_rec.part_path:
            raise ValueError(f"Target sheet '{target_sheet}' is not a worksheet with a resolvable XML part.")

        if target_rec.part_path not in names:
            raise ValueError(f"Target sheet XML part not found in package: {target_rec.part_path}")

        sheet_order = [s.name for s in sheet_records]
        worksheet_names = {s.name for s in sheet_records if s.is_worksheet}

        _, local_name_map, global_name_map = parse_defined_names(workbook_root)

        target_formulas = gather_formulas(zin.read(target_rec.part_path))
        predecessors = compute_level1_predecessors(
            target_formulas=target_formulas,
            target_sheet_index=target_rec.old_index,
            sheet_order=sheet_order,
            worksheet_names=worksheet_names,
            local_names=local_name_map,
            global_names=global_name_map,
        )

        predecessors.discard(target_sheet)

        predecessor_records = [sheet_by_name[name] for name in sorted(predecessors) if name in sheet_by_name]
        predecessor_records = [rec for rec in predecessor_records if rec.is_worksheet and rec.part_path]

        keep_sheet_names = {target_sheet} | {rec.name for rec in predecessor_records}
        old_to_new_idx = build_old_to_new_index_map(sheet_records, keep_sheet_names)

        removed_sheet_records = [rec for rec in sheet_records if rec.name not in keep_sheet_names]
        removed_sheet_paths = {rec.part_path for rec in removed_sheet_records if rec.part_path}

        # Hardcode formulas in level-1 predecessors.
        formulas_hardcoded = 0
        for rec in predecessor_records:
            part = rec.part_path
            if not part or part not in names:
                continue
            new_xml, changed = strip_formulas_keep_cached_values(zin.read(part))
            replacements[part] = new_xml
            formulas_hardcoded += changed

        # Update workbook metadata.
        remap_workbook_sheets_and_defined_names(workbook_root, keep_sheet_names, old_to_new_idx)

        keep_rel_ids = {sheet_by_name[n].rel_id for n in keep_sheet_names if sheet_by_name[n].rel_id}
        removed_sheet_rel_ids = {rec.rel_id for rec in removed_sheet_records if rec.rel_id}
        removed_calcchain_part_paths: Set[str] = set()
        prune_workbook_relationships(
            workbook_rels_root,
            keep_rel_ids,
            removed_sheet_rel_ids,
            removed_calcchain_part_paths,
            workbook_part,
        )

        removed_part_paths = set(removed_sheet_paths)
        removed_part_paths.update(removed_calcchain_part_paths)

        # Skip removed sheet XMLs and their rel files.
        for part in list(removed_sheet_paths):
            skipped_paths.add(part)
            dirname = posixpath.dirname(part)
            basename = posixpath.basename(part)
            rels_part = posixpath.join(dirname, "_rels", f"{basename}.rels")
            if rels_part in names:
                skipped_paths.add(rels_part)

        for calc_part in removed_calcchain_part_paths:
            if calc_part in names:
                skipped_paths.add(calc_part)

        # Prune kept sheet relationship files that directly target removed parts.
        for rec in sheet_records:
            if rec.name not in keep_sheet_names:
                continue
            if not rec.part_path:
                continue
            rels_part = posixpath.join(
                posixpath.dirname(rec.part_path),
                "_rels",
                f"{posixpath.basename(rec.part_path)}.rels",
            )
            if rels_part not in names:
                continue
            rels_bytes = zin.read(rels_part)
            pruned = prune_sheet_relationship_file(rels_bytes, rec.part_path, removed_part_paths)
            replacements[rels_part] = pruned

        # Content types cleanup.
        prune_content_types(content_types_root, removed_part_paths)

        replacements[workbook_part] = xml_bytes(workbook_root)
        replacements[workbook_rels_part] = xml_bytes(workbook_rels_root)
        replacements[content_types_part] = xml_bytes(content_types_root)

        # Write output.
        with zipfile.ZipFile(output_xlsx, "w") as zout:
            for info in zin.infolist():
                if info.filename in skipped_paths:
                    continue
                data = replacements.get(info.filename)
                if data is None:
                    data = zin.read(info.filename)
                zout.writestr(info, data)

    errors = scan_target_cached_errors(output_xlsx, target_sheet)
    if fail_on_target_errors and errors:
        details = ", ".join([f"{cell}={val}" for cell, val in errors[:10]])
        raise ValueError(
            "Target sheet contains cached error values after transformation: "
            f"{details}"
        )

    return {
        "target_sheet": target_sheet,
        "predecessors": sorted({rec.name for rec in predecessor_records}),
        "kept_sheets": sorted(keep_sheet_names),
        "removed_sheets": sorted({rec.name for rec in removed_sheet_records}),
        "formulas_hardcoded": formulas_hardcoded,
        "target_error_cells": errors,
        "output": str(output_xlsx),
    }


def scan_target_cached_errors(xlsx_path: Path, target_sheet: str) -> List[Tuple[str, str]]:
    workbook_part = "xl/workbook.xml"
    workbook_rels_part = "xl/_rels/workbook.xml.rels"

    with zipfile.ZipFile(xlsx_path, "r") as zf:
        workbook_root = parse_xml(zf.read(workbook_part))
        workbook_rels_root = parse_xml(zf.read(workbook_rels_part))

        sheet_records = collect_sheet_records(workbook_root, workbook_rels_root)
        by_name = {s.name: s for s in sheet_records}
        if target_sheet not in by_name:
            return []
        target_rec = by_name[target_sheet]
        if not target_rec.part_path or target_rec.part_path not in zf.namelist():
            return []

        root = parse_xml(zf.read(target_rec.part_path))
        bad: List[Tuple[str, str]] = []
        for cell in root.xpath(".//x:c[x:f]", namespaces=NSMAP):
            ref = cell.get("r", "?")
            v = cell.find(f"{{{NS_MAIN}}}v")
            if v is None or v.text is None:
                continue
            value = v.text.strip()
            if value in TARGET_ERROR_VALUES:
                bad.append((ref, value))
        return bad


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}.level1_hardcoded{input_path.suffix}")


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Hardcode formulas in direct predecessor sheets of a target worksheet, "
            "preserve target formulas, and prune non-relevant sheets."
        )
    )
    parser.add_argument("input", type=Path, help="Path to input .xlsx file")
    parser.add_argument("--target", required=True, help="Target worksheet name")
    parser.add_argument("--output", type=Path, help="Path to output .xlsx file")
    parser.add_argument(
        "--allow-target-errors",
        action="store_true",
        help="Allow output even if target cached formula values contain #VALUE! or #NAME?",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    input_path: Path = args.input
    output_path: Path = args.output or default_output_path(input_path)

    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 2

    try:
        result = transform_workbook(
            input_xlsx=input_path,
            output_xlsx=output_path,
            target_sheet=args.target,
            fail_on_target_errors=not args.allow_target_errors,
        )
    except Exception as exc:
        print(f"Transformation failed: {exc}", file=sys.stderr)
        return 1

    print("Transformation complete")
    print(f"Output: {result['output']}")
    print(f"Target: {result['target_sheet']}")
    print(f"Level1 predecessors: {', '.join(result['predecessors']) or '(none)'}")
    print(f"Kept sheets: {', '.join(result['kept_sheets'])}")
    print(f"Removed sheets: {', '.join(result['removed_sheets']) or '(none)'}")
    print(f"Formulas hardcoded in predecessors: {result['formulas_hardcoded']}")

    target_errors = result["target_error_cells"]
    if target_errors:
        preview = ", ".join([f"{cell}={val}" for cell, val in target_errors[:10]])
        print(f"Target cached error cells: {preview}")
    else:
        print("Target cached error cells: none (#VALUE!/#NAME?)")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
