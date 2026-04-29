#!/usr/bin/env python3
"""Extract speaker form fields from many directories.

Expected layout:
    root_dir/
      speaker-one/
        form.docx
        photo.jpg
      speaker-two/
        speaker.docx
        portrait.png

The script uses only Python's standard library. It extracts only:
- name
- job_title
- short_bio

It also records the detected DOCX file and photo file for each directory.

Usage:
    python3 extract_speaker_dirs.py /path/to/root
    python3 extract_speaker_dirs.py /path/to/root --json speakers.json
    python3 extract_speaker_dirs.py /path/to/root --csv speakers.csv
    python3 extract_speaker_dirs.py /path/to/root --recursive
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple


def load_existing(json_path: Optional[Path]) -> Dict[str, Dict[str, str]]:
    """Load an existing speakers.json, indexed by `directory`, if present."""
    if not json_path or not json_path.is_file():
        return {}
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"warning: could not read existing JSON {json_path}: {exc}",
              file=sys.stderr)
        return {}
    return {row["directory"]: row for row in data if isinstance(row, dict) and row.get("directory")}

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

PHOTO_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".gif", ".webp", ".tif", ".tiff", ".bmp", ".heic"
}

LABEL_ALIASES = {
    "name": "name",
    "name (as wanted in the program)": "name",
    "job title": "job_title",
    "job title (as wanted in the program)": "job_title",
    "short bio": "short_bio",
    "short bio (max 12 lines)": "short_bio",
}

OUTPUT_FIELDS = ("name", "job_title", "short_bio")


def clean(text: str) -> str:
    text = text.replace("\xa0", " ").replace("\u200b", "")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line).strip()


def norm_label(label: str) -> str:
    label = clean(label).strip(" :\t").lower()
    label = re.sub(r"\s+", " ", label)
    return label


def canonical_label(label: str) -> Optional[str]:
    return LABEL_ALIASES.get(norm_label(label))


def paragraph_text(p: ET.Element) -> str:
    parts: List[str] = []
    for node in p.iter():
        if node.tag == f"{{{NS['w']}}}t" and node.text:
            parts.append(node.text)
        elif node.tag == f"{{{NS['w']}}}tab":
            parts.append("\t")
        elif node.tag == f"{{{NS['w']}}}br":
            parts.append("\n")
    return clean("".join(parts))


def cell_text(tc: ET.Element) -> str:
    return clean("\n".join(paragraph_text(p) for p in tc.findall(".//w:p", NS)))


def load_document_xml(docx_path: Path) -> ET.Element:
    with zipfile.ZipFile(docx_path) as zf:
        xml_bytes = zf.read("word/document.xml")
    return ET.fromstring(xml_bytes)


def iter_table_rows(root: ET.Element) -> Iterable[List[str]]:
    for tr in root.findall(".//w:tr", NS):
        row = [cell_text(tc) for tc in tr.findall("./w:tc", NS)]
        yield [c for c in row if c]


def iter_body_paragraphs(root: ET.Element) -> Iterable[str]:
    body = root.find("w:body", NS)
    if body is None:
        return
    for child in body:
        if child.tag == f"{{{NS['w']}}}p":
            text = paragraph_text(child)
            if text:
                yield text


def split_label_value(text: str) -> Tuple[Optional[str], Optional[str]]:
    text = clean(text)
    if not text:
        return None, None

    if ":" in text:
        label, value = text.split(":", 1)
        key = canonical_label(label)
        if key:
            return key, clean(value)

    dotted = re.split(r"\.{5,}", text, maxsplit=1)
    if len(dotted) == 2:
        key = canonical_label(dotted[0])
        if key:
            return key, clean(dotted[1])

    return None, None


def value_from_row(cells: List[str]) -> Tuple[Optional[str], Optional[str]]:
    if not cells:
        return None, None

    for c in cells:
        key, value = split_label_value(c)
        if key:
            return key, value

    for i, cell in enumerate(cells):
        key = canonical_label(cell)
        if not key:
            continue
        value_parts: List[str] = []
        for following in cells[i + 1:]:
            f = clean(following).strip()
            if not f or f == ":":
                continue
            if canonical_label(f):
                break
            value_parts.append(f.lstrip(":").strip())
        return key, clean("\n".join(value_parts))

    joined = clean("\n".join(cells))
    return split_label_value(joined)


def extract_fields(docx_path: Path) -> Dict[str, str]:
    root = load_document_xml(docx_path)
    fields: Dict[str, str] = {}
    all_chunks: List[str] = []

    for cells in iter_table_rows(root):
        all_chunks.extend(cells)
        key, value = value_from_row(cells)
        if key:
            fields[key] = value or ""

    for para in iter_body_paragraphs(root):
        all_chunks.append(para)
        key, value = split_label_value(para)
        if key and key not in fields:
            fields[key] = value or ""

    full_text = clean("\n".join(all_chunks))
    fallback_patterns = {
        "name": r"Name\s*\(as wanted in the program\)\s*:?\s*(.+?)(?:\n\s*Job title|$)",
        "job_title": r"Job title\s*\(as wanted in the program\)\s*:?\s*(.+?)(?:\n\s*Personal mobile|\n\s*Short BIO|$)",
        "short_bio": r"Short\s+BIO\s*\(max 12 lines\)\s*:?\s*(.+?)(?:\n\s*Can you attach|\n\s*Participation in|$)",
    }

    for key, pattern in fallback_patterns.items():
        if not fields.get(key):
            match = re.search(pattern, full_text, flags=re.IGNORECASE | re.DOTALL)
            if match:
                value = clean(match.group(1)).strip(" :")
                if value:
                    fields[key] = value

    return {key: fields.get(key, "") for key in OUTPUT_FIELDS}


def candidate_directories(root: Path, recursive: bool) -> Iterable[Path]:
    if recursive:
        for path in sorted(root.rglob("*")):
            if path.is_dir():
                yield path
    else:
        for path in sorted(root.iterdir()):
            if path.is_dir():
                yield path


def find_one_file(directory: Path, extensions: set[str]) -> Tuple[Optional[Path], List[Path]]:
    matches = sorted(
        p for p in directory.iterdir()
        if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() in extensions
    )
    return (matches[0] if matches else None), matches


def extract_directory(
    directory: Path,
    root: Path,
    existing: Optional[Dict[str, str]] = None,
) -> Dict[str, str]:
    docx_path, docx_matches = find_one_file(directory, {".docx"})
    photo_path, photo_matches = find_one_file(directory, PHOTO_EXTENSIONS)

    row: Dict[str, str] = {
        "directory": str(directory.relative_to(root)),
        "docx": str(docx_path.relative_to(root)) if docx_path else "",
        "photo": str(photo_path.relative_to(root)) if photo_path else "",
        "status": "ok",
        "notes": "",
        "name": "",
        "job_title": "",
        "short_bio": "",
    }

    notes: List[str] = []
    if len(docx_matches) > 1:
        notes.append(f"multiple DOCX files found; used {docx_path.name}")
    if len(photo_matches) > 1:
        notes.append(f"multiple photo files found; used {photo_path.name}")
    if not docx_path:
        row["status"] = "missing_docx"
        row["notes"] = "no DOCX file found"
        if existing:
            for field in OUTPUT_FIELDS:
                if existing.get(field):
                    row[field] = existing[field]
        return row
    if not photo_path:
        notes.append("no photo file found")

    try:
        row.update(extract_fields(docx_path))
    except Exception as exc:  # noqa: BLE001 - CLI should report and continue.
        row["status"] = "error"
        notes.append(f"could not read DOCX: {exc}")

    if existing:
        for field in OUTPUT_FIELDS:
            existing_value = (existing.get(field) or "").strip()
            if existing_value:
                row[field] = existing_value

    missing_fields = [field for field in OUTPUT_FIELDS if not row.get(field)]
    if row["status"] not in ("error",):
        row["status"] = "partial" if missing_fields else "ok"
    if missing_fields:
        notes.append("missing fields: " + ", ".join(missing_fields))

    row["notes"] = "; ".join(notes)
    return row


def write_csv(path: Path, rows: List[Dict[str, str]]) -> None:
    fieldnames = ["directory", "docx", "photo", "name", "job_title", "short_bio", "status", "notes"]
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract name, job title, short BIO, and photo path from speaker directories."
    )
    parser.add_argument("root", help="Root directory containing one subdirectory per speaker")
    parser.add_argument("--recursive", action="store_true", help="Scan all nested directories, not only direct children")
    parser.add_argument("--json", dest="json_path", help="Optional path to save results as JSON")
    parser.add_argument("--csv", dest="csv_path", help="Optional path to save results as CSV")
    args = parser.parse_args()

    root = Path(args.root).expanduser().resolve()
    if not root.is_dir():
        print(f"error: root is not a directory: {root}", file=sys.stderr)
        return 2

    json_path = Path(args.json_path).expanduser().resolve() if args.json_path else None
    existing_by_dir = load_existing(json_path)

    rows = []
    for directory in candidate_directories(root, args.recursive):
        # Skip directories that contain neither a DOCX nor a photo.
        has_docx = any(p.is_file() and p.suffix.lower() == ".docx" for p in directory.iterdir())
        has_photo = any(p.is_file() and p.suffix.lower() in PHOTO_EXTENSIONS for p in directory.iterdir())
        if not has_docx and not has_photo:
            continue
        rel_dir = str(directory.relative_to(root))
        rows.append(extract_directory(directory, root, existing_by_dir.get(rel_dir)))

    output = json.dumps(rows, ensure_ascii=False, indent=2)
    print(output)

    if json_path:
        json_path.write_text(output + "\n", encoding="utf-8")
    if args.csv_path:
        write_csv(Path(args.csv_path), rows)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
