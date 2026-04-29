#!/usr/bin/env python3
"""Render speaker HTML blocks from a speakers.json file using a template.

Reads a JSON list of speakers (as produced by extract_speaker_dirs.py) and
renders one HTML block per speaker by substituting fields into a template.

Usage:
    python3 render_speakers.py
    python3 render_speakers.py --input scripts/2026/output/speakers.json \\
                               --template scripts/speaker_template.html \\
                               --output scripts/2026/output/speakers.html \\
                               --source-images "/home/user/Downloads/NISDUC Conference" \\
                               --images-dir scripts/2026/output/images \\
                               --image-url-prefix images/ \\
                               --year 26
"""

from __future__ import annotations

import argparse
import html
import json
import shutil
import sys
from pathlib import Path
from typing import Dict, List, Optional

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT = REPO_ROOT / "scripts" / "2026" / "output" / "speakers.json"
DEFAULT_TEMPLATE = REPO_ROOT / "scripts" / "speaker_template.html"
DEFAULT_OUTPUT = REPO_ROOT / "scripts" / "2026" / "output" / "speakers.html"
DEFAULT_SOURCE_IMAGES = Path("/home/user/Downloads/NISDUC Conference")
DEFAULT_IMAGES_DIR = REPO_ROOT / "scripts" / "2026" / "output" / "images"
DEFAULT_IMAGE_URL_PREFIX = "images/"
DEFAULT_YEAR = "26"
DEFAULT_WIDTH = "1280"
DEFAULT_HEIGHT = "984"


def render_paragraphs(speaker: Dict[str, str]) -> str:
    """Build the inner <p> block for a speaker.

    The job title becomes the lead paragraphs (split on newlines so multi-line
    titles render as separate <p> tags). The short bio, when present, is added
    as additional paragraphs.
    """
    indent = " " * 28
    lines: List[str] = []

    for source in (speaker.get("job_title", ""), speaker.get("short_bio", "")):
        for raw in source.splitlines():
            text = raw.strip()
            if text:
                lines.append(f"{indent}<p>{html.escape(text)}</p>")

    return "\n".join(lines)


def copy_speaker_image(
    speaker: Dict[str, str],
    speaker_id: str,
    source_root: Path,
    dest_dir: Path,
) -> Optional[str]:
    """Copy the speaker's photo to dest_dir as speaker<id>.<ext>.

    Returns the destination filename (e.g. "speaker2601.jpg") or None when no
    source photo is available.
    """
    photo_rel = speaker.get("photo", "").strip()
    if not photo_rel:
        return None

    source_path = source_root / photo_rel
    if not source_path.is_file():
        print(f"warning: missing source image for {speaker_id}: {source_path}",
              file=sys.stderr)
        return None

    dest_dir.mkdir(parents=True, exist_ok=True)
    extension = source_path.suffix.lower()
    dest_name = f"speaker{speaker_id[1:]}{extension}"
    shutil.copy2(source_path, dest_dir / dest_name)
    return dest_name


def render_speaker(
    template: str,
    speaker: Dict[str, str],
    speaker_id: str,
    photo_url: str,
) -> str:
    name = speaker.get("name", "").strip()

    return (
        template
        .replace("{{ID}}", speaker_id)
        .replace("{{PHOTO}}", html.escape(photo_url, quote=True))
        .replace("{{NAME}}", html.escape(name))
        .replace("{{WIDTH}}", DEFAULT_WIDTH)
        .replace("{{HEIGHT}}", DEFAULT_HEIGHT)
        .replace("{{PARAGRAPHS}}", render_paragraphs(speaker))
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--source-images", type=Path, default=DEFAULT_SOURCE_IMAGES,
                        help="Directory containing original speaker subdirectories with photos")
    parser.add_argument("--images-dir", type=Path, default=DEFAULT_IMAGES_DIR,
                        help="Destination directory for copied speaker images")
    parser.add_argument("--image-url-prefix", default=DEFAULT_IMAGE_URL_PREFIX,
                        help="URL prefix used when referencing images in the rendered HTML")
    parser.add_argument("--year", default=DEFAULT_YEAR,
                        help="Two-digit year used as id prefix (e.g. 26 -> s2601, s2602...)")
    args = parser.parse_args()

    speakers = json.loads(args.input.read_text(encoding="utf-8"))
    template = args.template.read_text(encoding="utf-8")

    blocks: List[str] = []
    rendered = 0
    skipped: List[str] = []

    for index, speaker in enumerate(speakers, start=1):
        if not speaker.get("name", "").strip():
            skipped.append(speaker.get("directory", f"index {index}"))
            continue

        rendered += 1
        speaker_id = f"s{args.year}{rendered:02d}"
        copied = copy_speaker_image(speaker, speaker_id, args.source_images, args.images_dir)
        photo_url = f"{args.image_url_prefix}{copied}" if copied else ""
        blocks.append(render_speaker(template, speaker, speaker_id, photo_url))

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text("\n".join(blocks) + "\n", encoding="utf-8")

    print(f"Rendered {rendered} speakers to {args.output}", file=sys.stderr)
    if skipped:
        print(f"Skipped {len(skipped)} entries with no name: {', '.join(skipped)}",
              file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
