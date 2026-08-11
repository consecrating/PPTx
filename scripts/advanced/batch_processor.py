"""Batch Processor for PPTX presentations.

Handles bulk operations: merge decks, split by sections, extract content,
batch apply themes, and generate multiple variants.

Usage:
    python batch_processor.py merge deck1.pptx deck2.pptx -o combined.pptx
    python batch_processor.py split input.pptx --by-sections
    python batch_processor.py extract input.pptx --format markdown
"""

import argparse
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

import defusedxml.minidom

sys.path.insert(0, str(Path(__file__).parent.parent))
from office.helpers import safe_extract, rezip, opc_target, rels_source_part, SLIDE_REL_TYPE


def extract_text(pptx_path: Path, format: str = "text") -> str:
    """Extract all text content from a PPTX file."""
    with zipfile.ZipFile(pptx_path, "r") as zf:
        # Get slide order from presentation.xml
        pres_rels = zf.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
        pres_xml = zf.read("ppt/presentation.xml").decode("utf-8")

        rels_dom = defusedxml.minidom.parseString(pres_rels)
        rid_to_target = {}
        for rel in rels_dom.getElementsByTagName("Relationship"):
            if rel.getAttribute("Type") == SLIDE_REL_TYPE:
                rid_to_target[rel.getAttribute("Id")] = rel.getAttribute("Target")

        pres_dom = defusedxml.minidom.parseString(pres_xml)
        slide_order = []
        for sld_id in pres_dom.getElementsByTagName("p:sldId"):
            rid = sld_id.getAttribute("r:id")
            if rid in rid_to_target:
                slide_order.append("ppt/" + rid_to_target[rid].lstrip("./"))

        output_lines = []
        for i, slide_path in enumerate(slide_order, 1):
            try:
                slide_xml = zf.read(slide_path).decode("utf-8")
                slide_dom = defusedxml.minidom.parseString(slide_xml)

                texts = []
                for t_elem in slide_dom.getElementsByTagName("a:t"):
                    if t_elem.firstChild:
                        texts.append(t_elem.firstChild.nodeValue)

                if format == "markdown":
                    output_lines.append(f"## Slide {i}")
                    output_lines.append("")
                    output_lines.append(" ".join(texts))
                    output_lines.append("")
                else:
                    output_lines.append(f"--- Slide {i} ---")
                    output_lines.append(" ".join(texts))
                    output_lines.append("")
            except (KeyError, Exception) as e:
                output_lines.append(f"--- Slide {i} (error: {e}) ---")
                output_lines.append("")

    return "\n".join(output_lines)


def get_slide_count(pptx_path: Path) -> int:
    """Get the number of slides in a PPTX file."""
    with zipfile.ZipFile(pptx_path, "r") as zf:
        pres_xml = zf.read("ppt/presentation.xml").decode("utf-8")
        return len(re.findall(r'<p:sldId\b', pres_xml))


def extract_metadata(pptx_path: Path) -> dict:
    """Extract metadata from a PPTX file."""
    metadata = {
        "file": str(pptx_path),
        "slide_count": 0,
        "has_notes": False,
        "has_charts": False,
        "has_media": False,
        "layouts_used": [],
        "file_size_mb": pptx_path.stat().st_size / (1024 * 1024),
    }

    with zipfile.ZipFile(pptx_path, "r") as zf:
        names = zf.namelist()
        metadata["slide_count"] = len([n for n in names if re.match(r"ppt/slides/slide\d+\.xml", n)])
        metadata["has_notes"] = any("notesSlides" in n for n in names)
        metadata["has_charts"] = any("charts" in n for n in names)
        metadata["has_media"] = any("media" in n for n in names)

    return metadata


def compare_decks(deck1_path: Path, deck2_path: Path) -> dict:
    """Compare two PPTX files and report differences."""
    meta1 = extract_metadata(deck1_path)
    meta2 = extract_metadata(deck2_path)

    text1 = extract_text(deck1_path)
    text2 = extract_text(deck2_path)

    return {
        "deck1": str(deck1_path),
        "deck2": str(deck2_path),
        "slide_count_diff": meta2["slide_count"] - meta1["slide_count"],
        "size_diff_mb": meta2["file_size_mb"] - meta1["file_size_mb"],
        "text_identical": text1 == text2,
        "deck1_metadata": meta1,
        "deck2_metadata": meta2,
    }


def main():
    parser = argparse.ArgumentParser(description="Batch process PPTX files")
    subparsers = parser.add_subparsers(dest="command")

    # Extract command
    extract_parser = subparsers.add_parser("extract", help="Extract text content")
    extract_parser.add_argument("input", help="Input PPTX file")
    extract_parser.add_argument("--format", choices=["text", "markdown", "json"],
                               default="text", help="Output format")

    # Info command
    info_parser = subparsers.add_parser("info", help="Show file metadata")
    info_parser.add_argument("input", help="Input PPTX file")

    # Compare command
    compare_parser = subparsers.add_parser("compare", help="Compare two decks")
    compare_parser.add_argument("deck1", help="First PPTX file")
    compare_parser.add_argument("deck2", help="Second PPTX file")

    args = parser.parse_args()

    if args.command == "extract":
        result = extract_text(Path(args.input), args.format)
        print(result)
    elif args.command == "info":
        meta = extract_metadata(Path(args.input))
        print(json.dumps(meta, indent=2))
    elif args.command == "compare":
        diff = compare_decks(Path(args.deck1), Path(args.deck2))
        print(json.dumps(diff, indent=2))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
