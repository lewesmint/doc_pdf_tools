from __future__ import annotations

import argparse
import json
from collections import defaultdict
from pathlib import Path
from typing import Any


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def parse_color(color_value: str) -> tuple[int, int, int]:
    value = color_value.strip()
    if value.startswith("#"):
        value = value[1:]
    if len(value) == 6 and all(ch in "0123456789abcdefABCDEF" for ch in value):
        return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)

    parts = [p.strip() for p in value.split(",")]
    if len(parts) != 3:
        raise ValueError("Color must be #RRGGBB or r,g,b")
    rgb = tuple(int(p) for p in parts)
    if any(channel < 0 or channel > 255 for channel in rgb):
        raise ValueError("Each RGB channel must be between 0 and 255")
    return rgb


def italic_matches(style: dict[str, Any], mode: str) -> bool:
    is_italic = bool(style.get("italic", False))
    if mode == "any":
        return True
    if mode == "true":
        return is_italic
    return not is_italic


def color_matches(style: dict[str, Any], target_rgb: tuple[int, int, int], tolerance: int) -> bool:
    rgb = style.get("color_rgb", {})
    if not isinstance(rgb, dict):
        return False
    try:
        r = int(rgb.get("r"))
        g = int(rgb.get("g"))
        b = int(rgb.get("b"))
    except (TypeError, ValueError):
        return False

    return (
        abs(r - target_rgb[0]) <= tolerance
        and abs(g - target_rgb[1]) <= tolerance
        and abs(b - target_rgb[2]) <= tolerance
    )


def collect_docx_paragraph_texts(extracted_runs: list[dict[str, Any]]) -> dict[int, str]:
    grouped: dict[int, list[tuple[int, str]]] = defaultdict(list)
    for run in extracted_runs:
        paragraph_index = run.get("paragraph_index")
        run_index = run.get("run_index")
        text = run.get("text", "")
        if not isinstance(paragraph_index, int) or not isinstance(run_index, int):
            continue
        if not isinstance(text, str):
            continue
        grouped[paragraph_index].append((run_index, text))

    paragraphs: dict[int, str] = {}
    for paragraph_index, pieces in grouped.items():
        ordered = [text for _, text in sorted(pieces, key=lambda item: item[0])]
        paragraphs[paragraph_index] = "".join(ordered)
    return paragraphs


def extract_matching_paragraphs(
    data: dict[str, Any],
    color_rgb: tuple[int, int, int],
    italic_mode: str,
    tolerance: int,
) -> list[str]:
    extracted_runs = data.get("extracted_text_runs", [])
    if not isinstance(extracted_runs, list):
        raise ValueError("Invalid JSON: extracted_text_runs must be a list")

    paragraph_map = collect_docx_paragraph_texts(extracted_runs)
    if not paragraph_map:
        raise ValueError(
            "No DOCX paragraph-indexed runs found. Use JSON generated from a DOCX input."
        )

    matching_indexes: set[int] = set()
    for run in extracted_runs:
        paragraph_index = run.get("paragraph_index")
        if not isinstance(paragraph_index, int):
            continue
        style = run.get("style", {})
        if not isinstance(style, dict):
            continue

        if italic_matches(style, italic_mode) and color_matches(style, color_rgb, tolerance):
            matching_indexes.add(paragraph_index)

    return [paragraph_map[idx].strip() for idx in sorted(matching_indexes) if paragraph_map[idx].strip()]


def write_output(paragraphs: list[str], output_path: Path, output_format: str) -> None:
    if output_format == "json":
        output_path.write_text(json.dumps(paragraphs, indent=2, ensure_ascii=False), encoding="utf-8")
        return

    output_text = "\n".join(paragraphs)
    if output_text:
        output_text += "\n"
    output_path.write_text(output_text, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract paragraph texts from extractor JSON where at least one run matches "
            "a target color and italic setting."
        )
    )
    parser.add_argument("input_json", type=Path, help="Path to extractor JSON output")
    parser.add_argument("output_file", type=Path, help="File to write the paragraph list")
    parser.add_argument(
        "--color",
        required=True,
        help="Target color as #RRGGBB or r,g,b (for example: #0000FF or 0,0,255)",
    )
    parser.add_argument(
        "--italic",
        choices=["true", "false", "any"],
        default="true",
        help="Italic filter for matching runs (default: true)",
    )
    parser.add_argument(
        "--tolerance",
        type=int,
        default=0,
        help="Per-channel RGB tolerance, 0-255 (default: 0)",
    )
    parser.add_argument(
        "--format",
        choices=["text", "json"],
        default="text",
        help="Output format (default: text)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.tolerance < 0 or args.tolerance > 255:
        raise ValueError("--tolerance must be between 0 and 255")

    input_path: Path = args.input_json
    output_path: Path = args.output_file
    if not input_path.exists():
        raise FileNotFoundError(f"Input JSON not found: {input_path}")

    data = load_json(input_path)
    color_rgb = parse_color(args.color)
    paragraphs = extract_matching_paragraphs(
        data,
        color_rgb=color_rgb,
        italic_mode=args.italic,
        tolerance=args.tolerance,
    )

    write_output(paragraphs, output_path, args.format)
    print(f"Wrote {len(paragraphs)} paragraph(s) to {output_path}")


if __name__ == "__main__":
    main()