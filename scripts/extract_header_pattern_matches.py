from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict, deque
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass
class ParagraphRuns:
    text: str
    runs: list[dict[str, Any]]


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


def parse_bool_mode(value: str) -> str:
    normalized = value.strip().lower()
    if normalized not in {"true", "false", "any"}:
        raise ValueError("Value must be one of: true, false, any")
    return normalized


def bool_mode_matches(actual: bool, mode: str) -> bool:
    if mode == "any":
        return True
    if mode == "true":
        return actual
    return not actual


def color_matches(
    style: dict[str, Any],
    target_rgb: tuple[int, int, int] | None,
    tolerance: int,
) -> bool:
    if target_rgb is None:
        return True

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


def style_matches_run(
    run_style: dict[str, Any],
    italic_mode: str,
    bold_mode: str,
    underline_mode: str,
    target_rgb: tuple[int, int, int] | None,
    tolerance: int,
) -> bool:
    italic_ok = bool_mode_matches(bool(run_style.get("italic", False)), italic_mode)
    bold_ok = bool_mode_matches(bool(run_style.get("bold", False)), bold_mode)
    underline_ok = bool_mode_matches(bool(run_style.get("underline", False)), underline_mode)
    color_ok = color_matches(run_style, target_rgb, tolerance)
    return italic_ok and bold_ok and underline_ok and color_ok


def collect_paragraph_runs(extracted_runs: list[dict[str, Any]]) -> dict[int, ParagraphRuns]:
    grouped: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for run in extracted_runs:
        paragraph_index = run.get("paragraph_index")
        run_index = run.get("run_index")
        text = run.get("text", "")
        if not isinstance(paragraph_index, int):
            continue
        if not isinstance(run_index, int):
            continue
        if not isinstance(text, str):
            continue
        grouped[paragraph_index].append(run)

    result: dict[int, ParagraphRuns] = {}
    for paragraph_index, runs in grouped.items():
        ordered_runs = sorted(runs, key=lambda r: int(r["run_index"]))
        text = "".join(str(r.get("text", "")) for r in ordered_runs)
        result[paragraph_index] = ParagraphRuns(text=text, runs=ordered_runs)
    return result


def normalize_text(text: str) -> str:
    return text.replace("\xa0", " ").strip()


def collect_paragraph_texts_under_heading(
    node: dict[str, Any],
    header_pattern: re.Pattern[str],
) -> list[str]:
    matches: list[str] = []

    if node.get("type") == "heading" and header_pattern.search(str(node.get("text", ""))):
        matches.extend(collect_descendant_paragraph_texts(node))

    for child in node.get("children", []):
        if isinstance(child, dict):
            matches.extend(collect_paragraph_texts_under_heading(child, header_pattern))

    return matches


def collect_descendant_paragraph_texts(node: dict[str, Any]) -> list[str]:
    texts: list[str] = []
    for child in node.get("children", []):
        if not isinstance(child, dict):
            continue
        if child.get("type") == "paragraph":
            texts.append(str(child.get("text", "")))
        texts.extend(collect_descendant_paragraph_texts(child))
    return texts


def map_paragraph_texts_to_indexes(
    paragraph_texts: list[str],
    paragraph_runs: dict[int, ParagraphRuns],
) -> list[int]:
    by_text: dict[str, deque[int]] = defaultdict(deque)
    for index in sorted(paragraph_runs):
        key = normalize_text(paragraph_runs[index].text)
        by_text[key].append(index)

    mapped_indexes: list[int] = []
    for text in paragraph_texts:
        key = normalize_text(text)
        if by_text[key]:
            mapped_indexes.append(by_text[key].popleft())

    return mapped_indexes


def run_overlaps(match_start: int, match_end: int, run: dict[str, Any]) -> bool:
    run_start = run.get("char_start")
    run_end = run.get("char_end")
    if not isinstance(run_start, int) or not isinstance(run_end, int):
        return True
    return run_start < match_end and run_end > match_start


def extract_matches(
    data: dict[str, Any],
    header_regex: str,
    text_pattern: str,
    capture_group: int,
    ignore_case: bool,
    italic_mode: str,
    bold_mode: str,
    underline_mode: str,
    target_rgb: tuple[int, int, int] | None,
    tolerance: int,
    unique: bool,
) -> list[str]:
    extracted_runs = data.get("extracted_text_runs", [])
    structure = data.get("structure", {})

    if not isinstance(extracted_runs, list):
        raise ValueError("Invalid JSON: extracted_text_runs must be a list")
    if not isinstance(structure, dict):
        raise ValueError("Invalid JSON: structure must be an object")

    paragraph_runs = collect_paragraph_runs(extracted_runs)
    if not paragraph_runs:
        raise ValueError(
            "No DOCX paragraph-indexed runs found. Use JSON generated from a DOCX input."
        )

    regex_flags = re.IGNORECASE if ignore_case else 0
    header_pattern = re.compile(header_regex, regex_flags)
    value_pattern = re.compile(text_pattern, regex_flags)

    paragraph_texts = collect_paragraph_texts_under_heading(structure, header_pattern)
    paragraph_indexes = map_paragraph_texts_to_indexes(paragraph_texts, paragraph_runs)

    results: list[str] = []
    seen: set[str] = set()

    for paragraph_index in paragraph_indexes:
        paragraph = paragraph_runs[paragraph_index]
        for match in value_pattern.finditer(paragraph.text):
            if capture_group > match.re.groups:
                raise ValueError(
                    f"Pattern has only {match.re.groups} capture groups, cannot use {capture_group}."
                )

            token = match.group(capture_group) if capture_group > 0 else match.group(0)
            if token is None:
                continue
            token = token.strip()
            if not token:
                continue

            overlaps_with_matching_style = False
            for run in paragraph.runs:
                if not run_overlaps(match.start(), match.end(), run):
                    continue
                style = run.get("style", {})
                if not isinstance(style, dict):
                    continue
                if style_matches_run(
                    style,
                    italic_mode=italic_mode,
                    bold_mode=bold_mode,
                    underline_mode=underline_mode,
                    target_rgb=target_rgb,
                    tolerance=tolerance,
                ):
                    overlaps_with_matching_style = True
                    break

            if not overlaps_with_matching_style:
                continue

            if unique:
                if token in seen:
                    continue
                seen.add(token)

            results.append(token)

    return results


def write_output(matches: list[str], output_path: Path, output_format: str) -> None:
    if output_format == "json":
        output_path.write_text(json.dumps(matches, indent=2, ensure_ascii=False), encoding="utf-8")
        return

    output = "\n".join(matches)
    if output:
        output += "\n"
    output_path.write_text(output, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Find regex matches under a target heading in extractor JSON and filter matches "
            "by overlapping run style/color."
        )
    )
    parser.add_argument("input_json", type=Path, help="Path to extractor JSON output")
    parser.add_argument("output_file", type=Path, help="File to write the extracted matches list")
    parser.add_argument("--header", required=True, help="Header regex to scope matching paragraphs")
    parser.add_argument(
        "--pattern",
        default=r"\[Covers:\s*(.*?)\]",
        help="Regex to extract text (default: \\[Covers:\\s*(.*?)\\])",
    )
    parser.add_argument(
        "--capture-group",
        type=int,
        default=1,
        help="Capture group index to output (0 for full match, default: 1)",
    )
    parser.add_argument(
        "--ignore-case",
        action="store_true",
        help="Case-insensitive matching for header and pattern",
    )
    parser.add_argument(
        "--italic",
        default="any",
        help="Italic filter: true, false, or any (default: any)",
    )
    parser.add_argument(
        "--bold",
        default="any",
        help="Bold filter: true, false, or any (default: any)",
    )
    parser.add_argument(
        "--underline",
        default="any",
        help="Underline filter: true, false, or any (default: any)",
    )
    parser.add_argument(
        "--color",
        default=None,
        help="Optional color filter as #RRGGBB or r,g,b",
    )
    parser.add_argument(
        "--tolerance",
        type=int,
        default=0,
        help="Per-channel RGB tolerance (0-255, default: 0)",
    )
    parser.add_argument(
        "--unique",
        action="store_true",
        help="Return unique values only, preserving first-seen order",
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
    input_path: Path = args.input_json
    output_path: Path = args.output_file

    if not input_path.exists():
        raise FileNotFoundError(f"Input JSON not found: {input_path}")
    if args.capture_group < 0:
        raise ValueError("--capture-group must be >= 0")
    if args.tolerance < 0 or args.tolerance > 255:
        raise ValueError("--tolerance must be between 0 and 255")

    italic_mode = parse_bool_mode(args.italic)
    bold_mode = parse_bool_mode(args.bold)
    underline_mode = parse_bool_mode(args.underline)
    target_rgb = parse_color(args.color) if args.color else None

    data = load_json(input_path)
    matches = extract_matches(
        data,
        header_regex=args.header,
        text_pattern=args.pattern,
        capture_group=args.capture_group,
        ignore_case=bool(args.ignore_case),
        italic_mode=italic_mode,
        bold_mode=bold_mode,
        underline_mode=underline_mode,
        target_rgb=target_rgb,
        tolerance=args.tolerance,
        unique=bool(args.unique),
    )

    write_output(matches, output_path, args.format)
    print(f"Wrote {len(matches)} match(es) to {output_path}")


if __name__ == "__main__":
    main()