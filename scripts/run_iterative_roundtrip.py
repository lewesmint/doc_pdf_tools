from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from regenerate_from_json import regenerate_docx, regenerate_pdf

from extract_structure import extract_file


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def canonical_special_text(text: str) -> str:
    norm = normalize_text(text)
    if ":" in norm:
        tail = norm.split(":")[-1].strip()
        if tail:
            return tail
    return norm


def normalize_structure(node: dict[str, Any], collapse_tables: bool = False) -> dict[str, Any]:
    node_type = node.get("type", "")
    text = normalize_text(str(node.get("text", "")))

    if collapse_tables and node_type == "table":
        node_type = "paragraph"
        text = "table"

    norm: dict[str, Any] = {
        "type": node_type,
        "text": text,
    }

    if node_type == "heading":
        norm["level"] = int(node.get("level", 1) or 1)
    if node_type == "table":
        rows_raw = node.get("rows", [])
        rows_norm: list[list[str]] = []
        if isinstance(rows_raw, list):
            for row in rows_raw:
                if not isinstance(row, list):
                    continue
                rows_norm.append([normalize_text(str(cell)) for cell in row])
        norm["rows"] = rows_norm

    children = [normalize_structure(c, collapse_tables=collapse_tables) for c in node.get("children", [])]
    if children:
        norm["children"] = children

    return norm


def normalize_special(data: dict[str, Any], canonical: bool = False) -> list[str]:
    if canonical:
        specials = [
            canonical_special_text(str(item.get("text", ""))) for item in data.get("special_formatted_text", [])
        ]
    else:
        specials = [normalize_text(str(item.get("text", ""))) for item in data.get("special_formatted_text", [])]
    return sorted([s for s in specials if s])


def _normalize_style(item: dict[str, Any], include_font_size: bool = True) -> dict[str, Any]:
    style = item.get("style", {}) if isinstance(item, dict) else {}
    rgb = style.get("color_rgb", {}) if isinstance(style, dict) else {}
    normalized: dict[str, Any] = {
        "bold": bool(style.get("bold", False)) if isinstance(style, dict) else False,
        "italic": bool(style.get("italic", False)) if isinstance(style, dict) else False,
        "underline": bool(style.get("underline", False)) if isinstance(style, dict) else False,
        "subscript": bool(style.get("subscript", False)) if isinstance(style, dict) else False,
        "superscript": bool(style.get("superscript", False)) if isinstance(style, dict) else False,
        "r": int(rgb.get("r", -1)) if isinstance(rgb, dict) else -1,
        "g": int(rgb.get("g", -1)) if isinstance(rgb, dict) else -1,
        "b": int(rgb.get("b", -1)) if isinstance(rgb, dict) else -1,
    }
    if include_font_size:
        normalized["font_size"] = (
            round(float(style["font_size"]), 1)
            if isinstance(style, dict) and style.get("font_size") is not None
            else None
        )
    return normalized


def normalize_special_styles(
    data: dict[str, Any],
    canonical_text: bool = False,
    include_font_size: bool = True,
) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for item in data.get("special_formatted_text", []):
        raw_text = str(item.get("text", ""))
        text = canonical_special_text(raw_text) if canonical_text else normalize_text(raw_text)
        if not text:
            continue
        normalized.append({"text": text, "style": _normalize_style(item, include_font_size=include_font_size)})
    return sorted(normalized, key=lambda x: (x["text"], json.dumps(x["style"], sort_keys=True)))


def flatten_structure(node: dict[str, Any], path: str = "root") -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    node_type = str(node.get("type", ""))
    text = normalize_text(str(node.get("text", "")))
    level = int(node.get("level", 0) or 0)
    rows = node.get("rows") if node_type == "table" else None
    items.append({"path": path, "type": node_type, "level": level, "text": text, "rows": rows})

    for idx, child in enumerate(node.get("children", [])):
        child_path = f"{path}.children[{idx}]"
        items.extend(flatten_structure(child, child_path))
    return items


def compare_data(expected: dict[str, Any], actual: dict[str, Any]) -> dict[str, Any]:
    expected_type = str(expected.get("type", "")).lower()
    actual_type = str(actual.get("type", "")).lower()
    cross_format = expected_type != actual_type

    expected_struct = normalize_structure(expected.get("structure", {}), collapse_tables=cross_format)
    actual_struct = normalize_structure(actual.get("structure", {}), collapse_tables=cross_format)

    # Root names are format-specific filenames, not meaningful for parity checks.
    expected_struct["text"] = "DOCUMENT"
    actual_struct["text"] = "DOCUMENT"

    expected_special = normalize_special(expected, canonical=cross_format)
    actual_special = normalize_special(actual, canonical=cross_format)
    expected_special_styles = normalize_special_styles(
        expected,
        canonical_text=cross_format,
        include_font_size=not cross_format,
    )
    actual_special_styles = normalize_special_styles(
        actual,
        canonical_text=cross_format,
        include_font_size=not cross_format,
    )

    expected_flat = flatten_structure(expected_struct)
    actual_flat = flatten_structure(actual_struct)

    first_structure_mismatch: dict[str, Any] | None = None
    if expected_flat != actual_flat:
        limit = min(len(expected_flat), len(actual_flat))
        mismatch_index = None
        for i in range(limit):
            if expected_flat[i] != actual_flat[i]:
                mismatch_index = i
                break
        if mismatch_index is None:
            mismatch_index = limit

        first_structure_mismatch = {
            "index": mismatch_index,
            "expected": expected_flat[mismatch_index] if mismatch_index < len(expected_flat) else None,
            "actual": actual_flat[mismatch_index] if mismatch_index < len(actual_flat) else None,
        }

    first_special_mismatch: dict[str, Any] | None = None
    if expected_special != actual_special:
        limit = min(len(expected_special), len(actual_special))
        mismatch_index = None
        for i in range(limit):
            if expected_special[i] != actual_special[i]:
                mismatch_index = i
                break
        if mismatch_index is None:
            mismatch_index = limit

        first_special_mismatch = {
            "index": mismatch_index,
            "expected": expected_special[mismatch_index] if mismatch_index < len(expected_special) else None,
            "actual": actual_special[mismatch_index] if mismatch_index < len(actual_special) else None,
        }

    first_style_mismatch: dict[str, Any] | None = None
    if expected_special_styles != actual_special_styles:
        limit = min(len(expected_special_styles), len(actual_special_styles))
        mismatch_index = None
        for i in range(limit):
            if expected_special_styles[i] != actual_special_styles[i]:
                mismatch_index = i
                break
        if mismatch_index is None:
            mismatch_index = limit

        first_style_mismatch = {
            "index": mismatch_index,
            "expected": (
                expected_special_styles[mismatch_index] if mismatch_index < len(expected_special_styles) else None
            ),
            "actual": actual_special_styles[mismatch_index] if mismatch_index < len(actual_special_styles) else None,
        }

    return {
        "structure_match": expected_struct == actual_struct,
        "special_match": expected_special == actual_special,
        "style_match": expected_special_styles == actual_special_styles,
        "expected_node_count": len(expected_flat),
        "actual_node_count": len(actual_flat),
        "expected_special_count": len(expected_special),
        "actual_special_count": len(actual_special),
        "first_structure_mismatch": first_structure_mismatch,
        "first_special_mismatch": first_special_mismatch,
        "first_style_mismatch": first_style_mismatch,
        "expected_special": expected_special,
        "actual_special": actual_special,
        "expected_special_styles": expected_special_styles,
        "actual_special_styles": actual_special_styles,
    }


def run_for_json(input_json: Path, output_dir: Path, iterations: int) -> dict[str, Any]:
    base = json.loads(input_json.read_text(encoding="utf-8"))
    src_type = str(base.get("type", "")).lower()
    if src_type not in {"docx", "pdf"}:
        raise ValueError(f"Unsupported source type in {input_json}: {src_type}")

    input_stem = input_json.stem
    current_json_data = base
    iteration_results: list[dict[str, Any]] = []

    for i in range(1, iterations + 1):
        ext = "docx" if src_type == "docx" else "pdf"
        regenerated_file = output_dir / f"{input_stem}.iter{i}.{ext}"
        regenerated_json_file = output_dir / f"{input_stem}.iter{i}.{ext}.json"

        if src_type == "docx":
            regenerate_docx(current_json_data, regenerated_file)
        else:
            regenerate_pdf(current_json_data, regenerated_file)

        extracted = extract_file(regenerated_file)
        regenerated_json_file.write_text(json.dumps(extracted, indent=2, ensure_ascii=False), encoding="utf-8")

        cmp = compare_data(base, extracted)
        iteration_results.append(
            {
                "iteration": i,
                "regenerated_file": str(regenerated_file),
                "regenerated_json": str(regenerated_json_file),
                **cmp,
            }
        )

        current_json_data = extracted

    return {
        "input_json": str(input_json),
        "source_type": src_type,
        "iterations": iteration_results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Run iterative JSON -> file -> JSON roundtrip tests.")
    parser.add_argument(
        "--inputs",
        nargs="+",
        default=[
            "sample_inputs/sample_structured.docx.json",
            "sample_inputs/sample_structured.pdf.json",
            "sample_inputs/sample_challenging.docx.json",
            "sample_inputs/sample_challenging.pdf.json",
        ],
        help="Input JSON files to test",
    )
    parser.add_argument("--iterations", type=int, default=2, help="Roundtrip iterations per input")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("roundtrip_outputs"),
        help="Directory for regenerated artifacts and reports",
    )
    args = parser.parse_args()

    output_dir: Path = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    reports: list[dict[str, Any]] = []
    for raw_path in args.inputs:
        input_json = Path(raw_path)
        if not input_json.exists():
            print(f"Skipping missing input: {input_json}")
            continue
        reports.append(run_for_json(input_json, output_dir, args.iterations))

    final_report = {"reports": reports}
    report_path = output_dir / "roundtrip_report.json"
    report_path.write_text(json.dumps(final_report, indent=2, ensure_ascii=False), encoding="utf-8")

    failed = False
    for rep in reports:
        for iteration in rep["iterations"]:
            if (
                not iteration["structure_match"]
                or not iteration["special_match"]
                or not iteration["style_match"]
            ):
                failed = True

    print(f"Wrote report: {report_path}")
    for rep in reports:
        print(f"\nInput: {rep['input_json']} ({rep['source_type']})")
        for iteration in rep["iterations"]:
            print(
                "  Iter {i}: structure_match={sm} special_match={pm} style_match={stm}".format(
                    i=iteration["iteration"],
                    sm=iteration["structure_match"],
                    pm=iteration["special_match"],
                    stm=iteration["style_match"],
                )
            )

    if failed:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
