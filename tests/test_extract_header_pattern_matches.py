from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


def _script_path() -> Path:
    return Path(__file__).resolve().parents[1] / "scripts" / "extract_header_pattern_matches.py"


def _build_docx_style_json() -> dict[str, object]:
    return {
        "type": "docx",
        "structure": {
            "type": "document",
            "text": "sample.docx",
            "children": [
                {
                    "type": "heading",
                    "text": "Scope",
                    "level": 1,
                    "children": [
                        {
                            "type": "paragraph",
                            "text": "[Covers: Alpha] and [Covers: Beta]",
                        },
                        {
                            "type": "paragraph",
                            "text": "[Covers: Alpha] repeated",
                        },
                    ],
                },
                {
                    "type": "heading",
                    "text": "Other",
                    "level": 1,
                    "children": [
                        {
                            "type": "paragraph",
                            "text": "[Covers: Ignored]",
                        }
                    ],
                },
            ],
        },
        "extracted_text_runs": [
            {
                "text": "[Covers: Alpha]",
                "paragraph_index": 0,
                "run_index": 0,
                "char_start": 0,
                "char_end": 15,
                "style": {"italic": True, "bold": False, "underline": False, "color_rgb": {"r": 0, "g": 0, "b": 255}},
            },
            {
                "text": " and ",
                "paragraph_index": 0,
                "run_index": 1,
                "char_start": 15,
                "char_end": 20,
                "style": {"italic": False, "bold": False, "underline": False, "color_rgb": {"r": 0, "g": 0, "b": 0}},
            },
            {
                "text": "[Covers: Beta]",
                "paragraph_index": 0,
                "run_index": 2,
                "char_start": 20,
                "char_end": 34,
                "style": {"italic": True, "bold": False, "underline": False, "color_rgb": {"r": 0, "g": 0, "b": 255}},
            },
            {
                "text": "[Covers: Alpha]",
                "paragraph_index": 1,
                "run_index": 0,
                "char_start": 0,
                "char_end": 15,
                "style": {"italic": True, "bold": False, "underline": False, "color_rgb": {"r": 0, "g": 0, "b": 255}},
            },
            {
                "text": " repeated",
                "paragraph_index": 1,
                "run_index": 1,
                "char_start": 15,
                "char_end": 24,
                "style": {"italic": False, "bold": False, "underline": False, "color_rgb": {"r": 0, "g": 0, "b": 0}},
            },
            {
                "text": "[Covers: Ignored]",
                "paragraph_index": 2,
                "run_index": 0,
                "char_start": 0,
                "char_end": 17,
                "style": {"italic": True, "bold": False, "underline": False, "color_rgb": {"r": 0, "g": 0, "b": 255}},
            },
        ],
    }


def test_extracts_covers_values_under_target_header(tmp_path: Path) -> None:
    input_json = tmp_path / "input.json"
    output_file = tmp_path / "covers.txt"
    input_json.write_text(json.dumps(_build_docx_style_json()), encoding="utf-8")

    subprocess.run(
        [
            sys.executable,
            str(_script_path()),
            str(input_json),
            str(output_file),
            "--header",
            "^Scope$",
            "--pattern",
            r"\[Covers:\s*(.*?)\]",
            "--capture-group",
            "1",
            "--italic",
            "true",
            "--color",
            "0,0,255",
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    lines = output_file.read_text(encoding="utf-8").splitlines()
    assert lines == ["Alpha", "Beta", "Alpha"]


def test_extracts_unique_values_as_json(tmp_path: Path) -> None:
    input_json = tmp_path / "input.json"
    output_file = tmp_path / "covers.json"
    input_json.write_text(json.dumps(_build_docx_style_json()), encoding="utf-8")

    subprocess.run(
        [
            sys.executable,
            str(_script_path()),
            str(input_json),
            str(output_file),
            "--header",
            "scope",
            "--ignore-case",
            "--pattern",
            r"\[Covers:\s*(.*?)\]",
            "--capture-group",
            "1",
            "--italic",
            "true",
            "--color",
            "#0000FF",
            "--unique",
            "--format",
            "json",
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    values = json.loads(output_file.read_text(encoding="utf-8"))
    assert values == ["Alpha", "Beta"]


def test_style_overlap_filter_excludes_non_matching_runs(tmp_path: Path) -> None:
    payload = _build_docx_style_json()
    runs = payload["extracted_text_runs"]
    assert isinstance(runs, list)

    # Make Beta non-italic so it no longer matches the style filter.
    runs[2]["style"]["italic"] = False

    input_json = tmp_path / "input.json"
    output_file = tmp_path / "covers.txt"
    input_json.write_text(json.dumps(payload), encoding="utf-8")

    subprocess.run(
        [
            sys.executable,
            str(_script_path()),
            str(input_json),
            str(output_file),
            "--header",
            "^Scope$",
            "--pattern",
            r"\[Covers:\s*(.*?)\]",
            "--italic",
            "true",
            "--color",
            "0,0,255",
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    lines = output_file.read_text(encoding="utf-8").splitlines()
    assert lines == ["Alpha", "Alpha"]