from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


def _script_path() -> Path:
    return Path(__file__).resolve().parents[1] / "scripts" / "extract_paragraphs_by_style.py"


def test_extracts_matching_docx_paragraphs_to_text_file(tmp_path: Path) -> None:
    input_json = tmp_path / "input.json"
    output_file = tmp_path / "matches.txt"

    data = {
        "type": "docx",
        "extracted_text_runs": [
            {
                "text": "First paragraph. ",
                "paragraph_index": 0,
                "run_index": 0,
                "style": {"italic": False, "color_rgb": {"r": 0, "g": 0, "b": 0}},
            },
            {
                "text": "Blue italic fragment",
                "paragraph_index": 0,
                "run_index": 1,
                "style": {"italic": True, "color_rgb": {"r": 0, "g": 0, "b": 255}},
            },
            {
                "text": "Second paragraph plain",
                "paragraph_index": 1,
                "run_index": 0,
                "style": {"italic": False, "color_rgb": {"r": 255, "g": 0, "b": 0}},
            },
            {
                "text": "Third paragraph",
                "paragraph_index": 2,
                "run_index": 0,
                "style": {"italic": True, "color_rgb": {"r": 0, "g": 0, "b": 255}},
            },
        ],
    }
    input_json.write_text(json.dumps(data), encoding="utf-8")

    subprocess.run(
        [
            sys.executable,
            str(_script_path()),
            str(input_json),
            str(output_file),
            "--color",
            "0,0,255",
            "--italic",
            "true",
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    lines = output_file.read_text(encoding="utf-8").splitlines()
    assert lines == ["First paragraph. Blue italic fragment", "Third paragraph"]


def test_extracts_with_hex_color_and_json_output(tmp_path: Path) -> None:
    input_json = tmp_path / "input.json"
    output_file = tmp_path / "matches.json"

    data = {
        "type": "docx",
        "extracted_text_runs": [
            {
                "text": "Para A",
                "paragraph_index": 0,
                "run_index": 0,
                "style": {"italic": False, "color_rgb": {"r": 255, "g": 0, "b": 0}},
            },
            {
                "text": "Para B",
                "paragraph_index": 1,
                "run_index": 0,
                "style": {"italic": False, "color_rgb": {"r": 255, "g": 0, "b": 0}},
            },
        ],
    }
    input_json.write_text(json.dumps(data), encoding="utf-8")

    subprocess.run(
        [
            sys.executable,
            str(_script_path()),
            str(input_json),
            str(output_file),
            "--color",
            "#FF0000",
            "--italic",
            "false",
            "--format",
            "json",
        ],
        check=True,
        capture_output=True,
        text=True,
    )

    payload = json.loads(output_file.read_text(encoding="utf-8"))
    assert payload == ["Para A", "Para B"]


def test_rejects_non_docx_style_json_without_paragraph_indexes(tmp_path: Path) -> None:
    input_json = tmp_path / "input.json"
    output_file = tmp_path / "matches.txt"

    data = {
        "type": "pdf",
        "extracted_text_runs": [
            {
                "text": "Line one",
                "paragraph_index": None,
                "run_index": None,
                "style": {"italic": True, "color_rgb": {"r": 0, "g": 0, "b": 255}},
            }
        ],
    }
    input_json.write_text(json.dumps(data), encoding="utf-8")

    proc = subprocess.run(
        [
            sys.executable,
            str(_script_path()),
            str(input_json),
            str(output_file),
            "--color",
            "0,0,255",
        ],
        check=False,
        capture_output=True,
        text=True,
    )

    assert proc.returncode != 0
    assert "DOCX paragraph-indexed" in proc.stderr