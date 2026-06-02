# Document Structure Extractor (PDF + DOCX -> JSON)

Extracts:
- headings and child headings (hierarchy)
- paragraph content under each heading
- style metadata for extracted text runs (font/color/bold/italic where available)
- specially formatted text runs (blue + italic + bold subset)

`extracted_text_runs` preserves exact run/span text (including spaces) and includes
character offsets (`char_start`/`char_end`) so mid-paragraph style changes can be
reconstructed without splitting the paragraph tree.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python extract_structure.py /path/to/file.docx
python extract_structure.py /path/to/file.pdf
```

Optional output path:

```bash
python extract_structure.py /path/to/file.pdf -o output.json
```

## Roundtrip Validation (JSON -> file -> JSON)

This workflow tests whether extracted JSON can regenerate a document and then
be re-extracted with equivalent structure.

Run the full process:

```bash
./run_roundtrip.sh
```

What it does:
- regenerates sample input documents
- extracts them to JSON
- regenerates DOCX/PDF from those JSON files
- re-extracts regenerated files for multiple iterations
- writes a report to `roundtrip_outputs/roundtrip_report.json`

You can also run the iterative engine directly:

```bash
python scripts/run_iterative_roundtrip.py --iterations 3 --output-dir roundtrip_outputs
```

## Extract Paragraph List By Style

If you already have extractor JSON from a DOCX file, you can export paragraph text
that contains at least one run with a target color and italic rule.

Text output (one paragraph per line):

```bash
python scripts/extract_paragraphs_by_style.py input.docx.json matches.txt --color 0,0,255 --italic true
```

JSON list output:

```bash
python scripts/extract_paragraphs_by_style.py input.docx.json matches.json --color #FF0000 --italic false --format json
```

Notes:
- `--color` accepts `#RRGGBB` or `r,g,b`
- `--italic` can be `true`, `false`, or `any`
- `--tolerance` can be used for near-color matching (0-255 per channel)

## Extract Pattern Values Under A Header

You can scope extraction to a specific heading and pull repeated values from a
pattern such as `[Covers: ???]`, with optional style/color overlap filters.

Example (extract `???` values under header `Scope` where matching text is blue + italic):

```bash
python scripts/extract_header_pattern_matches.py input.docx.json covers.txt --header "^Scope$" --pattern "\[Covers:\s*(.*?)\]" --capture-group 1 --italic true --color 0,0,255
```

Useful options:
- `--unique` to de-duplicate values while keeping first-seen order
- `--bold` / `--underline` style filters (`true` / `false` / `any`)
- `--ignore-case` for case-insensitive header and pattern matching
- `--format json` for JSON list output

## Output Shape

```json
{
  "source_file": "...",
  "type": "pdf|docx",
  "structure": {
    "type": "document",
    "text": "file-name",
    "children": [
      {
        "type": "heading",
        "text": "Heading 1",
        "level": 1,
        "children": [
          {
            "type": "paragraph",
            "text": "Paragraph under heading"
          },
          {
            "type": "heading",
            "text": "Child heading",
            "level": 2,
            "children": []
          }
        ]
      }
    ]
  },
  "special_formatted_text": [
    {
      "text": "Important phrase",
      "style": {
        "bold": true,
        "italic": true,
        "color_rgb": { "r": 0, "g": 0, "b": 255 }
      }
    }
  ],
  "extracted_text_runs": [
    {
      "text": "Important phrase",
      "paragraph_index": 3,
      "run_index": 1,
      "char_start": 18,
      "char_end": 34,
      "style": {
        "bold": true,
        "italic": true,
        "color_rgb": { "r": 0, "g": 0, "b": 255 },
        "font": "Calibri",
        "font_size": 12.0
      }
    }
  ]
}
```

## Notes

- `.docx` extraction uses Word style names (`Heading 1`, `Heading 2`, etc.).
- PDF heading detection is heuristic (font size/weight based), because PDFs do not have semantic heading tags.
- Legacy `.doc` is not parsed directly by this script; convert `.doc` to `.docx` first.
- On Windows, you can use [scripts/convert_doc_to_docx.ps1](scripts/convert_doc_to_docx.ps1) to convert `.doc` files with installed Microsoft Word before running the extractor.
