# Document Structure Extractor (PDF + DOCX -> JSON)

Extracts:
- headings and child headings (hierarchy)
- paragraph content under each heading
- specially formatted text runs (blue + italic + bold)

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
  ]
}
```

## Notes

- `.docx` extraction uses Word style names (`Heading 1`, `Heading 2`, etc.).
- PDF heading detection is heuristic (font size/weight based), because PDFs do not have semantic heading tags.
- Legacy `.doc` is not parsed directly by this script; convert `.doc` to `.docx` first.
- On Windows, you can use [scripts/convert_doc_to_docx.ps1](scripts/convert_doc_to_docx.ps1) to convert `.doc` files with installed Microsoft Word before running the extractor.
