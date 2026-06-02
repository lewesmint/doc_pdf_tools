#!/usr/bin/env bash

set -euo pipefail

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
cd "$script_dir"

python scripts/generate_sample_documents.py
python extract_structure.py sample_inputs/sample_structured.docx
python extract_structure.py sample_inputs/sample_structured.pdf
python extract_structure.py sample_inputs/sample_challenging.docx
python extract_structure.py sample_inputs/sample_challenging.pdf
python scripts/run_iterative_roundtrip.py --iterations 2 --output-dir roundtrip_outputs

echo
echo "Roundtrip report: roundtrip_outputs/roundtrip_report.json"
