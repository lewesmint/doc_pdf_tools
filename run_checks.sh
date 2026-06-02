#!/usr/bin/env bash

set -euo pipefail

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
repo_root="$script_dir"
cd "$repo_root"

target_py=(
  extract_structure.py
)

lint_targets=(
  extract_structure.py
  scripts/generate_sample_documents.py
  scripts/regenerate_from_json.py
  scripts/run_iterative_roundtrip.py
)

type_targets=(
  extract_structure.py
  scripts/generate_sample_documents.py
  scripts/regenerate_from_json.py
  scripts/run_iterative_roundtrip.py
)

for f in "${target_py[@]}"; do
  if [[ ! -f "$f" ]]; then
    echo "Missing target file: $f"
    exit 1
  fi
done

# Python resolution order:
# 1) explicit override via PYTHON_BIN
# 2) active virtualenv interpreter (if VIRTUAL_ENV is set)
# 3) workspace-local .venv
# 4) current shell python/python3
python_bin="${PYTHON_BIN:-}"
if [[ -z "$python_bin" && -n "${VIRTUAL_ENV:-}" && -x "${VIRTUAL_ENV}/bin/python" ]]; then
  python_bin="${VIRTUAL_ENV}/bin/python"
fi
if [[ -z "$python_bin" && -x "$script_dir/.venv/bin/python" ]]; then
  python_bin="$script_dir/.venv/bin/python"
fi
if [[ -z "$python_bin" ]]; then
  python_bin="$(command -v python || true)"
fi
if [[ -z "$python_bin" ]]; then
  python_bin="$(command -v python3)"
fi

echo "Using Python: $python_bin"
"$python_bin" --version
echo

echo "===== INSTALL_LINT_TOOLS ====="
"$python_bin" -m pip install --upgrade pip
"$python_bin" -m pip install -r requirements.txt
"$python_bin" -m pip install pytest ruff mypy bandit pylint pyright
echo "[exit:0]"
echo

run_check() {
  echo "===== $1 ====="
  shift
  if "$@"; then
    exit_code=0
  else
    exit_code=$?
  fi
  echo "[exit:$exit_code]"
  echo
  return 0
}

run_check PY_COMPILE "$python_bin" -m py_compile "${target_py[@]}"
run_check RUFF_FIX "$python_bin" -m ruff check --fix "${lint_targets[@]}"
run_check RUFF "$python_bin" -m ruff check "${lint_targets[@]}"
run_check MYPY "$python_bin" -m mypy
run_check PYTEST "$python_bin" -m pytest -q
run_check BANDIT "$python_bin" -m bandit "${lint_targets[@]}"
run_check PYLINT "$python_bin" -m pylint "${lint_targets[@]}"
run_check PYRIGHT "$python_bin" -m pyright "${type_targets[@]}"