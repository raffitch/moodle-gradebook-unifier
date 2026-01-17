#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INPUT_DIR="${1:-.}"
OUTPUT_PATH="${2:-${SCRIPT_DIR}/consolidated.xlsx}"

cd "${SCRIPT_DIR}"

echo "Using input directory: ${INPUT_DIR}"
echo "Output file will be: ${OUTPUT_PATH}"

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip >/dev/null
pip install -r requirements.txt

python consolidate_grades.py --input-dir "${INPUT_DIR}" --output "${OUTPUT_PATH}"

deactivate
echo "Consolidation complete. Output: ${OUTPUT_PATH}"
