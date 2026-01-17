# Moodle Gradebook Unifier

Python utility to combine multiple Moodle rubric exports into a single consolidated workbook. It reads the numbered assignment XLSX files, applies rubric labels from the matching CSVs, pulls overall grades from the `00-*.xlsx` course total file, and writes a styled, side-by-side gradebook.

## Prerequisites
- Python 3.9+ and `bash`
- Moodle exports in one folder:
  - `00-<course> Grades.xlsx` (course totals with percentage and letter columns)
  - `01-...`, `02-...`, etc. rubric XLSX files
  - Matching rubric CSVs for each assignment (same name as the XLSX without the numeric prefix, e.g. `PrW301_20251_ALL - Assignment - Phase 1 - Midterm - 45% - Rubric Percentage.csv`)

## Quick start (recommended)
```bash
./run_consolidation.sh /path/to/exports /path/to/output.xlsx
```
- First argument: directory containing the Moodle XLSX/CSV exports (defaults to current directory if omitted).
- Second argument: output file path (defaults to `consolidated.xlsx` in the repo root).
- The script creates a `.venv`, installs dependencies, runs the consolidation, and deactivates the venv when done.
- If LibreOffice (`soffice`) is on your PATH, a PDF (`<output>.pdf`) is produced in landscape, fitted to one page.

## Manual usage
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python consolidate_grades.py --input-dir /path/to/exports --output consolidated.xlsx
deactivate
```
- A PDF will also be written next to the XLSX when LibreOffice is available.

## Output layout
- Top banner: course name + “Assignment Grade Breakdown Per Criteria”.
- Name block once on the left, assignment blocks side-by-side, course total on the right.
- Criterion names pulled from rubric CSVs (fallback to `Criterion N`), totals, over-100 scores, and letters included.
- Gray palette, zebra striping, and thick separators between sections; headers auto-size horizontally.

## Rubric weight export (Tampermonkey)
Use the companion userscript to capture rubric criterion weights as the CSVs the Python tool expects.

1) Install the Tampermonkey browser extension.
2) Add the script from `tampermonkey-rubric-export.user.js` (raw URL: https://raw.githubusercontent.com/raffitch/moodle-gradebook-unifier/main/tampermonkey-rubric-export.user.js).
3) On a Moodle assignment’s Advanced Grading page (Rubric), the script adds an “Export rubric % CSV” button above the rubric.
4) Click the button to download a CSV of criterion labels and weights. Rename the file to match your assignment export (same name as the assignment XLSX without the numeric prefix, ending with `- Rubric Percentage.csv`) and place it alongside the XLSX exports before running the consolidation.

## Notes
- Rows containing “Raffi” are removed.
- Grade values are left untouched; numeric cells are shown with two decimal places.
- If rubric CSVs are missing, the script falls back to generic criterion labels.
