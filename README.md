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

## Manual usage
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python consolidate_grades.py --input-dir /path/to/exports --output consolidated.xlsx
deactivate
```

## Output layout
- Top banner: course name + “Assignment Grade Breakdown Per Criteria”.
- Name block once on the left, assignment blocks side-by-side, course total on the right.
- Criterion names pulled from rubric CSVs (fallback to `Criterion N`), totals, over-100 scores, and letters included.
- Rotated headers, gray palette, zebra striping, and thick separators between sections.

## Notes
- Rows containing “Raffi” are removed.
- Grade values are left untouched; numeric cells are shown with two decimal places.
- If rubric CSVs are missing, the script falls back to generic criterion labels.
