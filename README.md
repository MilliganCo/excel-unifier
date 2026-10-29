# Excel Unifier

Python utility that turns invoice/packing-list tables into six columns:
`№`, `Наименование`, `Артикул`, `Цвет`, `Размер`, `Количество`.
Each apparel size becomes a separate output row. Non-apparel dimensions remain
part of the product description rather than an apparel size variation.

The parser recognizes Russian/English headers, size distributions, subtotals and
product summaries. Its existing article-priority grammar favors codes beginning
with `MJM`; this source release preserves that behavior. All bundled test codes,
quantities and prices are invented. No invoice, customer table or output workbook
is included.

## Setup and run

Use Python 3.10+. On Windows:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python excel_unifier.py
```

On Linux/macOS, use `python3 -m venv .venv` and `.venv/bin/python` for the same
commands. The graphical mode needs Tkinter; install it separately through your
platform's Python distribution if absent. The launchers `start.bat` / `bash start.sh`
use the local `.venv` when present and never install packages automatically.

For explicit files without a graphical dialog:

```sh
python excel_unifier.py input.xlsx --output-dir Result
```

Results are written beside the source in the named subfolder. Source books are
not overwritten; an existing result at the chosen output path is replaced.
`.xlsx` and `.xlsm` inputs are read through openpyxl; output is `.xlsx` without
macros. Legacy `.xls` uses xlrd, or optionally LibreOffice when xlrd is unavailable.
No Microsoft Excel/COM automation is required.

## Checks

```sh
python -B excel_unifier.py --self-test
python -B -m unittest discover -s tests -v
```

Tests create synthetic tables and temporary workbooks. The optional private-file
smoke test skips unless `EXCEL_UNIFIER_PRIVATE_FIXTURE` is explicitly set by the
caller. That file is read only, output stays in a temporary directory, and no
customer-specific totals or filenames are embedded in the test.

See [VALIDATION.md](VALIDATION.md) for actual results. Header/table detection and
article reconciliation are heuristics: inspect output before operational use.
Graphical interaction and LibreOffice/legacy workbook behavior require separate
acceptance. No project-wide reuse license was supplied with the local source.
