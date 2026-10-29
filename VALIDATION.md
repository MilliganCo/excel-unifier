# Validation

Checked on Windows with CPython 3.12.10, openpyxl 3.1.5 and xlrd 2.0.2.
The installed packages already satisfied requirements; no package installation
or network access was needed for validation.

| Check | Result |
| --- | --- |
| `python -B excel_unifier.py --self-test` | Passed |
| `python -B -m unittest discover -s tests -v` | 11 passed, 1 optional test skipped |
| `cmd /d /c start.bat --self-test` | Passed using the existing system Python |
| `bash -n start.sh` | Syntax passed with Git Bash; application launch on Unix was not run |
| Python AST parsing | Both Python source files passed |

Synthetic coverage includes article-prefix priority, literal leading-zero codes,
size normalization/order, invalid quantities, duplicate merging, renamed-sheet
selection, subtotal exclusion, one-character article reconciliation and a
negative reconciliation case, towel dimensions, input overwrite protection,
macro-extension output naming and a real temporary XLSX read/write roundtrip.
The roundtrip verifies source SHA-256 preservation, six output headers, invented
quantities and text article cells. Temporary files stay in a disposable test
directory. No customer or supplier workbook was copied or used.

The optional private-fixture structural test was deliberately skipped because no
private workbook was explicitly supplied. Its assertions contain no private
article identifiers, filenames or expected totals.

Not run: GUI interaction, Linux/macOS execution, `.xls` or `.xlsm` content
roundtrips, LibreOffice fallback and real operational workbook acceptance.
No Excel COM, graphical application, browser or external service was started.
Do not treat these checks as proof that every vendor's spreadsheet is recognized.

The parser's production functions are unchanged from the selected local source.
Only its illustrative self-test codes and a source comment were generalized.
Public launchers now use an existing Python environment and leave dependency
installation to the caller. No project-wide license was present in the source;
this snapshot does not invent one.
