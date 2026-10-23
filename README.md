# compare-migration: SysML v1.7 to SysML v2 migration success auditor

[![DOI](https://zenodo.org/badge/1430626406.svg)](https://doi.org/10.5281/zenodo.23587761)

The CATIA Magic / MSOSA 2026x SysML v2 migration plugin reports only what
failed to migrate. This tool produces the inverse: an Excel workbook that
shows which SysML 1.7 elements provably arrived in the SysML v2 text, which
ones fell through the cracks, and a per-metatype success rate. Matching is
deterministic (element IDs first, then qualified names), so results are
reproducible.

Supporting software for the doctoral praxis "Metamodel Scope Using Model-Based Decision on Intended-Use & Requirements Engineering Categorization Tool (MB-DIRECT)"
(Evan Matthews, The George Washington University, School of Engineering and
Applied Science, 2026).

## What is here

| Path | Contents |
|---|---|
| `src/compare_migration.py` | The auditor. Python 3.8+, needs openpyxl. One self-contained file. |
| `templates/AllElements_Export.vm` | Report Wizard template that exports every element in scope as a 7-column CSV. |
| `examples/AllElements_1_7.csv` | Inventory export of a small model. |
| `examples/MigrationReportOfHomegrown_Claude.xlsx` | The not-migrated report for the same model, in the CATIA report layout. |
| `examples/Homegrown_generated_pilot.txt` | The SysML v2 text for the same model. |
| `examples/expected_AllElements_1_7_MigrationComparison.xlsx` | The auditor's output for these three inputs. |
| `docs/data_dictionary.md` | Column-by-column description of every input and of the output workbook. |

## Install

```
python -m venv .venv
.venv\Scripts\activate        (Windows)    or    source .venv/bin/activate   (macOS, Linux)
pip install -r requirements.txt
```

## Inputs

1. **Inventory CSV.** In MSOSA, Tools > Report Wizard > New template, point
   it at `templates/AllElements_Export.vm`, set the output extension to
   `.csv` (UTF-8), and select as scope exactly the packages you migrated.
   The 33-column TSV produced by the companion `tsv-to-sysmlv2` template is
   accepted as well.
2. **Not-migrated report.** `MigrationReportOf<project>.xlsx` from the CATIA
   plugin, or the report written by `tsv_to_sysmlv2.py`.
3. **SysML v2 source text.** The migrated model as `.sysml` or `.txt`.

All three must come from the same project and scope.

## Run

Build the workbook and run the comparison in one step:

```
python src/compare_migration.py path/to/AllElements_1_7.csv --catia path/to/MigrationReportOfX.xlsx --v2 path/to/Model.sysml
```

This writes `AllElements_1_7_MigrationComparison.xlsx` next to the CSV. If
the report and the v2 file sit next to the CSV and are the only candidates
there, `--catia` and `--v2` can be omitted. Without them the script builds
the workbook, pre-fills the CONFIG sheet, and tells you to rerun on the
workbook once the two yellow CONFIG cells are correct:

```
python src/compare_migration.py path/to/AllElements_1_7_MigrationComparison.xlsx
```

Rerunning on an existing workbook is safe: the result sheets are rebuilt.

## Smoke test

```
cd examples
python ../src/compare_migration.py AllElements_1_7.csv --catia MigrationReportOfHomegrown_Claude.xlsx --v2 Homegrown_generated_pilot.txt
```

The console must end with:

```
Excluded metatypes: 15 element(s) set aside on EXCLUDED_METATYPES; 29 remain in ALL_ELEMENTS_1_7.
  MIGRATED_SUCCESS:  15 elements confirmed in the v2 source (15 by qualified name, 0 by name)
  MISSED_MIGRATIONS: 2 rows (0 MISSING, 0 REFERENCE_ONLY, 2 UNVERIFIABLE_UNNAMED)
  Excluded (already in CATIA not-migrated report): 12
Appended per-metatype section to MIGRATION_SUMMARY: 69 metatype rows (header at row 23).
```

The new `AllElements_1_7_MigrationComparison.xlsx` matches
`expected_AllElements_1_7_MigrationComparison.xlsx` on every sheet except
the Generated timestamp on MIGRATION_SUMMARY.

## Output workbook

| Sheet | Contents |
|---|---|
| CONFIG | The two input paths (yellow cells). |
| ALL_ELEMENTS_1_7 | The inventory, minus the excluded metatypes. |
| EXCLUDED_METATYPES | Rows set aside before any metric is computed: Boolean, Element, Integer and String Tagged Values, Diagrams, and Literal Unlimited Naturals. |
| MIGRATED_SUCCESS | Every v1.7 element found in the v2 text, with the v2 construct it became, its v2 qualified name, line numbers, match type, and lost properties noted by the report. |
| MISSED_MIGRATIONS | Elements with no trace in the v2 text that the report also does not list. MISSING (red) are the real findings; REFERENCE_ONLY (amber) appear only as references; UNVERIFIABLE_UNNAMED (blue) are unnamed elements that name matching cannot judge. |
| MIGRATION_SUMMARY | Counts and percentages as live Excel formulas, a note on how many elements were excluded, and a per-metatype section (Element Metatype, Quantity, % Migration Success). |

Column definitions are in `docs/data_dictionary.md`.

## AI assistance

Portions of `src/compare_migration.py` were developed with assistance from Anthropic's
Claude (Claude Code). All AI-assisted code was specified, reviewed, tested,
and validated by the author. This statement is provided in accordance with
the AI Policy for George Washington University Online Engineering Programs.

## License

Code: MIT (see `LICENSE`). Templates and example data: CC BY 4.0 (see
`DATA_LICENSE.md`).

## How to cite

See `CITATION.cff`, or use the "Cite this repository" button on GitHub.
The archived release is at https://doi.org/10.5281/zenodo.XXXXXXX.

## Author

Evan Matthews, ORCID https://orcid.org/0009-0005-2265-0754
