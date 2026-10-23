# Data dictionary: compare-migration

## Input 1: inventory CSV from `templates/AllElements_Export.vm`

Comma-separated, UTF-8, quoted fields, one header row, one row per element or diagram.

| # | Column | Meaning |
|---|---|---|
| 1 | Element MetaType | Metaclass or stereotype label, for example Block, Part Property, Diagram. |
| 2 | Element Name | Element name. Line breaks are exported as the token `<NL>`. |
| 3 | Qualified Name | Full `::`-separated qualified name. |
| 4 | Element ID | MagicDraw element ID; used for exact matching against the report. |
| 5 | Kind | `Element` or `Diagram`. |
| 6 | Owner Qualified Name | Qualified name of the owning element. |
| 7 | Applied Stereotypes | Applied stereotype names, if any. |

The 33-column TSV from the `tsv-to-sysmlv2` template is accepted in place of the CSV;
its `SysML 1.7 Element Type`, `Name`, `Qualified Name`, `Element ID`, `Owner Qualified Name`
and `Applied Stereotype` columns map onto the seven columns above.

## Input 2: not-migrated report (`MigrationReportOf<project>.xlsx`)

Written by the CATIA Magic migration plugin or by `tsv_to_sysmlv2.py`. Elements are
matched by the `Element ID` column, or by the ID inside the `mdel://` link in
`Element Url` when the ID column is absent. Sheets and columns of the example report:

- `Not Migrated Elements`: `Element MetaType` | `Element Name` | `Qualified Name` | `Element Url` | `Element ID`
- `Not Migrated Properties`: `Element Name` | `Element MetaType` | `Property Name` | `Qualified Name` | `Element Url`
- `Not Migrated Diagrams`: `Diagram Type` | `Element Name` | `Qualified Name` | `Element Url` | `Element ID`

## Input 3: SysML v2 source text

The migrated model as `.sysml` or `.txt`. Parsed by brace nesting into qualified names;
comments and `doc` bodies are ignored; quoted names such as `'Op-Commands'` are unquoted
before matching.

## Output: comparison workbook (`<inventory name>_MigrationComparison.xlsx`)

### Sheet `CONFIG`

B3 = path to the not-migrated report; B4 = path to the SysML v2 source. Yellow cells.

### Sheet `ALL_ELEMENTS_1_7`

Columns: `Element MetaType` | `Element Name` | `Qualified Name` | `Element ID` | `Kind` | `Owner Qualified Name` | `Applied Stereotypes`

### Sheet `MIGRATED_SUCCESS`

Columns: `Element MetaType (v1.7)` | `Element Name` | `v1.7 Qualified Name` | `Element ID` | `Migrated To (v2 construct)` | `v2 Qualified Name` | `v2 Line(s)` | `Match Type` | `Notes`

### Sheet `MISSED_MIGRATIONS`

Columns: `Category` | `Element MetaType (v1.7)` | `Element Name` | `v1.7 Qualified Name` | `Element ID` | `Reason`

### Sheet `MIGRATION_SUMMARY`

Columns: `Metric` | `Count` | `% of v1.7 elements`

### Sheet `EXCLUDED_METATYPES`

Columns: `Element MetaType` | `Element Name` | `Qualified Name` | `Element ID` | `Kind` | `Owner Qualified Name` | `Applied Stereotypes`

### Value definitions

- `Category` on MISSED_MIGRATIONS: `MISSING` (named element, no trace in the v2 text,
  not in the report), `REFERENCE_ONLY` (the name appears in the v2 text only as a
  reference, never as a declaration), `UNVERIFIABLE_UNNAMED` (unnamed v1.7 element;
  name matching cannot prove or disprove it).
- `Match Type` on MIGRATED_SUCCESS: `Qualified name` (exact), `Name (unique)`,
  `Name (ambiguous)`.
- `Notes` on MIGRATED_SUCCESS: flags elements that the report also lists, and names
  properties the report says were lost.
- MIGRATION_SUMMARY counts are live `COUNTIF`/`COUNTA` formulas over the other sheets,
  followed by static parse statistics, the input file paths, a Generated timestamp, a
  note on how many rows were excluded, and the per-metatype section
  (`Element Metatype`, `Quantity`, `% Migration Success`).
- Excluded metatypes (moved to EXCLUDED_METATYPES before any metric): Boolean Tagged
  Value, Diagram, Element Tagged Value, Integer Tagged Value, Literal Unlimited Natural,
  String Tagged Value.
