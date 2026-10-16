# xlsx-textdump

[pre-commit](https://pre-commit.com) hook that writes a plain-text dump next to each Excel workbook so changes are readable in Git diffs and on GitHub.

`Tool.xlsx` → `Tool.xlsx.txt` containing:

- **Defined names** (workbook and sheet scoped)
- **Cells, per sheet:** formulas, and typed constants (text, numbers, dates, booleans). Cached formula results are excluded, so recalculation doesn't create diff noise.
- **Data validation**, including Excel 2010+ (`x14`) rules such as cross-sheet dropdowns
- **Conditional formatting** rules: type, operator, formula, priority, thresholds. Colours and fonts are excluded.

Sheet order and hidden/very-hidden state are recorded. Output is deterministic, with no timestamps and LF line endings.

Supports `.xlsx` and `.xlsm`. `.xls` and `.xlsb` are not supported. VBA is not dumped.

## Use in a repo

`.pre-commit-config.yaml`:
```yaml
repos:
  - repo: https://github.com/johartanto-cpu/xlsx-textdump
    rev: v0.1.0
    hooks:
      - id: xlsx-textdump
```
Then once per clone: `pip install pre-commit && pre-commit install`

Or start new tools from [excel-tool-template](https://github.com/johartanto-cpu/excel-tool-template).

## Behaviour

When a staged workbook's dump is missing or out of date, the hook rewrites it and **fails the commit**. Review the dump, `git add` it, and commit again.

## CLI

```
xlsx-textdump Tool.xlsx            # write/refresh dump; exit 1 if it changed
xlsx-textdump --check Tool.xlsx    # don't write; exit 1 if stale (for CI)
```

## Releasing changes

Edit, commit, then tag a new version, e.g. `git tag v0.1.1 && git push --tags`. Tool repos pick it up with `pre-commit autoupdate`.
