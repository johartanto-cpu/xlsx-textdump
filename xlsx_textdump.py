"""Dump the logic of Excel workbooks to a plain-text file for Git diffs.

For each Tool.xlsx / Tool.xlsm, writes Tool.xlsx.txt next to it containing:
  - defined names (named ranges)
  - per sheet: formulas and typed constants (cached formula results excluded)
  - data validation rules
  - conditional formatting rules

Output is deterministic (no timestamps, sorted, LF line endings) so the diff
only changes when the workbook's logic changes.

Usage:
  xlsx-textdump FILE [FILE ...]          write/refresh dumps; exit 1 if any changed
  xlsx-textdump --check FILE [FILE ...]  don't write; exit 1 if any dump is stale
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import warnings
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

DUMP_SUFFIX = ".txt"
SUPPORTED = {".xlsx", ".xlsm"}


def dump_path(workbook: Path) -> Path:
    return workbook.with_name(workbook.name + DUMP_SUFFIX)


def fmt_value(v) -> str | None:
    """Format a cell value. Returns None for empty cells."""
    if v is None:
        return None
    if isinstance(v, ArrayFormula):
        return f"{{array {v.ref}}} {v.text}"
    if isinstance(v, DataTableFormula):
        attrs = {k: getattr(v, k, None) for k in ("ref", "r1", "r2", "dt2D", "dtr")}
        return "{data table} " + json.dumps({k: x for k, x in attrs.items() if x})
    if isinstance(v, str):
        if v.startswith("="):
            return v  # formula
        if v == "":
            return None
        return json.dumps(v, ensure_ascii=False)  # quoted text constant
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return repr(v)
    if isinstance(v, (int,)):
        return str(v)
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return f"#{v.isoformat()}"
    if isinstance(v, dt.timedelta):
        return f"#duration {v}"
    return f"{type(v).__name__}: {v!r}"


def dump_defined_names(wb) -> list[str]:
    lines = []
    entries = []
    for name, dn in wb.defined_names.items():
        entries.append((name, "(workbook)", dn.attr_text))
    for ws in wb.worksheets:
        for name, dn in ws.defined_names.items():
            entries.append((name, ws.title, dn.attr_text))
    for name, scope, ref in sorted(entries, key=lambda e: (e[0].lower(), e[1])):
        lines.append(f"{name} [{scope}] = {ref}")
    return lines


def dump_cells(ws) -> list[str]:
    lines = []
    for row in ws.iter_rows():
        for cell in row:
            s = fmt_value(cell.value)
            if s is not None:
                lines.append(f"{cell.coordinate}: {s}")
    return lines


NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
}
# Attributes that are presentation or internal IDs, not logic -> skipped to avoid noise
SKIP_ATTRS = {"dxfId", "id", "sqref"}
# Child elements that are presentation only
SKIP_TAGS = {"color", "dxf", "extLst", "fillColor", "negativeFillColor", "borderColor",
             "negativeBorderColor", "axisColor", "sqref"}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


# OOXML attribute defaults. An attribute at its default is omitted, so files written
# by Excel (which omits defaults) and by other tools (which write them out) dump identically.
DEFAULTS = {
    # data validation
    "type": "none", "errorStyle": "stop", "imeMode": "noControl", "operator": "between",
    "allowBlank": "0", "showDropDown": "0", "showErrorMessage": "0", "showInputMessage": "0",
    # conditional formatting rules
    "aboveAverage": "1", "bottom": "0", "equalAverage": "0", "percent": "0", "stopIfTrue": "0",
    # cfvo / colour scale / data bar / icon set
    "gte": "1", "showValue": "1", "reverse": "0", "minLength": "10", "maxLength": "90",
    "iconSet": "3TrafficLights1",
}
# Attributes that only mean something for particular rule/validation types
ONLY_FOR_TYPE = {
    "aboveAverage": {"aboveAverage"}, "equalAverage": {"aboveAverage"}, "stdDev": {"aboveAverage"},
    "rank": {"top10"}, "percent": {"top10"}, "bottom": {"top10"},
    "text": {"containsText", "notContainsText", "beginsWith", "endsWith"},
    "timePeriod": {"timePeriod"},
}
NO_OPERATOR_TYPES = {"list", "custom", "none", "any"}  # DV types where operator is ignored
BOOL = {"true": "1", "false": "0"}


def _norm(name: str, value: str) -> str:
    """Normalise true/false -> 1/0 for boolean attributes only."""
    if DEFAULTS.get(name) in ("0", "1"):
        return BOOL.get(value.lower(), value)
    return value


def _attrs(el) -> str:
    """Logic-bearing attributes only: defaults, empties and irrelevant ones dropped."""
    a = {_local(k): _norm(_local(k), v) for k, v in el.attrib.items() if _local(k) not in SKIP_ATTRS}
    kind = _local(el.tag)
    typ = a.get("type")
    if kind == "cfvo" and typ in ("min", "max"):
        a.pop("val", None)  # min/max thresholds have no value
    if kind == "dataValidation" and typ in NO_OPERATOR_TYPES:
        a.pop("operator", None)
    if kind == "cfRule":
        for attr, types in ONLY_FOR_TYPE.items():
            if typ not in types:
                a.pop(attr, None)
    items = sorted((k, v) for k, v in a.items() if v != "" and DEFAULTS.get(k) != v)
    return " ".join(f"{k}={v}" for k, v in items)


def _dv_skip(el) -> set[str]:
    """formula2 only applies to between/notBetween on numeric-style validations."""
    typ = el.get("type", "none")
    op = el.get("operator", "between")
    if typ in NO_OPERATOR_TYPES or op not in ("between", "notBetween"):
        return {"formula2"}
    return set()


def _children(el, skip: set[str] = frozenset()) -> str:
    """Flatten logic-bearing children: formulas and cfvo thresholds etc."""
    parts = []
    for ch in el:
        tag = _local(ch.tag)
        if tag in SKIP_TAGS or tag in skip:
            continue
        if tag in ("formula", "formula1", "formula2", "f"):
            # x14 wraps formulas as <x14:formula1><xm:f>..</xm:f></x14:formula1>
            f = ch.find("xm:f", NS)
            text = (f.text if f is not None else ch.text) or ""
            parts.append(f"{tag}: {text}")
            continue
        inner = _children(ch)
        a = _attrs(ch)
        label = tag + (f"({a})" if a else "")
        parts.append(label + (f" [{inner}]" if inner else ""))
    return " | ".join(parts)


def _sqref(el) -> str:
    if el.get("sqref"):
        return el.get("sqref")
    s = el.find("xm:sqref", NS)
    return s.text if s is not None and s.text else "?"


def sheet_xml_paths(workbook: Path) -> dict[str, str]:
    """Map sheet title -> worksheet XML path inside the zip."""
    with zipfile.ZipFile(workbook) as z:
        wbx = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels.findall("rel:Relationship", NS)}
    out = {}
    for sh in wbx.findall("m:sheets/m:sheet", NS):
        t = targets.get(sh.get(f"{{{NS['r']}}}id"), "")
        out[sh.get("name")] = t.lstrip("/") if t.startswith("/") else "xl/" + t
    return out


def dump_validations(root) -> list[str]:
    rows = []
    found = [(dv, "") for dv in root.findall("m:dataValidations/m:dataValidation", NS)]
    found += [(dv, " (x14)") for dv in
              root.findall("m:extLst/m:ext/x14:dataValidations/x14:dataValidation", NS)]
    for dv, tag in found:
        body = _children(dv, _dv_skip(dv))
        rows.append(f"{_sqref(dv)}: {_attrs(dv)}" + (f" | {body}" if body else "") + tag)
    return sorted(rows)


def dump_conditional_formatting(root) -> list[str]:
    rows = []
    for cf in root.findall("m:conditionalFormatting", NS):
        for rule in cf.findall("m:cfRule", NS):
            rows.append((int(rule.get("priority", 0)), _sqref(cf), rule, ""))
    for cf in root.findall("m:extLst/m:ext/x14:conditionalFormattings/x14:conditionalFormatting", NS):
        for rule in cf.findall("x14:cfRule", NS):
            rows.append((int(rule.get("priority", 0)), _sqref(cf), rule, " (x14)"))
    out = []
    for prio, rng, rule, tag in sorted(rows, key=lambda r: (r[0], r[1], _attrs(r[2]))):
        body = _children(rule)
        out.append(f"{rng}: {_attrs(rule)}" + (f" | {body}" if body else "") + tag)
    return out


def render(workbook: Path) -> str:
    with warnings.catch_warnings():
        # openpyxl warns when it drops rules it can't parse; we read DV/CF from XML ourselves
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(workbook, data_only=False, keep_links=False)
    xml_paths = sheet_xml_paths(workbook)
    zf = zipfile.ZipFile(workbook)
    out = [f"# {workbook.name}", "# Generated by xlsx-textdump. Do not edit by hand.", ""]

    out.append("## Defined names")
    out.extend(dump_defined_names(wb) or ["(none)"])
    out.append("")

    for ws in wb.worksheets:
        state = "" if ws.sheet_state == "visible" else f" [{ws.sheet_state}]"
        out.append(f"## Sheet: {ws.title}{state}")
        out.append("")
        out.append("### Cells")
        out.extend(dump_cells(ws) or ["(empty)"])
        out.append("")
        root = ET.fromstring(zf.read(xml_paths[ws.title]))
        dv = dump_validations(root)
        if dv:
            out.append("### Data validation")
            out.extend(dv)
            out.append("")
        cf = dump_conditional_formatting(root)
        if cf:
            out.append("### Conditional formatting")
            out.extend(cf)
            out.append("")

    for cs in getattr(wb, "chartsheets", []):
        out.append(f"## Chart sheet: {cs.title} (contents not dumped)")
        out.append("")

    zf.close()
    return "\n".join(out).rstrip("\n") + "\n"


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("files", nargs="*", type=Path)
    p.add_argument("--check", action="store_true", help="don't write; fail if any dump is missing or stale")
    args = p.parse_args(argv)

    rc = 0
    for f in args.files:
        if f.suffix.lower() not in SUPPORTED or f.name.startswith("~$"):
            continue
        if not f.exists():
            continue
        try:
            text = render(f)
        except Exception as e:  # corrupt/password-protected workbook etc.
            print(f"xlsx-textdump: ERROR reading {f}: {e}", file=sys.stderr)
            rc = 1
            continue

        target = dump_path(f)
        current = target.read_text(encoding="utf-8") if target.exists() else None
        if current == text:
            continue

        rc = 1
        if args.check:
            print(f"xlsx-textdump: STALE {target} (run xlsx-textdump {f})", file=sys.stderr)
        else:
            with open(target, "w", encoding="utf-8", newline="\n") as fh:
                fh.write(text)
            print(f"xlsx-textdump: updated {target} -- git add it and commit again")
    return rc


if __name__ == "__main__":
    sys.exit(main())
