"""Leitet aus Excel-Gliederungen und ausgeblendeten Zeilen/Spalten Claude-Sheets-Gruppen ab.

Claude Sheets blendet nur über Gruppen aus: eine zugeklappte Gruppe ist "ausgeblendet",
mit einem "+" zum Aufklappen. Daher:

* Excel-Gliederung (outlineLevel)  -> Gruppe je Ebene; zugeklappt, wenn alles darin
                                      ausgeblendet ist.
* ausgeblendet ohne Gliederung      -> eigene zugeklappte Gruppe (wenn sie keine
                                      Gliederungsgruppe nur teilweise überlappt).

Aufruf:
    python3 -I tools/build_groups.py VORLAGE.xlsx > group_edits.json
"""

import json
import sys
import warnings

import openpyxl
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")


def runs(indices):
    out, start, prev = [], None, None
    for i in sorted(indices):
        if start is None:
            start = prev = i
        elif i == prev + 1:
            prev = i
        else:
            out.append((start, prev))
            start = prev = i
    if start is not None:
        out.append((start, prev))
    return out


def axis_groups(levels, hidden, collapsed_marks):
    """levels: {index: outlineLevel}, hidden: set(index), collapsed_marks: set(index) mit
    Excel-Attribut collapsed (sitzt auf der Zeile/Spalte direkt nach der Gruppe)
    -> [(start, end, collapsed)]"""
    groups = []
    max_level = max(levels.values(), default=0)
    for lvl in range(1, max_level + 1):
        for a, b in runs([i for i, l in levels.items() if l >= lvl]):
            all_hidden = all(i in hidden for i in range(a, b + 1))
            if lvl == 1:
                collapsed = all_hidden
            else:
                # innere Gruppe: zugeklappt nur, wenn Excel sie selbst als zugeklappt markiert
                collapsed = all_hidden and (b + 1 in collapsed_marks or a - 1 in collapsed_marks)
            groups.append((a, b, collapsed))
    outlined = set(levels)
    for a, b in runs([i for i in hidden if i not in outlined]):
        # nicht teilweise mit einer Gliederungsgruppe überlappen
        if any((a <= ga <= b) != (a <= gb <= b) and not (ga <= a and b <= gb) for ga, gb, _ in groups):
            print(f"  übersprungen (Überlappung): {a}-{b}", file=sys.stderr)
            continue
        groups.append((a, b, True))
    # Außen vor innen, damit verschachtelte Gruppen sauber entstehen
    groups.sort(key=lambda g: (g[0], -(g[1] - g[0])))
    return groups


def main(src):
    wb = openpyxl.load_workbook(src)
    edits = {}
    for ws in wb.worksheets:
        col_levels, col_hidden, col_marks = {}, set(), set()
        for _, d in ws.column_dimensions.items():
            for i in range(d.min or 0, (d.max or 0) + 1):
                if i == 0:
                    continue
                if d.outlineLevel:
                    col_levels[i] = d.outlineLevel
                if d.hidden:
                    col_hidden.add(i)
                if d.collapsed:
                    col_marks.add(i)
        row_levels, row_hidden, row_marks = {}, set(), set()
        for r, d in ws.row_dimensions.items():
            if d.outlineLevel:
                row_levels[r] = d.outlineLevel
            if d.hidden:
                row_hidden.add(r)
            if d.collapsed:
                row_marks.add(r)
        cg = axis_groups(col_levels, col_hidden, col_marks)
        rg = axis_groups(row_levels, row_hidden, row_marks)
        sheet_edits = []
        if cg:
            sheet_edits.append({"kind": "groupColumns", "sheet": ws.title, "groups": [
                {"columns": f"{get_column_letter(a)}:{get_column_letter(b)}", "isCollapsed": c} for a, b, c in cg]})
        if rg:
            sheet_edits.append({"kind": "groupRows", "sheet": ws.title, "groups": [
                {"rows": f"{a}:{b}", "isCollapsed": c} for a, b, c in rg]})
        if sheet_edits:
            edits[ws.title] = sheet_edits
    json.dump(edits, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(sys.argv[1])
