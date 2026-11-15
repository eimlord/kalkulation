"""Übersetzt die Excel-Prüfregeln der Vorlage in Claude-Sheets-Regeln.

Quelle: rules.json aus extract_rules.py (bedingte Formatierungen) und die Datenüberprüfungen
mit Wertgrenzen. Claude-Sheets-Regeln können nur die Zelle selbst oder eine Zelle derselben
Zeile prüfen; Vergleiche mit festen Zellen laufen daher über Hilfsspalten.

Ausgabe: JSON mit
  "helpers": applyEdits-Edits (Hilfsformeln, Gruppen)
  "rules":   {Blatt: [Regel, ...]} für replaceConditionalFormats

Aufruf:
    python3 -I tools/build_rules.py rules.json > rule_edits.json
"""

import json
import re
import sys

from openpyxl.utils import column_index_from_string, get_column_letter

KALK = "Kalkulation Herstellungskosten"

PINK = {"fillColor": {"kind": "rgb", "hex": "#ffc7ce"}, "font": {"color": {"kind": "rgb", "hex": "#9c0006"}}}
RED = {"fillColor": {"kind": "rgb", "hex": "#ff0000"}}


def cells_of(sqref):
    out = []
    for part in sqref.split():
        a, b = (part.split(":") + [part])[:2]
        ma, mb = re.match(r"([A-Z]+)(\d+)", a), re.match(r"([A-Z]+)(\d+)", b)
        for r in range(int(ma.group(2)), int(mb.group(2)) + 1):
            for c in range(column_index_from_string(ma.group(1)), column_index_from_string(mb.group(1)) + 1):
                out.append((c, r))
    return out


def col_runs(cells):
    """Zellen einer Spalte -> Liste 'Q8', 'Q15:Q25', ..."""
    by_col = {}
    for c, r in cells:
        by_col.setdefault(c, set()).add(r)
    out = []
    for c in sorted(by_col):
        rows = sorted(by_col[c])
        start = prev = rows[0]
        for r in rows[1:] + [None]:
            if r is not None and r == prev + 1:
                prev = r
                continue
            L = get_column_letter(c)
            out.append(f"{L}{start}" if start == prev else f"{L}{start}:{L}{prev}")
            if r is not None:
                start = prev = r
    return out


def chunks(ranges, n=20):
    return [ranges[i:i + n] for i in range(0, len(ranges), n)]


def same_row(col, condition):
    return {"tested": {"kind": "sameRowInColumn", "c": col}, "condition": condition}


def itself(condition):
    return {"tested": {"kind": "theCellItself"}, "condition": condition}


def gt(operand):
    return {"kind": "numberIs", "comparison": "greaterThan", "operand": operand}


def lt(operand):
    return {"kind": "numberIs", "comparison": "lessThan", "operand": operand}


def num(x):
    return {"kind": "number", "number": x}


def col(c):
    return {"kind": "sameRowInColumn", "c": c}


def is_bool(v):
    return {"kind": "booleanIs", "value": v}


def rule(cells, tests, fmt):
    return [{"cells": part, "effect": {"kind": "formatWhere", "allOf": tests, "format": fmt}} for part in chunks(cells)]


def main(rules_path):
    xl = json.load(open(rules_path))

    def collect(pred, exclude=()):
        s = set()
        for r in xl:
            if pred(r):
                s.update(c for c in cells_of(r["sqref"]) if c not in exclude)
        return s

    Q = column_index_from_string("Q")
    # Q > G derselben Zeile (ÖFI+-Betrag über Gesamtbetrag); Streufragmente außerhalb von Q entfallen
    q_gt_g = {c for c in collect(lambda r: r["sheet"] == KALK and r["type"] == "cellIs"
                                 and r["formulas"] and re.fullmatch(r"\$?G\$?\d+", r["formulas"][0]))
              if c[0] == Q}
    q_gt_f = collect(lambda r: r["sheet"] == KALK and r["type"] == "cellIs"
                     and r["formulas"] and re.fullmatch(r"\$F\$\d+", r["formulas"][0]))
    ctrl = collect(lambda r: r["sheet"] == "CONTROLLING-DATA" and r["type"] == "cellIs")

    gage_rows = ["F32:F53", "F56:F167"]
    dev_rows = ["J32:J53", "J56:J167"]

    rules = {
        KALK: (
            rule(col_runs(q_gt_g), [itself(gt(col("G")))], PINK)
            + rule(col_runs(q_gt_f), [itself(gt(col("F")))], PINK)
            # ILV: Abweichung zur KV-Gage über -12 % (max. KV x 1,1 x 0,8)
            + rule(dev_rows, [same_row("AW", is_bool(True)), same_row("F", gt(num(0))), itself(gt(num(-0.12)))], RED)
            + rule(["J55"], [same_row("AW", is_bool(True)), same_row("F", gt(num(0))), itself(gt(num(0.30)))], RED)
            # Eingabegrenzen der Gagen (Excel: Datenüberprüfung)
            + rule(gage_rows, [same_row("AW", is_bool(False)), itself(gt(num(0))), itself(lt(col("H")))], PINK)
            + rule(gage_rows, [same_row("AW", is_bool(True)), itself(gt(num(0))), same_row("J", gt(num(-0.12)))], PINK)
            # Vergleiche mit festen Zellen über Hilfsspalte BF
            + rule(["F17", "F54"], [same_row("BF", is_bool(True))], RED)
            + rule(["G82:G83"], [same_row("BF", is_bool(True))],
                   {"fillColor": {"kind": "rgb", "hex": "#ffc7ce"}, "font": {"color": {"kind": "rgb", "hex": "#c00000"}}})
            + rule(["G438:G439"], [itself(gt(col("BF")))], RED)
            + rule(["F171"], [itself(gt(col("AX")))], PINK)
            + rule(["G438"], [itself(gt(num(3000)))], PINK)
            + rule(["E273"], [itself(gt(num(600)))], PINK)
            + rule(["F277"], [itself(gt(num(10000)))], PINK)
            # kostenmindernde Erträge: negatives Vorzeichen erforderlich
            + rule(["G483:G487", "P483:P487", "R483:S487", "U483:V487", "Y483:AB487",
                    "AE483:AG487", "AI483:AK487", "AM483:AO487"], [itself(gt(num(0)))], PINK)
            # Excel-Ampelsymbole (relativ im Bereich) -> Farbskala
            + [{"cells": ["J283:J326"], "effect": {"kind": "colorScale",
                "lowest": {"anchor": {"kind": "smallestNumber"}, "color": {"kind": "standard", "hue": "red", "shade": 200}},
                "middle": {"anchor": {"kind": "percentile", "percentile": 50}, "color": {"kind": "standard", "hue": "yellow", "shade": 200}},
                "highest": {"anchor": {"kind": "largestNumber"}, "color": {"kind": "standard", "hue": "green", "shade": 200}}}}]
        ),
        "Zusammenfassung": (
            rule(["T24"], [itself(gt(num(0.8)))], PINK)
            + rule(["C21"], [itself(gt(num(0.1)))], PINK)
        ),
        "Stammblatt": (
            rule(["A33:A35"], [same_row("N", is_bool(False))],
                 {"font": {"bold": True, "color": {"kind": "rgb", "hex": "#669ad8"}}})
            + rule(["B33:F35"], [same_row("N", is_bool(True))],
                   {"fillColor": {"kind": "rgb", "hex": "#a6a6a6"}, "font": {"bold": True, "color": {"kind": "rgb", "hex": "#ffffff"}}})
        ),
        "CONTROLLING-DATA": rule(col_runs(ctrl), [itself(gt(col("I")))],
                                 {"fillColor": {"kind": "rgb", "hex": "#ff0000"}, "font": {"bold": True, "color": {"kind": "rgb", "hex": "#ffffff"}}}),
        "DarstellerInnen": (
            rule(["I4:I53"], [itself(gt(col("G")))], {"fillColor": {"kind": "rgb", "hex": "#ffc7ce"}})
            + rule(["I59:I60", "I62"], [itself(gt(col("AE")))], PINK)
        ),
        "GREEN FILMING Mehrkosten": rule(["D7", "N7"], [itself(gt(num(600)))], PINK),
    }

    # AW (gerade Zeilen) spiegelt die ILV-Checkbox der Position, damit Regeln dieselbe Zeile prüfen können
    aw_rows = []
    for r in range(32, 168):
        if r % 2 == 0 and r != 54:
            aw_rows.append((r, f"=AW{r + 1}"))
    helpers = [
        {"kind": "writeCells", "sheet": KALK, "topLeft": f"AW{r}", "rows": [[f]]} for r, f in aw_rows
    ]
    helpers += [
        {"kind": "writeCells", "sheet": KALK, "topLeft": "BF16", "rows": [["Hilfswerte Prüfregeln"], ["=$F$6>15000"]]},
        {"kind": "writeCells", "sheet": KALK, "topLeft": "BF54", "rows": [["=$F$7>15000"]]},
        {"kind": "writeCells", "sheet": KALK, "topLeft": "BF82", "rows": [["=AND(Stammblatt!$I$8=TRUE,$AZ$82=FALSE)"], ["=AND(Stammblatt!$I$8=TRUE,$AZ$82=FALSE)"]]},
        {"kind": "writeCells", "sheet": KALK, "topLeft": "BF438", "rows": [["=Stammblatt!$B$26*17"], ["=Stammblatt!$B$26*50"]]},
        {"kind": "groupColumns", "sheet": KALK, "groups": [{"columns": "BF:BF", "isCollapsed": True}]},
        {"kind": "writeCells", "sheet": "Stammblatt", "topLeft": "N33", "rows": [["=$I$31"], ["=$I$31"], ["=$I$31"]]},
        {"kind": "writeCells", "sheet": "DarstellerInnen", "topLeft": "AE58", "rows": [["Hilfswert Prüfregel"], ["=$I$19"], ["=$I$19"], [None], ["=$I$19"]]},
        {"kind": "groupColumns", "sheet": "DarstellerInnen", "groups": [{"columns": "AE:AE", "isCollapsed": True}]},
    ]
    count = sum(len(v) for v in rules.values())
    print(f"Regeln: {count}", file=sys.stderr)
    json.dump({"helpers": helpers, "rules": rules}, sys.stdout, ensure_ascii=False, separators=(",", ":"))


if __name__ == "__main__":
    main(sys.argv[1])
