"""Erzeugt die Claude-Sheets-Edits, die Excel-Formularelemente ersetzen.

Excel-Checkboxen und -Kombinationsfelder liegen als Zeichenobjekte über den Zellen und
schreiben ihren Zustand in eine (meist ausgeblendete) Hilfszelle. Claude Sheets kennt
solche Objekte nicht. Darum:

* Checkbox  -> echte Zell-Checkbox in der Zelle, über der sie in Excel lag;
               die Hilfszelle bekommt die Formel =<diese Zelle>.
* Auswahl   -> Zell-Dropdown mit den Einträgen des Excel-Listenbereichs;
               die Hilfszelle bekommt =IFERROR(MATCH(<Zelle>,<Liste>,0),<Standard>).

Die Zielzellen ergeben sich aus extract_controls.py (Ankerpositionen); Sonderfälle
stehen unten in OVERRIDES. Ausgabe: JSON-Liste von applyEdits-Paketen je Blatt.

Aufruf:
    python3 -I tools/build_controls.py VORLAGE.xlsx controls.json > control_edits.json
"""

import json
import re
import sys
import warnings

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

warnings.filterwarnings("ignore")

STAMM = "Stammblatt"
KALK = "Kalkulation Herstellungskosten"

# Zielzelle je verknüpfter Zelle (Blatt!Zelle -> Blatt, Zelle), wo die Excel-Position
# nicht eindeutig ist oder auf Text/verbundene Zellen fällt.
STAMMBLATT_TARGETS = {
    "Stammblatt!I7": "B7", "Stammblatt!I8": "B8", "Stammblatt!I9": "B9",
    "Stammblatt!I10": "F9",
    f"{KALK}!AW55": "F13",   # Regie durch Produzent*in
    "Stammblatt!I14": "F14",  # Debutfilm
    f"{KALK}!AW17": "F15",   # Drehbuch durch Produzent*in
    "Stammblatt!I16": "F16",  # Debut (Drehbuch)
    f"{KALK}!AW33": "F17",   # Herstellungsleitung
    f"{KALK}!AW35": "F18",   # Produktionsleitung intern
    f"{KALK}!AW53": "F19",   # Filmgeschäftsführung intern
    "Stammblatt!I30": "E30", "Stammblatt!I31": "E31",
}
FIN_TARGETS = {"Finanzierungsplan!Y22": "B20", "Finanzierungsplan!Y23": "B21",
               "Finanzierungsplan!AA22": "R20", "Finanzierungsplan!AA23": "R21"}


def norm_link(ctrl):
    link = (ctrl["link"] or "").replace("$", "")
    if not link or "#REF" in link:
        return None
    if "!" in link:
        sheet, cell = link.rsplit("!", 1)
        sheet = sheet.strip("'")
    else:
        sheet, cell = ctrl["sheet"], link
    return f"{sheet}!{cell}"


def ref(sheet, cell, from_sheet):
    if sheet == from_sheet:
        return cell
    return f"'{sheet}'!{cell}"


def list_labels(ws, rng):
    rng = rng.replace("$", "").split("!")[-1]
    a, b = rng.split(":")
    ca, ra = coordinate_from_string(a)
    cb, rb = coordinate_from_string(b)
    out = []
    for r in range(ra, rb + 1):
        v = ws.cell(r, column_index_from_string(ca)).value
        out.append(None if v is None else str(v).strip())
    return out


def compact(edits, existing=lambda sheet, cell: None, max_gap=1):
    """Fasst Einzelzellen-Schreibvorgänge untereinanderliegender Zellen zu Blöcken zusammen
    (Claude Sheets erlaubt höchstens 64 Edits pro Update). Lücken bis max_gap Zeilen werden
    mit dem vorhandenen Zellinhalt gefüllt, damit nichts überschrieben wird."""
    singles, rest = [], []
    for e in edits:
        if (e["kind"] == "writeCells" and len(e["rows"]) == 1 and len(e["rows"][0]) == 1
                and isinstance(e["rows"][0][0], str) and e["rows"][0][0].startswith("=")):
            singles.append(e)
        else:
            rest.append(e)
    cells = {}
    for e in singles:
        col, row = coordinate_from_string(e["topLeft"])
        cells[(e["sheet"], col, row)] = e["rows"][0][0]
    blocks = []
    for (sheet, col, row) in sorted(cells, key=lambda k: (k[0], column_index_from_string(k[1]), k[2])):
        last = blocks[-1] if blocks else None
        if last and last["sheet"] == sheet and last["_col"] == col and row - last["_end"] - 1 <= max_gap:
            for gap_row in range(last["_end"] + 1, row):
                last["rows"].append([existing(sheet, f"{col}{gap_row}")])
            last["rows"].append([cells[(sheet, col, row)]])
            last["_end"] = row
        else:
            blocks.append({"kind": "writeCells", "sheet": sheet, "topLeft": f"{col}{row}",
                           "rows": [[cells[(sheet, col, row)]]], "_col": col, "_end": row})
    for b in blocks:
        del b["_col"], b["_end"]
    # Struktur (verbinden, Checkbox-Format, Dropdowns) zuerst, dann die Werte
    return rest + blocks


def main(src, controls_path):
    wb = openpyxl.load_workbook(src)
    cached = openpyxl.load_workbook(src, data_only=True)
    ctrls = json.load(open(controls_path))

    def value_of(link):
        sheet, cell = link.split("!")
        v = cached[sheet][cell].value
        return v is True or (isinstance(v, str) and v.upper() == "TRUE")

    checkboxes = {}   # link -> (sheet, cell)
    mirrors = []      # (sheet, cell, primary sheet, primary cell)
    dropdowns = []    # dict
    seen = set()
    for c in ctrls:
        link = norm_link(c)
        if link is None:
            continue  # Excel-Element ohne gültige Verknüpfung (#REF!)
        if c["type"] == "CheckBox":
            if c["sheet"] == STAMM:
                checkboxes[link] = (STAMM, STAMMBLATT_TARGETS[link])
            elif c["sheet"] == "Finanzierungsplan":
                checkboxes[link] = ("Finanzierungsplan", FIN_TARGETS[link])
            elif c["sheet"] == "DarstellerInnen":
                row = int(re.sub(r"\D", "", link.split("!")[1]))
                checkboxes.setdefault(link, ("DarstellerInnen", f"H{row}"))
            elif c["sheet"] == KALK:
                if link == f"{KALK}!AZ82":
                    continue  # wird Dropdown "Originalton I/II", siehe unten
                fr = int(re.sub(r"\D", "", c["from"]))
                tr = int(re.sub(r"\D", "", c["to"]))
                lr = int(re.sub(r"\D", "", link.split("!")[1]))
                row = fr if tr - fr == 1 else lr
                target = (KALK, f"E{row}")
                if link in STAMMBLATT_TARGETS:
                    mirrors.append((KALK, f"E{row}", link))
                else:
                    checkboxes[link] = target
        elif c["type"] == "Drop":
            if link in seen:
                continue
            seen.add(link)
            dropdowns.append({"sheet": c["sheet"], "link": link, "list": c["listRange"], "sel": int(c["sel"] or 1)})

    per_sheet = {}

    def add(sheet, edit):
        per_sheet.setdefault(sheet, []).append(edit)

    # Stammblatt: verbundene Eingabefelder B18:F18 / B19:F19 auf B:E kürzen, damit F frei wird
    add(STAMM, {"kind": "unmergeCells", "sheet": STAMM, "ranges": ["B18:F18", "B19:F19"]})
    add(STAMM, {"kind": "mergeCells", "sheet": STAMM, "ranges": ["B18:E18", "B19:E19"]})

    by_sheet_cells = {}
    for link, (sheet, cell) in sorted(checkboxes.items()):
        by_sheet_cells.setdefault(sheet, []).append((cell, link))
    for sheet, items in by_sheet_cells.items():
        # formatCells schreibt FALSE in jede leere Zelle; nur angehakte Boxen brauchen ein TRUE
        add(sheet, {"kind": "formatCells", "sheet": sheet, "ranges": [c for c, _ in items],
                    "format": {"cellControl": {"kind": "checkbox"}}})
        for cell, link in items:
            if value_of(link):
                add(sheet, {"kind": "writeCells", "sheet": sheet, "topLeft": cell, "rows": [[True]]})
    # Hilfszellen
    for link, (sheet, cell) in sorted(checkboxes.items()):
        lsheet, lcell = link.split("!")
        add(lsheet, {"kind": "writeCells", "sheet": lsheet, "topLeft": lcell,
                     "rows": [[f"={ref(sheet, cell, lsheet)}"]]})
    # Anzeige des Stammblatt-Zustands in der Kalkulation
    for sheet, cell, link in mirrors:
        psheet, pcell = checkboxes[link]
        add(sheet, {"kind": "writeCells", "sheet": sheet, "topLeft": cell,
                    "rows": [[f"=IF('{psheet}'!{pcell}=TRUE,\"☑\",\"☐\")"]]})
        add(sheet, {"kind": "formatCells", "sheet": sheet, "ranges": [cell],
                    "format": {"alignment": {"horizontal": "center"}}})

    # Dropdowns
    green = [d for d in dropdowns if d["sheet"] == "GREEN FILMING Mehrkosten"]
    other = [d for d in dropdowns if d["sheet"] != "GREEN FILMING Mehrkosten"]
    for d in other:
        sheet = d["sheet"]
        ws = wb[sheet]
        labels = list_labels(ws, d["list"])
        cell = {"L4": "B5", "H26": "E26", "H27": "E27"}[d["link"].split("!")[1]]
        opts = [l for l in labels if l]
        add(sheet, {"kind": "makeDropdown", "sheet": sheet, "range": cell,
                    "options": [{"label": l, "color": {"kind": "none"}} for l in opts]})
        add(sheet, {"kind": "writeCells", "sheet": sheet, "topLeft": cell, "rows": [[labels[d["sel"] - 1]]]})
        rng = d["list"].replace("Stammblatt!", "")
        default = 1 if d["link"].endswith("L4") else 0
        add(sheet, {"kind": "writeCells", "sheet": sheet, "topLeft": d["link"].split("!")[1],
                    "rows": [[f"=IFERROR(MATCH({cell},{rng},0),{default})"]]})
    if green:
        sheet = "GREEN FILMING Mehrkosten"
        ws = wb[sheet]
        labels = list_labels(ws, "$L$12:$L$63")
        rows = sorted(int(re.sub(r"\D", "", d["link"].split("!")[1])) for d in green)
        add(sheet, {"kind": "makeDropdown", "sheet": sheet, "range": f"A{rows[0]}:A{rows[-1]}",
                    "options": [{"label": l, "color": {"kind": "none"}} for l in labels if l]})
        for r in rows:
            add(sheet, {"kind": "writeCells", "sheet": sheet, "topLeft": f"H{r}",
                        "rows": [[f"=IFERROR(MATCH(A{r},$L$12:$L$63,0),1)"]]})

    # Originalton I / II (Excel: Checkbox mitten im Text von B82)
    add(KALK, {"kind": "writeCells", "sheet": KALK, "topLeft": "B82", "rows": [["Originalton I"]]})
    add(KALK, {"kind": "makeDropdown", "sheet": KALK, "range": "B82", "options": [
        {"label": "Originalton I", "color": {"kind": "none"}},
        {"label": "Originalton II", "color": {"kind": "none"}}]})
    add(KALK, {"kind": "writeCells", "sheet": KALK, "topLeft": "AZ82", "rows": [["=B82=\"Originalton II\""]]})

    def existing(sheet, cell):
        v = wb[sheet][cell].value
        if isinstance(v, bool) or v is None or isinstance(v, (int, float)):
            return v
        return str(v)

    per_sheet = {s: compact(edits, existing) for s, edits in per_sheet.items()}
    json.dump(per_sheet, sys.stdout, ensure_ascii=False, indent=1)
    print(f"\nCheckboxen: {len(checkboxes)}, Anzeigen: {len(mirrors)}, Dropdowns: {len(dropdowns) + 1}",
          file=sys.stderr)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
