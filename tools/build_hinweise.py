"""Sammelt Zellkommentare und Eingabehilfen (Datenüberprüfung) einer .xlsx für ein Hinweise-Blatt.

Je Hinweis: Blatt, Zelle, Bezug (Zeilen- und Spaltenbeschriftung der Zelle), Art, Text.
Beschriftungen werden aus den zwischengespeicherten Werten gelesen, weil sie in der
Vorlage oft Formeln sind (Sprachumschaltung Deutsch/Englisch).

Aufruf:
    python3 -I tools/build_hinweise.py VORLAGE.xlsx > hinweise.json
"""

import html
import json
import re
import sys
import warnings
import zipfile

import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string

warnings.filterwarnings("ignore")

AUTHOR_PREFIX = re.compile(r"^\s*(Werner Zappe|zapp)\s*:\s*", re.I)

# Von Hand nachgeschärfte Bezüge, wo die automatische Erkennung zu ungenau ist
CONTEXT = {
    ("Stammblatt", "A1"): "Projekttitel (Kopfzeile)",
    ("Stammblatt", "B15, C15:E16"): "Drehbuch",
    ("Stammblatt", "B17:E17"): "Herstellungsleitung",
    ("Stammblatt", "B18:F18"): "Produktionsleitung",
    ("Stammblatt", "B19:F19"): "Filmgeschäftsführung",
    ("Stammblatt", "B33"): "Ko-Partner 1",
    ("Zusammenfassung", "G3"): "Spaltenkopf „Effekte“: WFBE",
    ("Zusammenfassung", "C21"): "Überschreitungsreserve (Prozentsatz)",
    ("Zusammenfassung", "D23"): "Mitteltransfer (Zu-/Abflüsse)",
    ("Kalkulation Herstellungskosten", "E14"): "Spaltenkopf „ILV“",
    ("Kalkulation Herstellungskosten", "F17"): "Drehbuch · Abzug",
    ("Kalkulation Herstellungskosten", "H17"): "Drehbuch · Richtsatz",
    ("Kalkulation Herstellungskosten", "K17"): "Drehbuch · Spalte M",
    ("Kalkulation Herstellungskosten", "L17"): "Drehbuch · Spalte W",
    ("Kalkulation Herstellungskosten", "D31"): "Spaltenkopf „Wochen“",
    ("Kalkulation Herstellungskosten", "E31"): "Spaltenkopf „ILV“",
    ("Kalkulation Herstellungskosten", "F31"): "Spaltenkopf „Wochengage inkl. SZ, UEL“",
    ("Kalkulation Herstellungskosten", "K32"): "Herstellungsleitung · Spalte M",
    ("Kalkulation Herstellungskosten", "L32"): "Herstellungsleitung · Spalte W",
    ("Kalkulation Herstellungskosten", "F54"): "Regie · Abzug",
    ("Kalkulation Herstellungskosten", "H55"): "Regie · Richtsatz",
    ("Kalkulation Herstellungskosten", "BD82"): "Originalton I/II (englische Bezeichnung)",
    ("Kalkulation Herstellungskosten", "BD114"): "Innenrequisite (englische Bezeichnung)",
    ("Kalkulation Herstellungskosten", "E170"): "Spaltenkopf „ILV“ (Green Filming)",
    ("Kalkulation Herstellungskosten", "F171"): "Green Film Consultant · Gage",
    ("Kalkulation Herstellungskosten", "B173:C174"): "Green-Filming-Personal · Funktion",
    ("Kalkulation Herstellungskosten", "E273"): "Green Film Consultant · Tagsatz",
    ("Kalkulation Herstellungskosten", "G483, U483:U487"): "Kostenmindernde Erträge",
    ("BERECHNUNGSSÄTZE", "A14"): "DZ (Wien)",
    ("BERECHNUNGSSÄTZE", "C18"): "U-Bahnabgabe (Wien)",
    ("DarstellerInnen", "Q1"): "Spaltenkopf „weiblich“ (zugeklappt)",
    ("DarstellerInnen", "Q4"): "Spalte „weiblich“ (zugeklappt)",
    ("DarstellerInnen", "B4"): "Spalte „Rolle“",
    ("Stab_Cast_Ö_Reisekosten", "L5:Q5"): "Diätentage (Spalten L–Q)",
    ("Zu 10 (Reise-, Transportkosten)", "E28:H28"): "Auslandsdiäten",
}

# Was beim Umbau nach Claude Sheets anders gelöst wurde als in Excel
UMBAU = [
    ("Stammblatt", "B5", "Sprache", "Auswahlliste statt Excel-Kombinationsfeld. Die Hilfszelle L4 rechnet die Auswahl in 1 (Deutsch) bzw. 2 (Englisch) um."),
    ("Stammblatt", "B7:B9, F9, F13:F19, E30:E31", "Filmart, Förderung, Personalunion, Debut, ILV, Koproduktion",
     "Echte Checkboxen statt Excel-Formularelemente. Die bisherigen Hilfszellen (Spalte I bzw. Kalkulation AW) lesen sie per Formel."),
    ("Stammblatt", "B18:E18, B19:E19", "Produktionsleitung, Filmgeschäftsführung",
     "Eingabefeld um eine Spalte gekürzt (vorher bis F), damit die Checkbox in F Platz hat."),
    ("Stammblatt", "E26, E27", "Aufnahmeformat, Kopienformat", "Auswahllisten statt Excel-Kombinationsfelder; H26/H27 rechnen die Auswahl in die Listennummer um."),
    ("Kalkulation Herstellungskosten", "E17, E32, E34, E52, E55", "ILV bei Drehbuch, Herstellungs-/Produktionsleitung, Filmgeschäftsführung, Regie",
     "Nur Anzeige (☑/☐). In Excel gab es diese Häkchen doppelt; in Claude Sheets setzt man sie im Stammblatt (F13–F19)."),
    ("Kalkulation Herstellungskosten", "E36:E166, E171, E334:E398", "Spalte „ILV“", "Echte Checkboxen statt Excel-Formularelemente; Hilfsspalte AW liest sie per Formel."),
    ("Kalkulation Herstellungskosten", "B82", "Originalton I / II", "Auswahlliste statt Checkbox mitten im Text. „Originalton II“ entspricht dem gesetzten Häkchen (Hilfszelle AZ82)."),
    ("Kalkulation Herstellungskosten", "F264", "Produzent*innenhonorar",
     "Zirkelbezug aufgelöst: Excel rechnete das Honorar iterativ aus Kosten, die das Honorar enthalten. Die Formel berechnet jetzt direkt den gleichen Fixpunkt "
     "aus den Netto-Fertigungskosten ohne Honorar (BERECHNUNGSSÄTZE C30)."),
    ("BERECHNUNGSSÄTZE", "C27, C30", "Produzent*innenhonorar (Staffel)", "C30 = NFK ohne Produzent*innenhonorar; C27 löst die ÖFI-Staffel nach dem Honorar auf. Kein Zirkelbezug mehr nötig."),
    ("Finanzierungsplan", "B20:B21, R20:R21", "Green Bonus, Gender-Gap-Financing", "Echte Checkboxen statt Excel-Formularelemente; Y22:Y23 bzw. AA22:AA23 lesen sie per Formel."),
    ("GREEN FILMING Mehrkosten", "A15:A21", "Mehrkosten Stab · Funktion", "Auswahllisten statt Excel-Kombinationsfelder; Spalte H rechnet die Auswahl in die Listennummer um."),
    ("DarstellerInnen", "H4:H53", "Spalte „beschränkt EST-pflichtig“", "Echte Checkboxen statt Excel-Formularelemente; Spalte T liest sie per Formel."),
    ("Richt- und Höchstsätze", "A14:I67", "Produzent*innenpauschalen, Drehbuch- & Regie-Richtsätze", "In Excel nur als Bild vorhanden; hier abgetippt, %-Spalten als Formeln."),
    ("Tages- und Nächtigungsgelder", "A1:D182", "Diäten Österreich und Ausland", "In Excel nur als Bilder vorhanden; hier abgetippt (Stand September 2017)."),
    ("alle Blätter", "–", "Blattschutz", "Claude Sheets hat keinen Blattschutz. Bitte weiterhin nur die grauen bzw. weißen Eingabefelder befüllen; Änderungen lassen sich über den Versionsverlauf zurückholen."),
]


def clean(text):
    text = html.unescape(text or "").replace("_x000a_", " ").replace("_x000D_", " ")
    text = AUTHOR_PREFIX.sub("", text)
    return re.sub(r"\s+", " ", text).strip()


def is_label(v):
    return isinstance(v, str) and v.strip() and not v.strip().startswith("=") and not re.fullmatch(r"[\d.,%\s*)-]+", v.strip())


def context(ws, cell):
    col, row = coordinate_from_string(cell)
    c = column_index_from_string(col)
    merged_first = {}
    for mr in ws.merged_cells.ranges:
        for r in range(mr.min_row, mr.max_row + 1):
            for cc in range(mr.min_col, mr.max_col + 1):
                merged_first[(cc, r)] = (mr.min_col, mr.min_row)

    def val(cc, r):
        cc, r = merged_first.get((cc, r), (cc, r))
        return ws.cell(r, cc).value

    own = val(c, row)
    row_label = None
    for cc in range(c - 1, 0, -1):
        v = val(cc, row)
        if is_label(v):
            row_label = v.strip()
            break
    if row_label is None:  # Beschriftung steht oft eine Zeile darüber (2-zeilige Positionen)
        for cc in range(c - 1, 0, -1):
            v = val(cc, row - 1)
            if is_label(v):
                row_label = v.strip()
                break
    col_label = None
    for r in range(row - 1, max(0, row - 40), -1):
        v = val(c, r)
        if is_label(v):
            col_label = v.strip()
            break
    parts = []
    if is_label(own):
        parts.append(own.strip())
    if row_label and row_label not in parts:
        parts.append(row_label)
    if col_label and col_label not in parts:
        parts.append(f"Spalte „{col_label}“")
    return " · ".join(re.sub(r"\s+", " ", p) for p in parts) or "–"


def main(src):
    wb = openpyxl.load_workbook(src, data_only=True)
    zf = zipfile.ZipFile(src)
    wbxml = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target = {}
    for m in re.finditer(r"<Relationship ([^>]*)/>", rels):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        target[a["Id"]] = "xl/" + a["Target"].lstrip("/").replace("xl/", "")
    out = []
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]*)"[^>]*r:id="([^"]*)"', wbxml):
        sheet = html.unescape(name)
        ws = wb[sheet]
        path = target[rid]
        folder, fname = path.rsplit("/", 1)
        rel_path = f"{folder}/_rels/{fname}.rels"
        if rel_path in zf.namelist():
            srels = zf.read(rel_path).decode("utf-8")
            for cpath in re.findall(r'Target="\.\./(comments\d+\.xml)"', srels):
                cx = zf.read(f"xl/{cpath}").decode("utf-8")
                for m in re.finditer(r'<comment ref="([A-Z]+\d+)"[^>]*>(.*?)</comment>', cx, re.S):
                    text = clean(re.sub(r"<[^>]+>", "", m.group(2)))
                    out.append({"sheet": sheet, "cell": m.group(1), "kind": "Kommentar",
                                "context": context(ws, m.group(1)), "text": text})
        for dv in ws.data_validations.dataValidation:
            if dv.prompt and dv.showInputMessage:
                first = str(dv.sqref).split()[0].split(":")[0]
                title = clean(dv.promptTitle)
                text = clean(dv.prompt)
                out.append({"sheet": sheet, "cell": str(dv.sqref).replace(" ", ", "), "kind": "Eingabehilfe",
                            "context": context(ws, first),
                            "text": f"{title}: {text}" if title else text})
    for h in out:
        h["context"] = CONTEXT.get((h["sheet"], h["cell"]), h["context"])
    order = [ws.title for ws in wb.worksheets]
    out.sort(key=lambda h: (order.index(h["sheet"]), int(re.sub(r"\D", "", h["cell"].split(",")[0].split(":")[0])),
                            column_index_from_string(re.sub(r"\d", "", h["cell"].split(",")[0].split(":")[0]))))
    umbau = [{"sheet": s, "cell": c, "kind": "Umbau", "context": ctx, "text": t} for s, c, ctx, t in UMBAU]
    rows = [[h["sheet"], h["cell"], h["context"], h["kind"], h["text"]] for h in umbau + out]
    sheet = "Hinweise"
    end = len(rows) + 1
    edits = [
        {"kind": "addSheet", "name": sheet},
        {"kind": "moveSheets", "sheets": [sheet], "afterSheet": "Tages- und Nächtigungsgelder"},
        {"kind": "writeCells", "sheet": sheet, "topLeft": "A1", "rows": [["Blatt", "Zelle", "Bezug", "Art", "Hinweis"]]},
        {"kind": "formatCells", "sheet": sheet, "ranges": ["A1:E1"], "format": {"fillColor": {"kind": "rgb", "hex": "#f3f3f3"}}},
        {"kind": "setCellTable", "sheet": sheet, "cellTable": {"columns": "A:E", "headerRows": "1", "footerRows": None}},
        {"kind": "setSheetSettings", "sheet": sheet, "settings": {"lastFrozenRow": 1}},
    ]
    fill = [
        {"kind": "writeCells", "sheet": sheet, "topLeft": "A2", "rows": rows},
        {"kind": "makeDropdown", "sheet": sheet, "range": f"D2:D{end}", "options": [
            {"label": "Umbau", "color": {"kind": "standard", "hue": "orange", "shade": 200}},
            {"label": "Kommentar", "color": {"kind": "standard", "hue": "blue", "shade": 200}},
            {"label": "Eingabehilfe", "color": {"kind": "standard", "hue": "green", "shade": 200}}]},
        {"kind": "formatCells", "sheet": sheet, "ranges": [f"C2:C{end}", f"E2:E{end}"], "format": {"alignment": {"wrap": True}}},
        {"kind": "setColumnWidths", "sheet": sheet, "columns": "A", "widthInPixels": 190},
        {"kind": "setColumnWidths", "sheet": sheet, "columns": "B", "widthInPixels": 150},
        {"kind": "setColumnWidths", "sheet": sheet, "columns": "C", "widthInPixels": 240},
        {"kind": "setColumnWidths", "sheet": sheet, "columns": "D", "widthInPixels": 120},
        {"kind": "setColumnWidths", "sheet": sheet, "columns": "E", "widthInPixels": 420},
    ]
    json.dump({"setup": edits, "fill": fill}, sys.stdout, ensure_ascii=False, separators=(",", ":"))


if __name__ == "__main__":
    main(sys.argv[1])
