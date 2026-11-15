"""Liest die ÖFI-Positionsliste (Kostengruppe, Nr., Bezeichnung, Einheit, Art, KV-Funktion)
aus der Vorlage und schreibt sie als JSON (Grundlage für die Google-Tabelle).

Aufruf: python3 -I google/extract_positions.py vorlage/20270826_namen_ersetzt.xlsx > google/positions.json
"""
import json, re, sys, warnings
import openpyxl
warnings.filterwarnings("ignore")

GROUPS = [(1, 4, 11), (2, 14, 27), (3, 30, 329), (4, 332, 352), (5, 355, 374), (6, 377, 390),
          (7, 393, 409), (8, 412, 441), (9, 444, 454), (10, 457, 462), (11, 465, 479), (12, 482, 489)]


def kvrefs(f):
    return [int(x) for x in re.findall(r"Kollektivvertrag!\$?[A-Z]\$?(\d+)", f or "")]


def main(src):
    wb = openpyxl.load_workbook(src)
    wv = openpyxl.load_workbook(src, data_only=True)
    ws, wsv = wb["Kalkulation Herstellungskosten"], wv["Kalkulation Herstellungskosten"]
    kv = wv["Kollektivvertrag"]
    groups, out = [], []
    for g, a, b in GROUPS:
        name = re.sub(r"\*\)|\s+", " ", str(wsv.cell(a, 55).value)).strip()
        groups.append({"nr": g, "name": re.sub(r"^\d+\.\s*", "", name)})
        for r in range(a + 1, b):
            nr = wsv.cell(r, 1).value
            if nr is None or str(nr).strip() == "":
                continue
            nr = str(nr).strip()
            bc = str(wsv.cell(r, 55).value or "").strip()
            label = re.sub(r"\s+", " ", bc or str(wsv.cell(r, 2).value or "")).strip()
            if r == 82:
                label = "Originalton"
            if label in ("", "None"):
                continue
            unit = wsv.cell(r, 5).value
            unit = str(unit).strip() if isinstance(unit, str) and unit.strip() not in ("ILV", "#DIV/0!") else ""
            nxt = str(wsv.cell(r + 1, 55).value or "")
            two_row = nxt.strip().startswith("Wochenpauschale")
            h = str(ws.cell(r, 8).value or "")
            refs = [k for k in kvrefs(h) if k != 60]
            m = re.search(r"AZ(\d+)", h)
            if m:
                refs += kvrefs(str(ws.cell(int(m.group(1)), 52).value or ""))
            if r == 82:
                refs = [50]
            names = []
            for k in refs:
                n = str(kv.cell(k, 1).value).strip()
                if n not in names:
                    names.append(n)
            if r in (250, 264):
                art = "Berechnet"
            elif r == 247:
                art = "Cast"
            elif two_row or 170 <= r <= 241:
                art = "Stab"
            else:
                art = "Kosten"
            richtsatz = {17: "Drehbuch", 54: "Regie", 55: "Regie"}.get(r, "")
            out.append({"row": r, "gruppe": g, "nr": nr, "bezeichnung": label, "einheit": unit, "art": art,
                        "kv_spielfilm": names[0] if names else "", "kv_doku": names[-1] if names else "",
                        "richtsatz": richtsatz})
            if r == 5:  # Info-Zeilen für den Abzug bei Drehbuch/Regie (Vorförderungen)
                out.append({"row": 6, "gruppe": 1, "nr": "1.1a", "bezeichnung": "davon Drehbuchkosten gesamt",
                            "einheit": "", "art": "Info", "kv_spielfilm": "", "kv_doku": "", "richtsatz": ""})
                out.append({"row": 7, "gruppe": 1, "nr": "1.1b", "bezeichnung": "davon Regiekosten gesamt",
                            "einheit": "", "art": "Info", "kv_spielfilm": "", "kv_doku": "", "richtsatz": ""})
    # Positionen mit Unterpositionen (z. B. 142 -> 142.1) sind nur Überschriften
    nrs = {x["nr"] for x in out}
    for x in out:
        if any(n.startswith(x["nr"] + ".") for n in nrs) and x["art"] in ("Kosten", "Stab"):
            x["art"] = "Überschrift"
    json.dump({"gruppen": groups, "positionen": out}, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(sys.argv[1])
