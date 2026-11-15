"""Erzeugt Inhalt, Formeln und Formatvorgaben für die interne Kalkulation als Google-Tabelle.

Ausgabe (JSON): je Blatt die Zellen (Text, wie in Google Tabellen mit deutschem Gebietsschema
eingetippt: Formeln mit ";" und Dezimalkomma), Zellstile für den HTML-Einfügeweg und eine
Liste von Bedienschritten (Checkboxen, Dropdowns, Zahlenformate, Fixierung), die
google/fill.js im Browser ausführt.

Aufruf:
    python3 -I google/build_sheets.py vorlage/20270826_EMPTY_TEMPLATE.xlsx google/positions.json > google/sheets.json
"""

import json
import sys
import warnings

import openpyxl

warnings.filterwarnings("ignore")

BUNDESLAENDER = [("Burgenland", "0,40%"), ("Kärnten", "0,37%"), ("Niederösterreich", "0,33%"),
                 ("Oberösterreich", "0,31%"), ("Salzburg", "0,35%"), ("Steiermark", "0,34%"),
                 ("Tirol", "0,39%"), ("Vorarlberg", "0,33%"), ("Wien", "0,36%")]

HEAD = {"bold": True, "bg": "#f3f3f3"}
GROUP = {"bold": True, "bg": "#e8f0fe"}
CALC = {"bg": "#f8f9fa"}
TITLE = {"bold": True, "size": 14}
NOTE = {"color": "#5f6368", "italic": True}


def q(s):
    return f"'{s}'"


class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}   # (r, c) -> text
        self.styles = {}  # (r, c) -> style dict
        self.ops = []

    def set(self, r, c, text, style=None):
        self.cells[(r, c)] = text
        if style:
            self.styles[(r, c)] = {**self.styles.get((r, c), {}), **style}

    def style(self, r, c, style):
        self.styles[(r, c)] = {**self.styles.get((r, c), {}), **style}

    def grid(self):
        maxr = max(r for r, _ in self.cells)
        maxc = max(c for _, c in self.cells)
        return [[self.cells.get((r, c), "") for c in range(1, maxc + 1)] for r in range(1, maxr + 1)]

    def dump(self):
        return {"name": self.name, "rows": self.grid(),
                "styles": [[r, c, s] for (r, c), s in self.styles.items()], "ops": self.ops}


def col(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def build(template, positions_path):
    wv = openpyxl.load_workbook(template, data_only=True)
    pos = json.load(open(positions_path))

    # ------------------------------------------------------------ Projekt
    pr = Sheet("Projekt")
    pr.set(1, 1, "Projekt", TITLE)
    fields = [
        ("Projekttitel", "", ""),
        ("Produktionsfirma", "", ""),
        ("Filmart", "Spielfilm", "Spielfilm oder Dokumentarfilm (bestimmt die KV-Funktion bei HL, PL, AL, Kamera)"),
        ("Werkstattprojekt (§ 19 KV)", "", "Werkstatt-KV bis HK 1,22 Mio. €, halber KV bis 1,57 Mio. €"),
        ("gefördert durch BMWKMS/if", "", "FGK 7,5 % der NFK; Produzent*innenhonorar 5 % der HK"),
        ("Debutfilm (Regie)", "", "Regie-Richtsatz 85 %"),
        ("Debut (Drehbuch)", "", "Drehbuch-Richtsatz 85 %"),
        ("Bundesland der Betriebsstätte", "Wien", "bestimmt DZ-Satz; U-Bahn-Abgabe nur bei Wien"),
        ("Stichtag für Sätze", "01.01.2027", "z. B. Drehbeginn: gilt für KV-Gagen und alle Sätze"),
        ("Überschreitungsreserve", "5%", "max. 8 %, mit Completion Bond 10 %"),
        ("Completion Bond", "", ""),
        ("Finanzierungskosten", "", ""),
        ("Kostenstand per", "", ""),
    ]
    P = {}
    for i, (label, value, hint) in enumerate(fields):
        r = 3 + i
        pr.set(r, 1, label, {"bold": True})
        pr.set(r, 2, value)
        pr.set(r, 3, hint, NOTE)
        P[label] = f"'Projekt'!$B${r}"
    FILMART, WERKSTATT, BMWKMS = P["Filmart"], P["Werkstattprojekt (§ 19 KV)"], P["gefördert durch BMWKMS/if"]
    DEBUT_RE, DEBUT_DB = P["Debutfilm (Regie)"], P["Debut (Drehbuch)"]
    LAND, STICHTAG, UER = P["Bundesland der Betriebsstätte"], P["Stichtag für Sätze"], P["Überschreitungsreserve"]
    BOND, FIN = P["Completion Bond"], P["Finanzierungskosten"]
    pr.ops += [
        {"op": "checkbox", "range": "B6:B9"},
        {"op": "dropdown", "range": "B5", "options": ["Spielfilm", "Dokumentarfilm"]},
        {"op": "dropdown", "range": "B10", "options": [b for b, _ in BUNDESLAENDER]},
        {"op": "percent", "range": "B12"},
        {"op": "currency", "range": "B13:B14"},
    ]

    # Sätze (Schlüssel, Bezeichnung, Bundesland, Wert, Einheit, gültig ab, Quelle)
    S = [
        ("SV_HBGL", "SV-Höchstbeitragsgrundlage laufend (monatlich)", "", "6930", "€", "01.01.2026", "ÖFI-Vorlage 2027 (vorläufig)"),
        ("SV_LFD", "DG-Anteil SV laufend (AV 2,95 + KV 3,78 + UV 1,10 + PV 12,55 + IESG 0,10 + WF 0,50)", "", "20,98%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("SV_SZ", "DG-Anteil SV Sonderzahlungen (ohne WF)", "", "20,48%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("UEL", "Urlaubsersatzleistung", "", "10,41%", "", "01.01.2026", "ÖFI-Rechenlogik"),
        ("DB", "Dienstgeberbeitrag (FLAF)", "", "3,7%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
    ] + [("DZ", "Zuschlag zum DB", land, satz, "", "01.01.2026", "WKO, Hebesätze 2026") for land, satz in BUNDESLAENDER] + [
        ("KOMMST", "Kommunalsteuer", "", "3%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("MV", "Betriebliche Vorsorge (MV-Beitrag)", "", "1,53%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("UBAHN", "Dienstgeberabgabe Wien (U-Bahn-Abgabe) je Woche", "Wien", "2", "€", "01.06.2012", "Stadt Wien"),
        ("F_SZUEL", "Faktor SZ/UEL auf KV-Grundgage", "", "1,2881166", "", "01.01.2026", "ÖFI-Vorlage (Kollektivvertrag)"),
        ("F_60H", "Faktor Wochenpauschale 60 h (§ 7 KV)", "", "1,385", "", "01.01.2026", "ÖFI-Vorlage (Kollektivvertrag)"),
        ("ILV", "ILV: max. Anteil der KV-Gage (KV × 1,1 × 0,8)", "", "88%", "", "01.01.2026", "ÖFI-Richtlinie"),
        ("UEBER", "Überzahlung anerkannt bis (FRL Pkt. 6.2.3)", "", "110%", "", "01.01.2026", "ÖFI-Richtlinie"),
        ("WS1", "Werkstattprojekt: HK-Grenze Werkstatt-KV", "", "1220000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("WS2", "Werkstattprojekt: HK-Grenze halber KV", "", "1570000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("GFC_W", "Green Film Consultant: Höchstsatz je Woche", "", "3000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("GFC_ILV", "Green Film Consultant: Höchstsatz je Woche bei ILV", "", "2400", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("MD_MAX", "Producer for Marketing & Distribution: max. anerkannt", "", "10000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("FREIBETRAG", "Drehbuch/Regie: Freibetrag Vorförderungen", "", "15000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("DEBUT", "Richtsatz bei Debut", "", "85%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("ILV_RS", "Richtsatz Drehbuch/Regie bei ILV", "", "80%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("DIAET", "Tagesgeld Inland", "", "30", "€", "01.01.2025", "§ 26 Z 4 EStG"),
        ("NAECHT", "Nächtigungsgeld Inland (Pauschale)", "", "17", "€", "01.01.2025", "§ 26 Z 4 EStG"),
        ("KM", "Kilometergeld", "", "0,5", "€", "01.01.2025", "§ 26 Z 4 EStG"),
    ]
    keys = []
    for k, *_ in S:
        if k not in keys:
            keys.append(k)
    pr.set(17, 1, "Aktuelle Sätze zum Stichtag (automatisch, Werte im Blatt „Sätze“ pflegen)", {"bold": True})
    K = {}
    for i, k in enumerate(keys):
        r = 18 + i
        label = next(s[1] for s in S if s[0] == k)
        pr.set(r, 1, label)
        pr.set(r, 2, f"=SUMIFS('Sätze'!$D:$D;'Sätze'!$A:$A;$C{r};'Sätze'!$H:$H;TRUE())", CALC)
        pr.set(r, 3, k, NOTE)
        K[k] = f"'Projekt'!$B${r}"
    last = 18 + len(keys) - 1
    pr.ops += [{"op": "currency", "range": f"B18:B{last}"}]
    pct = [18 + keys.index(k) for k in ("SV_LFD", "SV_SZ", "UEL", "DB", "DZ", "KOMMST", "MV", "ILV", "UEBER", "DEBUT", "ILV_RS")]
    pr.ops += [{"op": "percent", "range": f"B{r}"} for r in pct]
    pr.ops += [{"op": "plain", "range": f"B{18 + keys.index(k)}"} for k in ("F_SZUEL", "F_60H")]

    # ------------------------------------------------------------ Sätze
    sa = Sheet("Sätze")
    head = ["Schlüssel", "Bezeichnung", "Bundesland", "Wert", "Einheit", "gültig ab", "Quelle", "gilt am Stichtag"]
    for c, h in enumerate(head, 1):
        sa.set(1, c, h, HEAD)
    for i, row in enumerate(S):
        r = 2 + i
        for c, v in enumerate(row, 1):
            sa.set(r, c, v)
        sa.set(r, 8, f"=AND($F{r}<={STICHTAG};OR($C{r}=\"\";$C{r}={LAND});"
                     f"$F{r}=IF($C{r}=\"\";MAXIFS($F:$F;$A:$A;$A{r};$F:$F;\"<=\"&{STICHTAG});"
                     f"MAXIFS($F:$F;$A:$A;$A{r};$C:$C;$C{r};$F:$F;\"<=\"&{STICHTAG})))", CALC)
    sa.ops += [{"op": "freeze", "rows": 1}]

    # ------------------------------------------------------------ KV
    kv = Sheet("KV")
    head = ["Funktion", "Grundgage 40 h (Woche)", "gültig ab", "Quelle", "40 h inkl. SZ/UEL",
            "60 h inkl. SZ/UEL", "Tag (1/5) inkl. SZ/UEL", "gilt am Stichtag"]
    for c, h in enumerate(head, 1):
        kv.set(1, c, h, HEAD)
    kvs = wv["Kollektivvertrag"]
    kv_rows = [(kvs.cell(r, 1).value, kvs.cell(r, 2).value) for r in range(11, 77)
               if kvs.cell(r, 1).value and isinstance(kvs.cell(r, 2).value, (int, float))]
    for i, (name, gage) in enumerate(kv_rows):
        r = 2 + i
        kv.set(r, 1, str(name).strip())
        kv.set(r, 2, f"{gage:.2f}".replace(".", ","))
        kv.set(r, 3, "01.01.2027")
        kv.set(r, 4, "KV Filmschaffende 2027 (ÖFI-Vorlage)")
        kv.set(r, 5, f"=B{r}*{K['F_SZUEL']}", CALC)
        kv.set(r, 6, f"=B{r}*{K['F_60H']}*{K['F_SZUEL']}", CALC)
        kv.set(r, 7, f"=B{r}/5*{K['F_SZUEL']}", CALC)
        kv.set(r, 8, f"=$C{r}=MAXIFS($C:$C;$A:$A;$A{r};$C:$C;\"<=\"&{STICHTAG})", CALC)
    kv_last = 1 + len(kv_rows)
    kv.ops += [{"op": "currency", "range": f"B2:B{kv_last}"}, {"op": "currency", "range": f"E2:G{kv_last}"},
               {"op": "freeze", "rows": 1}]

    # ------------------------------------------------------------ Positionen
    po = Sheet("Positionen")
    cols = ["Gr.", "Nr.", "Position", "Name / Beschreibung", "Partner", "Menge", "Einheit", "Preis",
            "Wochen 60 h", "Wochengage 60 h", "ILV", "M/W", "Plan brutto", "LNK", "Plan gesamt", "Prüfung",
            "Ist", "noch zu erwarten", "Endkosten", "± zu Plan", "Art", "KV-Funktion", "KV-Mindestgage 40 h",
            "Abw. zu KV", "ohne U-Bahn-Abgabe", "Hilfe: U-Bahn-Abgabe", "Hilfe: Menge × Preis", "Hilfe: 60 h"]
    C = {name: i + 1 for i, name in enumerate(cols)}
    L = {name: col(i) for name, i in C.items()}
    for c, h in enumerate(cols, 1):
        po.set(1, c, h, HEAD)
    po.set(2, 1, "Preis: Stab = Wochengage 40 h inkl. SZ/UEL · Cast = Tagesgage · sonst Einzelpreis (ohne Menge = Pauschale). "
                 "Neue Person/Partner: Zeile unter der Position einfügen und Nr. eintragen.", NOTE)
    r = 3
    first_data = r
    group_rows, stab_rows, person_rows, pos_rows = [], [], [], []
    by_group = {}
    for p in pos["positionen"]:
        by_group.setdefault(p["gruppe"], []).append(p)
    gnames = {g["nr"]: g["name"] for g in pos["gruppen"]}

    def F(name):
        return f"{L[name]}{{r}}"

    for g in range(1, 13):
        po.set(r, 1, str(g), GROUP)
        po.set(r, 3, f"{g}. {gnames[g]}", GROUP)
        po.set(r, C["Art"], "Gruppe", GROUP)
        for c in range(1, len(cols) + 1):
            po.style(r, c, GROUP)
        group_rows.append(r)
        r += 1
        for p in by_group.get(g, []):
            art = p["art"]
            po.set(r, 1, str(g))
            po.set(r, 2, "'" + p["nr"])
            po.set(r, 3, p["bezeichnung"])
            po.set(r, C["Art"], art)
            if art in ("Kosten", "Stab", "Cast", "Info"):
                po.set(r, C["Partner"], "AT")
                pos_rows.append(r)
            if art == "Stab":
                po.set(r, C["Einheit"], "Wochen 40 h")
                ks, kd = p["kv_spielfilm"], p["kv_doku"]
                if ks and kd and ks != kd:
                    po.set(r, C["KV-Funktion"], f"=IF({FILMART}=\"Dokumentarfilm\";\"{kd}\";\"{ks}\")")
                elif ks:
                    po.set(r, C["KV-Funktion"], ks)
                stab_rows.append(r)
                person_rows.append(r)
            elif art == "Cast":
                po.set(r, C["Einheit"], "Drehtage")
                person_rows.append(r)
            elif art == "Kosten":
                po.set(r, C["Einheit"], p["einheit"])
            elif art == "Berechnet" and p["nr"] == "84":
                po.set(r, 4, "je Person in Spalte LNK enthalten; Aufteilung siehe Zusammenfassung", NOTE)
            elif art == "Berechnet":
                po.set(r, 4, "wird automatisch berechnet (siehe Zusammenfassung)", NOTE)
            elif art == "Überschrift":
                po.set(r, 4, "Summe der Unterpositionen", NOTE)
            r += 1
    last = r - 1
    rng = lambda name: f"{L[name]}{first_data}:{L[name]}{last}"

    # Formeln je Zeile
    for rr in range(first_data, last + 1):
        f = lambda name: f"{L[name]}{rr}"
        art = f(f"Art")
        po.set(rr, C["Plan brutto"],
               f"=IF(OR({art}=\"Gruppe\";{art}=\"Überschrift\";{art}=\"Berechnet\";"
               f"AND({f('Preis')}=\"\";{f('Wochengage 60 h')}=\"\"));\"\";"
               f"IF({f('Menge')}=\"\";N({f('Preis')});{f('Menge')}*N({f('Preis')}))+N({f('Wochen 60 h')})*N({f('Wochengage 60 h')}))", CALC)
        if rr in group_rows:
            # Gruppenzeile: Summe des eigenen Blocks bis einschließlich der nächsten Gruppenzeile
            # (so zählen auch Zeilen mit, die direkt über der nächsten Gruppe eingefügt werden)
            nxt = next((g for g in group_rows if g > rr), last + 1)
            blk = lambda name: f"{L[name]}{rr + 1}:{L[name]}{nxt}"
            for name in ("Plan brutto", "LNK", "Plan gesamt", "Endkosten"):
                po.set(rr, C[name], f"=SUMIFS({blk(name)};{blk('Art')};\"<>Gruppe\";{blk('Art')};\"<>Info\")", GROUP)
            continue
        ub = f"IF({f('ohne U-Bahn-Abgabe')}=TRUE();0;{K['UBAHN']})"
        po.set(rr, C["Hilfe: U-Bahn-Abgabe"],
               f"=IF({art}=\"Stab\";(N({f('Menge')})+N({f('Wochen 60 h')}))*{ub};"
               f"IF({art}=\"Cast\";N({f('Menge')})/5*{ub};0))", CALC)
        po.set(rr, C["Hilfe: Menge × Preis"], f"=IF({f('Preis')}=\"\";0;IF({f('Menge')}=\"\";1;{f('Menge')})*{f('Preis')})", CALC)
        po.set(rr, C["Hilfe: 60 h"], f"=N({f('Wochen 60 h')})*N({f('Wochengage 60 h')})", CALC)
        stab = (f"LET(brutto;N({f('Plan brutto')});wochen;N({f('Menge')})+N({f('Wochen 60 h')});hbgl;{K['SV_HBGL']};"
                f"svl;{K['SV_LFD']};svs;{K['SV_SZ']};uel;{K['UEL']};lfd;brutto/((1+1/6)*(1+uel));sonder;lfd/6;"
                f"svlfd;IF(wochen=0;0;IF(lfd/wochen>hbgl/30*7;wochen*hbgl*svl/30*7;lfd*svl));"
                f"svsz;IF(sonder>2*hbgl;2*hbgl*svs;sonder*svs);"
                f"uellfd;IF(lfd*uel>0,5833*wochen*hbgl/30;0,5833*wochen*hbgl*svl/30;lfd*uel*svl);"
                f"uelsz;IF(sonder>2*hbgl;0;IF(sonder*uel>(2*hbgl-sonder);(2*hbgl-sonder)*svs;sonder*uel*svs));"
                f"IF(brutto=0;\"\";svlfd+svsz+uellfd+uelsz+brutto*({K['MV']}+{K['DB']}+{K['DZ']}+{K['KOMMST']})+{f('Hilfe: U-Bahn-Abgabe')}))")
        cast = (f"LET(brutto;N({f('Plan brutto')});tage;N({f('Menge')});hbgl;{K['SV_HBGL']};svl;{K['SV_LFD']};uel;{K['UEL']};"
                f"sv;IF(OR(tage=0;brutto/MAX(tage;1)<=hbgl/30);brutto*svl;tage*hbgl*svl/30);"
                f"ueltl;brutto-brutto/(1+uel);"
                f"uelsv;IF(ueltl>0,5833*tage/7*hbgl/30;0,5833*tage/7*hbgl*svl/30;ueltl*svl);"
                f"IF(brutto=0;\"\";sv+uelsv+brutto*({K['DB']}+{K['DZ']}+{K['KOMMST']})+{f('Hilfe: U-Bahn-Abgabe')}))")
        po.set(rr, C["LNK"], f"=IF({art}=\"Stab\";{stab};IF({art}=\"Cast\";{cast};\"\"))", CALC)
        po.set(rr, C["Plan gesamt"],
               f"=IF({art}=\"Berechnet\";IF({f('Nr.')}=\"87\";'Zusammenfassung'!$B$24;\"\");"
               f"IF(AND({f('Plan brutto')}=\"\";{f('LNK')}=\"\");\"\";N({f('Plan brutto')})+N({f('LNK')})))", CALC)
        kvmin = (f"=IF(OR({art}<>\"Stab\";{f('KV-Funktion')}=\"\");\"\";"
                 f"LET(kv;SUMIFS('KV'!$E:$E;'KV'!$A:$A;{f('KV-Funktion')};'KV'!$H:$H;TRUE());"
                 f"ws;SUMIFS('KV'!$E:$E;'KV'!$A:$A;\"Werkstattprojekt (§ 19 KV)\";'KV'!$H:$H;TRUE());"
                 f"hk;'Zusammenfassung'!$B$21;"
                 f"IF({WERKSTATT}=TRUE();IF(hk<={K['WS1']};ws;IF(hk<={K['WS2']};kv/2;kv));kv)))")
        po.set(rr, C["KV-Mindestgage 40 h"], kvmin, CALC)
        po.set(rr, C["Abw. zu KV"],
               f"=IF(OR({f('KV-Mindestgage 40 h')}=\"\";N({f('KV-Mindestgage 40 h')})=0;N({f('Preis')})=0);\"\";"
               f"{f('Preis')}/{f('KV-Mindestgage 40 h')}-1)", CALC)
        kvm = f"N({f('KV-Mindestgage 40 h')})"
        ilv = f"{f('ILV')}=TRUE()"
        pr_ = f"N({f('Preis')})"
        checks = [
            f"IF(AND({art}=\"Stab\";{ilv};{pr_}>0;{kvm}>0;{pr_}>{kvm}*{K['ILV']});\"ILV über 88 % KV\";\"\")",
            f"IF(AND({art}=\"Stab\";NOT({ilv});{pr_}>0;{kvm}>0;{pr_}<{kvm});\"unter KV (40 h)\";\"\")",
            f"IF(AND({art}=\"Stab\";NOT({ilv});N({f('Wochengage 60 h')})>0;{kvm}>0;"
            f"N({f('Wochengage 60 h')})<{kvm}*{K['F_60H']});\"unter KV (60 h)\";\"\")",
            f"IF(AND({art}=\"Stab\";{kvm}>0;OR({pr_}>{kvm}*{K['UEBER']};N({f('Wochengage 60 h')})>{kvm}*{K['F_60H']}*{K['UEBER']}));"
            f"\"Überzahlung über 10 % (nur mit Begründung)\";\"\")",
            f"IF(AND({f('Nr.')}=\"81.0\";{pr_}>IF({ilv};{K['GFC_ILV']};{K['GFC_W']}));\"über Höchstsatz Green Film Consultant\";\"\")",
            f"IF(AND({f('Nr.')}=\"94.4\";N({f('Plan brutto')})>{K['MD_MAX']});\"max. 10.000 € anerkannt\";\"\")",
            f"IF(AND({f('Gr.')}=12;N({f('Plan brutto')})>0);\"Erträge mit negativem Vorzeichen eintragen\";\"\")",
            f"IF(AND({f('Nr.')}=\"4\";N({f('Plan brutto')})>ROUND('Zusammenfassung'!$B$26*IF({ilv};{K['ILV_RS']};1);-1)-'Zusammenfassung'!$B$28);"
            f"\"über Richtsatz Drehbuch\";\"\")",
            f"IF(AND({f('Nr.')}=\"24\";N({f('Plan brutto')})>ROUND('Zusammenfassung'!$B$27*IF({ilv};{K['ILV_RS']};1);-1)-'Zusammenfassung'!$B$29);"
            f"\"über Richtsatz Regie\";\"\")",
            f"IF(AND({f('Nr.')}<>\"\";COUNTIF('ÖFI-Übertrag'!$B:$B;{f('Nr.')})=0);\"Nr. nicht in der ÖFI-Liste\";\"\")",
            f"IF(AND({f('Partner')}<>\"\";N({f('Plan brutto')})<>0;COUNTIF({{\"AT\";\"A\";\"B\";\"C\"}};{f('Partner')})=0);\"Partner AT, A, B oder C\";\"\")",
        ]
        po.set(rr, C["Prüfung"], "=TEXTJOIN(\", \";TRUE();" + ";".join(checks) + ")", CALC)
        po.set(rr, C["Endkosten"], f"=IF(AND({f('Ist')}=\"\";{f('noch zu erwarten')}=\"\");\"\";N({f('Ist')})+N({f('noch zu erwarten')}))", CALC)
        po.set(rr, C["± zu Plan"], f"=IF({f('Endkosten')}=\"\";\"\";N({f('Plan brutto')})-{f('Endkosten')})", CALC)
        for name in ("LNK", "Plan gesamt", "Prüfung", "KV-Mindestgage 40 h", "Abw. zu KV", "Endkosten", "± zu Plan",
                     "Hilfe: U-Bahn-Abgabe", "Hilfe: Menge × Preis", "Hilfe: 60 h"):
            po.style(rr, C[name], CALC)
        po.style(rr, C["Plan brutto"], CALC)

    g3 = [rr for rr in range(first_data, last + 1) if po.cells.get((rr, 1)) == "3" and rr not in group_rows]
    g3a, g3b = min(g3), max(g3)
    money = [rng(n) for n in ("Preis", "Wochengage 60 h", "Plan brutto", "LNK", "Plan gesamt", "Ist", "noch zu erwarten",
                              "Endkosten", "± zu Plan", "KV-Mindestgage 40 h", "Hilfe: U-Bahn-Abgabe", "Hilfe: Menge × Preis", "Hilfe: 60 h")]
    po.ops += [{"op": "text", "range": rng("Nr.")}]
    po.ops += [{"op": "currency", "range": m} for m in money]
    po.ops += [{"op": "percent", "range": rng("Abw. zu KV")}]
    po.ops += [
        {"op": "checkbox", "range": f"{L['ILV']}{g3a}:{L['ILV']}{g3b}"},
        {"op": "checkbox", "range": f"{L['ohne U-Bahn-Abgabe']}{g3a}:{L['ohne U-Bahn-Abgabe']}{g3b}"},
        {"op": "dropdown", "range": rng("Partner"), "options": ["AT", "A", "B", "C"]},
        {"op": "dropdown", "range": f"{L['M/W']}{g3a}:{L['M/W']}{g3b}", "options": ["M", "W"]},
        {"op": "dropdown_range", "range": f"{L['KV-Funktion']}{g3a}:{L['KV-Funktion']}{g3b}", "source": f"='KV'!$A$2:$A${kv_last}"},
        {"op": "dropdown", "range": rng("Art"), "options": ["Gruppe", "Kosten", "Stab", "Cast", "Info", "Überschrift", "Berechnet"]},
        {"op": "freeze", "rows": 1, "cols": 3},
    ]
    PO = {name: f"'Positionen'!${L[name]}${first_data}:${L[name]}" for name in cols}

    # ------------------------------------------------------------ Zusammenfassung
    zs = Sheet("Zusammenfassung")
    zs.set(1, 1, "Zusammenfassung", TITLE)
    partners = ["AT", "A", "B", "C"]
    for c, h in enumerate(["Kostengruppe", "AT", "A", "B", "C", "Gesamt", "Endkosten (brutto)", "± zu Plan"], 1):
        zs.set(3, c, h, HEAD)
    for g in range(1, 13):
        r = 3 + g
        zs.set(r, 1, f"{g}. {gnames[g]}")
        for i, pt in enumerate(partners):
            cc = 2 + i
            base = (f"SUMIFS({PO['Plan brutto']};{PO['Gr.']};\"{g}\";{PO['Partner']};\"{pt}\";{PO['Art']};\"<>Info\")"
                    f"+SUMIFS({PO['LNK']};{PO['Gr.']};\"{g}\";{PO['Partner']};\"{pt}\")")
            if g == 3 and pt == "AT":
                base += "+$B$24"
            zs.set(r, cc, "=" + base, CALC)
        zs.set(r, 6, f"=SUM(B{r}:E{r})", CALC)
        zs.set(r, 7, f"=SUMIFS({PO['Endkosten']};{PO['Gr.']};\"{g}\";{PO['Art']};\"<>Info\";{PO['Art']};\"<>Gruppe\")", CALC)
        zs.set(r, 8, f"=SUMIFS({PO['± zu Plan']};{PO['Gr.']};\"{g}\";{PO['Art']};\"<>Info\";{PO['Art']};\"<>Gruppe\")", CALC)
    rows = [
        (16, "Fertigungskosten (NFK)", ["=SUM(B4:B15)", "=SUM(C4:C15)", "=SUM(D4:D15)", "=SUM(E4:E15)"]),
        (17, "Fertigungsgemeinkosten (FGK)", [
            f"=IF({BMWKMS}=TRUE();B16*7,5%;ROUND(IF(B16>=5000000;375000;IF(B16<=800000;B16*9%;IF(B16<=1200000;72000+(B16-800000)*30000/400000;"
            f"IF(B16<=1600000;102000+(B16-1200000)*26000/400000;IF(B16<=2500000;128000+(B16-1600000)*59500/900000;B16*7,5%)))));-1))", "", "", ""]),
        (18, "Completion Bond", [f"=N({BOND})", "", "", ""]),
        (19, "Finanzierungskosten", [f"=N({FIN})", "", "", ""]),
        (20, "Überschreitungsreserve", [f"=ROUND(B16*{UER};-1)", "", "", ""]),
        (21, "Herstellungskosten (HK)", ["=SUM(B16:B20)", "=SUM(C16:C20)", "=SUM(D16:D20)", "=SUM(E16:E20)"]),
    ]
    for r, label, vals in rows:
        zs.set(r, 1, label, {"bold": r in (16, 21)})
        for i, v in enumerate(vals):
            if v:
                zs.set(r, 2 + i, v, {**CALC, "bold": r in (16, 21)})
        zs.set(r, 6, f"=SUM(B{r}:E{r})", {**CALC, "bold": r in (16, 21)})
    zs.set(16, 7, "=SUM(G4:G15)", CALC)
    zs.set(16, 8, "=SUM(H4:H15)", CALC)

    zs.set(23, 1, "Berechnete Werte", HEAD)
    zs.set(23, 2, "", HEAD)
    zs.set(23, 3, "Erklärung", HEAD)
    basis = (f"SUMIFS({PO['Plan brutto']};{PO['Partner']};\"AT\";{PO['Art']};\"<>Info\")"
             f"+SUMIFS({PO['LNK']};{PO['Partner']};\"AT\")")
    ph = (f"=IF({BMWKMS}=TRUE();LET(basis;$B$25;ueber;{UER};zusatz;N({BOND})+N({FIN});"
          f"schaetz;ROUND(0,05*(basis*(1,075+ueber)+zusatz)/(1-0,05*(1,075+ueber));-1);"
          f"erste;ROUND(0,05*((basis+schaetz)*1,075+ROUND((basis+schaetz)*ueber;-1)+zusatz);-1);"
          f"zweite;ROUND(0,05*((basis+erste)*1,075+ROUND((basis+erste)*ueber;-1)+zusatz);-1);zweite);"
          f"LET(basis;$B$25;IF(basis<=180000;basis/9;IF(basis<=925000;(6250+basis*0,06875)/0,93125;"
          f"IF(basis<=1400000;(25000+basis*0,05)/0,95;IF(basis<=2375000;(62500+basis*0,025)/0,975;"
          f"IF(basis<=4750000;basis/19;250000)))))))")
    calc = [
        (24, "Produzent*innenhonorar (Pos. 87)", ph, "ÖFI-Staffel auf die NFK inkl. Honorar, ohne Zirkelbezug direkt gelöst (BMWKMS: 5 % der HK)"),
        (25, "NFK AT ohne Produzent*innenhonorar", "=" + basis, "Basis für das Honorar"),
        (26, "Richtsatz Drehbuch", f"=LET(n;$F$16;rs;ROUND(IF(n<=200000;n*12%;IF(n<=500000;24000+(n-200000)*5%;IF(n<=800000;39000+(n-500000)*3%;"
             f"IF(n<=2800000;48000+(n-800000)*2%;IF(n<=6000000;88000+(n-2800000)*1%;n*2%)))));-1);IF({DEBUT_DB}=TRUE();ROUND(rs*{K['DEBUT']};-1);rs))",
         "auf die NFK gesamt (alle Partner); Debut 85 %"),
        (27, "Richtsatz Regie", f"=LET(n;$F$16;rs;ROUND(IF(n<=200000;n*18%;IF(n<=300000;36000+(n-200000)*6%;IF(n<=600000;42000+(n-300000)*4%;"
             f"IF(n<=2400000;54000+(n-600000)*3%;IF(n<=3000000;108000+(n-2400000)*2%;IF(n<=4500000;120000+(n-3000000)*1%;n*3%))))));-1);"
             f"IF({DEBUT_RE}=TRUE();ROUND(rs*{K['DEBUT']};-1);rs))", "auf die NFK gesamt (alle Partner); Debut 85 %"),
        (28, "Abzug Drehbuch (Vorförderung über Freibetrag)",
         f"=MAX(0;SUMIFS({PO['Plan brutto']};{PO['Nr.']};\"1.1a\")-{K['FREIBETRAG']})", "Pos. 1.1a „davon Drehbuchkosten“ abzüglich 15.000 €"),
        (29, "Abzug Regie (Vorförderung über Freibetrag)",
         f"=MAX(0;SUMIFS({PO['Plan brutto']};{PO['Nr.']};\"1.1b\")-{K['FREIBETRAG']})", "Pos. 1.1b „davon Regiekosten“ abzüglich 15.000 €"),
        (30, "Überzahlungen über 10 % (FRL Pkt. 6.2.3)",
         f"=SUMPRODUCT(({PO['Art']}=\"Stab\")*({PO['KV-Mindestgage 40 h']}<>\"\")*"
         f"IFERROR(({PO['Preis']}-{PO['KV-Mindestgage 40 h']}*{K['UEBER']})*{PO['Menge']}*(({PO['Preis']}-{PO['KV-Mindestgage 40 h']}*{K['UEBER']})>0);0))"
         f"+SUMPRODUCT(({PO['Art']}=\"Stab\")*({PO['KV-Mindestgage 40 h']}<>\"\")*"
         f"IFERROR(({PO['Wochengage 60 h']}-{PO['KV-Mindestgage 40 h']}*{K['F_60H']}*{K['UEBER']})*{PO['Wochen 60 h']}*"
         f"(({PO['Wochengage 60 h']}-{PO['KV-Mindestgage 40 h']}*{K['F_60H']}*{K['UEBER']})>0);0))",
         "nur anerkannt, wenn im Finanzierungsplan gedeckt"),
    ]
    for r, label, formula, note in calc:
        zs.set(r, 1, label)
        zs.set(r, 2, formula, CALC)
        zs.set(r, 3, note, NOTE)

    zs.set(32, 1, "Lohnnebenkosten (Pos. 84)", HEAD)
    for i, pt in enumerate(partners):
        zs.set(32, 2 + i, pt, HEAD)
    zs.set(32, 6, "Gesamt", HEAD)
    for i, pt in enumerate(partners):
        cc = col(2 + i)
        stab = f"SUMIFS({PO['Plan brutto']};{PO['Art']};\"Stab\";{PO['Partner']};\"{pt}\")"
        cast = f"SUMIFS({PO['Plan brutto']};{PO['Art']};\"Cast\";{PO['Partner']};\"{pt}\")"
        zs.set(34, 2 + i, f"={stab}*{K['MV']}", CALC)
        zs.set(35, 2 + i, f"=({stab}+{cast})*{K['DB']}", CALC)
        zs.set(36, 2 + i, f"=({stab}+{cast})*{K['DZ']}", CALC)
        zs.set(37, 2 + i, f"=({stab}+{cast})*{K['KOMMST']}", CALC)
        zs.set(38, 2 + i, f"=SUMIFS({PO['Hilfe: U-Bahn-Abgabe']};{PO['Partner']};\"{pt}\")", CALC)
        zs.set(39, 2 + i, f"=SUMIFS({PO['LNK']};{PO['Partner']};\"{pt}\")", {**CALC, "bold": True})
        zs.set(33, 2 + i, f"={cc}39-SUM({cc}34:{cc}38)", CALC)
    for r, label in [(33, "Sozialversicherung (inkl. SZ/UEL)"), (34, "MV-Beitrag"), (35, "DB"), (36, "DZ"),
                     (37, "Kommunalsteuer"), (38, "U-Bahn-Abgabe"), (39, "Summe Lohnnebenkosten")]:
        zs.set(r, 1, label, {"bold": r == 39})
        zs.set(r, 6, f"=SUM(B{r}:E{r})", {**CALC, "bold": r == 39})

    zs.set(41, 1, "Gender-Budgeting (Plan brutto Stab und Cast)", HEAD)
    zs.set(41, 2, "Betrag", HEAD)
    zs.set(41, 3, "Anteil", HEAD)
    for r, label, crit in [(42, "weiblich (W)", "\"W\""), (43, "männlich (M)", "\"M\""), (44, "ohne Angabe", "\"\"")]:
        zs.set(r, 1, label)
        zs.set(r, 2, f"=SUMIFS({PO['Plan brutto']};{PO['Art']};\"Stab\";{PO['M/W']};{crit})"
                     f"+SUMIFS({PO['Plan brutto']};{PO['Art']};\"Cast\";{PO['M/W']};{crit})", CALC)
        zs.set(r, 3, f"=IF(SUM($B$42:$B$44)=0;\"\";B{r}/SUM($B$42:$B$44))", CALC)
    zs.ops += [{"op": "currency", "range": "B4:H21"}, {"op": "currency", "range": "B24:B30"},
               {"op": "currency", "range": "B33:F39"}, {"op": "currency", "range": "B42:B44"},
               {"op": "percent", "range": "C42:C44"}, {"op": "freeze", "rows": 3}]

    # ------------------------------------------------------------ ÖFI-Übertrag
    ue = Sheet("ÖFI-Übertrag")
    head = ["Gr.", "Nr.", "Position", "Plan AT", "Wochen 40 h / Menge AT", "Ø Preis AT", "Wochen 60 h AT",
            "Ø Wochengage 60 h AT", "Ist AT", "noch zu erwarten AT", "Plan A", "Ist A", "noch zu erwarten A",
            "Plan B", "Ist B", "noch zu erwarten B", "Plan C", "Ist C", "noch zu erwarten C", "Plan gesamt"]
    for c, h in enumerate(head, 1):
        ue.set(1, c, h, HEAD)
    ue.set(2, 1, "Summen je ÖFI-Position in der Reihenfolge des ÖFI-Formulars. Werte blockweise ins offizielle Formular übertragen.", NOTE)
    r = 3
    for g in range(1, 13):
        ue.set(r, 1, str(g), GROUP)
        ue.set(r, 3, f"{g}. {gnames[g]}", GROUP)
        for c in range(1, len(head) + 1):
            ue.style(r, c, GROUP)
        r += 1
        for p in by_group.get(g, []):
            nr = p["nr"]
            ue.set(r, 1, str(g))
            ue.set(r, 2, "'" + nr)
            ue.set(r, 3, p["bezeichnung"])
            if p["art"] == "Berechnet" and nr == "87":
                ue.set(r, 4, "='Zusammenfassung'!$B$24", CALC)
                ue.set(r, 20, "='Zusammenfassung'!$B$24", CALC)
            elif p["art"] == "Berechnet" and nr == "84":
                ue.set(r, 4, "='Zusammenfassung'!$B$39", CALC)
                ue.set(r, 11, "='Zusammenfassung'!$C$39", CALC)
                ue.set(r, 14, "='Zusammenfassung'!$D$39", CALC)
                ue.set(r, 17, "='Zusammenfassung'!$E$39", CALC)
                ue.set(r, 20, "='Zusammenfassung'!$F$39", CALC)
                r += 1
                for zr, label in [(33, "Sozialversicherung"), (34, "MV-Beitrag"), (35, "DB"), (36, "DZ"),
                                  (37, "Kommunalsteuer"), (38, "U-Bahn-Abgabe")]:
                    ue.set(r, 1, "3")
                    ue.set(r, 3, f"   davon {label}", NOTE)
                    for cc, zc in [(4, "B"), (11, "C"), (14, "D"), (17, "E"), (20, "F")]:
                        ue.set(r, cc, f"='Zusammenfassung'!${zc}${zr}", CALC)
                    r += 1
                continue
            elif p["art"] != "Info":
                crit = f"{PO['Nr.']};$B{r}"

                def s(name, pt):
                    return f"=SUMIFS({PO[name]};{crit};{PO['Partner']};\"{pt}\")"
                if p["art"] == "Überschrift":
                    sub = f"=SUMIFS({PO['Plan brutto']};{PO['Nr.']};$B{r}&\".*\";{PO['Partner']};\"AT\")"
                    ue.set(r, 4, sub, CALC)
                else:
                    ue.set(r, 4, s("Plan brutto", "AT"), CALC)
                    ue.set(r, 5, s("Menge", "AT"), CALC)
                    ue.set(r, 6, f"=IF(N(E{r})=0;\"\";SUMIFS({PO['Hilfe: Menge × Preis']};{crit};{PO['Partner']};\"AT\")/E{r})", CALC)
                    ue.set(r, 7, s("Wochen 60 h", "AT"), CALC)
                    ue.set(r, 8, f"=IF(N(G{r})=0;\"\";SUMIFS({PO['Hilfe: 60 h']};{crit};{PO['Partner']};\"AT\")/G{r})", CALC)
                    ue.set(r, 9, s("Ist", "AT"), CALC)
                    ue.set(r, 10, s("noch zu erwarten", "AT"), CALC)
                    for base, pt in [(11, "A"), (14, "B"), (17, "C")]:
                        ue.set(r, base, s("Plan brutto", pt), CALC)
                        ue.set(r, base + 1, s("Ist", pt), CALC)
                        ue.set(r, base + 2, s("noch zu erwarten", pt), CALC)
                    ue.set(r, 20, f"=SUM(D{r};K{r};N{r};Q{r})", CALC)
            r += 1
    ue_last = r - 1
    r += 1
    ue.set(r, 1, "Für das ÖFI-Blatt „Zusammenfassung“", HEAD)
    for c in range(2, 5):
        ue.set(r, c, "", HEAD)
    r += 1
    for zr, label in [(16, "Fertigungskosten (NFK)"), (17, "Fertigungsgemeinkosten"), (18, "Completion Bond"),
                      (19, "Finanzierungskosten"), (20, "Überschreitungsreserve"), (21, "Herstellungskosten")]:
        ue.set(r, 3, label, {"bold": zr in (16, 21)})
        ue.set(r, 4, f"='Zusammenfassung'!$B${zr}", CALC)
        ue.set(r, 20, f"='Zusammenfassung'!$F${zr}", CALC)
        r += 1
    ue.ops += [{"op": "text", "range": f"B3:B{ue_last}"},
               {"op": "currency", "range": f"D3:D{r}"}, {"op": "currency", "range": f"F3:F{ue_last}"},
               {"op": "currency", "range": f"H3:T{r}"}, {"op": "freeze", "rows": 1, "cols": 3}]

    return [pr, sa, kv, po, zs, ue]


if __name__ == "__main__":
    sheets = build(sys.argv[1], sys.argv[2])
    json.dump([s.dump() for s in sheets], sys.stdout, ensure_ascii=False)
