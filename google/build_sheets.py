"""Erzeugt Inhalt, Formeln und Formatvorgaben für die interne Kalkulation als Google-Tabelle.

Ausgabe (JSON): je Blatt die Zellen (Text, wie in Google Tabellen mit deutschem Gebietsschema
eingetippt: Formeln mit ";" und Dezimalkomma), Zellstile für den HTML-Einfügeweg und eine
Liste von Bedienschritten (Checkboxen, Dropdowns, Zahlenformate, Fixierung, ausgeblendete
Spalten, bedingte Formatierung), die google/fill.js im Browser ausführt.

Die Regeln der Förderstelle (FISA+ oder ÖFI) stehen im Blatt „Sätze“ (Spalte „Förderstelle“)
und werden über das Dropdown im Blatt „Projekt“ gewählt.

Aufruf:
    python3 -I google/build_sheets.py vorlage/20270826_EMPTY_TEMPLATE.xlsx google/positions.json \
        [google/eingaben.json] > google/sheets.json

eingaben.json (optional) enthält Werte, die beim Neuaufbau erhalten bleiben sollen:
    {"Projekt": {"Feldname": "Wert"}, "Positionen": {"Nr|Zeilenindex": {"Spalte": "Wert"}}}
"""

import json
import sys
import warnings

import openpyxl

warnings.filterwarnings("ignore")

BUNDESLAENDER = [("Burgenland", "0,40%"), ("Kärnten", "0,37%"), ("Niederösterreich", "0,33%"),
                 ("Oberösterreich", "0,31%"), ("Salzburg", "0,35%"), ("Steiermark", "0,34%"),
                 ("Tirol", "0,39%"), ("Vorarlberg", "0,33%"), ("Wien", "0,36%")]
FOERDERSTELLEN = ["FISA+", "ÖFI"]

HEAD = {"bold": True, "bg": "#f3f3f3"}
GROUP = {"bold": True, "bg": "#e8f0fe"}
CALC = {"bg": "#f8f9fa"}
TITLE = {"bold": True, "size": 14}
NOTE = {"color": "#5f6368", "italic": True}
INPUT = {"bg": "#fff8e1"}
DEPT = {"bold": True, "bg": "#f1f3f4"}
N_RECHNUNGSLEGER = 10   # freie Zeilen 94.5 ff.
N_DARSTELLER = 15       # vorbereitete Zusatzzeilen unter Pos. 82

# Departments in Gruppe 3 (ÖFI-Reihenfolge): Name, erste Nr., erste Nr. des nächsten Departments
DEPARTMENTS = [("Produktion", 13, 24), ("Regie", 24, 28), ("Kamera", 28, 38), ("Ton", 38, 43),
               ("Schnitt & Post", 43, 47), ("Kostüm", 47, 53), ("Requisite", 53, 57), ("Szenenbild", 57, 59),
               ("Maske", 59, 65), ("Fahrer", 65, 69), ("Licht", 69, 75), ("Bühne", 75, 79),
               ("Musik & Sonstige", 79, 82), ("Cast", 82, 84), ("Sonstige Kosten", 84, 89),
               ("Honorare & Dienstleister", 89, 95)]

# Blatt „Reisen“: Catering-Zeilen und erste Personenzeile; Blatt „Motive“: erste Motivzeile
REI_CAT = (5, 7)
REI_FIRST, N_REISEN = 12, 40
MOT_FIRST, N_MOTIVE = 7, 30
# Spalten im Blatt Motive -> Position
MOTIV_COLS = [("E", "Motivgebühr", "116"), ("F", "Abgaben & Halteverbote", "116"), ("G", "Nebenräume", "116"),
              ("H", "Absperrhilfen", "117"), ("I", "Strom", "118"), ("J", "Baumaterial Kauf", "114"),
              ("K", "Baumaterial Miete", "115"), ("L", "Requisiten Kauf", "122"), ("M", "Requisiten Miete", "123"),
              ("N", "Spielfahrzeuge", "128"), ("O", "Stunts/SFX", "129")]


def feed(nr, partner):
    """Zufluss aus den Blättern Reisen/Motive für eine Position (je Partner der Zeile) oder None."""
    rs = lambda c: f"SUMIFS('Reisen'!${c}${REI_FIRST}:${c};'Reisen'!$D${REI_FIRST}:$D;{partner})"
    if nr == "160":
        return "Reisen", rs("P")
    if nr == "161":
        a, b = REI_CAT
        return "Reisen", rs("Q") + f"+SUMIFS('Reisen'!$G${a}:$G${b};'Reisen'!$B${a}:$B${b};{partner})"
    cols = [c for c, _, n in MOTIV_COLS if n == nr]
    if cols:
        return "Motive", "+".join(f"SUMIFS('Motive'!${c}${MOT_FIRST}:${c};'Motive'!$C${MOT_FIRST}:$C;{partner})" for c in cols)
    return None


# Zeilen der Zusammenfassung, auf die andere Blätter verweisen
ZR = {"ph_oefi": 42, "basis": 43, "ph_max": 44, "rs_db": 45, "rs_re": 46, "abz_db": 47, "abz_re": 48,
      "ueber": 49, "hl": 50, "lnk_head": 52, "lnk_first": 53, "lnk_sum": 59, "gender_head": 61, "dep_head": 66}


def department(nr):
    try:
        x = float(str(nr).lstrip("'"))
    except ValueError:
        return None
    return next((name for name, a, b in DEPARTMENTS if a <= x < b), None)


class Sheet:
    def __init__(self, name):
        self.name = name
        self.cells = {}
        self.styles = {}
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


def build(template, positions_path, inputs=None):
    inputs = inputs or {}
    wv = openpyxl.load_workbook(template, data_only=True)
    pos = json.load(open(positions_path))
    by_group = {}
    for p in pos["positionen"]:
        by_group.setdefault(p["gruppe"], []).append(p)
        if p["nr"] == "94.4":
            # freie Zeilen wie in der ÖFI-Vorlage (dort ohne Bezeichnung, daher beim Einlesen übersprungen)
            for k in range(5, 5 + N_RECHNUNGSLEGER):
                by_group[p["gruppe"]].append({"row": None, "gruppe": p["gruppe"], "nr": f"94.{k}", "bezeichnung": "Rechnungsleger",
                                              "einheit": "Tagsatz", "art": "Kosten", "kv_spielfilm": "", "kv_doku": "", "richtsatz": None})
    gnames = {g["nr"]: g["name"] for g in pos["gruppen"]}
    UE = "Übertrag Antrag"

    # ================================================================ Projekt
    pr = Sheet("Projekt")
    pr.set(1, 1, "Projekt", TITLE)
    fields = [
        ("Projekttitel", "", ""),
        ("Produktionsfirma", "", ""),
        ("Förderstelle (Regelwerk)", "FISA+", "bestimmt Honorar, FGK, Prüfungen und Anreizförderung (Werte im Blatt „Sätze“)"),
        ("Format", "Film", "Film oder Serie (FISA+: Höchstbetrag und Mindest-HK)"),
        ("Anzahl Folgen", "1", "bei Serie: Anzahl der Folgen der Staffel"),
        ("Filmart", "Spielfilm", "Spielfilm oder Dokumentarfilm"),
        ("Werkstattprojekt (§ 19 KV)", "nein", "ÖFI: Werkstatt-KV bis HK 1,22 Mio. €, halber KV bis 1,57 Mio. € (Prüfung in der Zusammenfassung)"),
        ("gefördert durch BMWKMS/if", "", "ÖFI: FGK 7,5 %, Honorar 5 % der HK; Voraussetzung für ÖFI+"),
        ("Debutfilm (Regie)", "", "ÖFI: Regie-Richtsatz 85 %"),
        ("Debut (Drehbuch)", "", "ÖFI: Drehbuch-Richtsatz 85 %"),
        ("Green Bonus beantragt", "", "Anreizförderung +5 Prozentpunkte"),
        ("Gender Gap Financing beantragt", "", "Anreizförderung +25.000 €"),
        ("Bundesland der Betriebsstätte", "Wien", "bestimmt DZ-Satz; U-Bahn-Abgabe nur bei Wien"),
        ("Stichtag für Sätze", "01.01.2027", "z. B. Drehbeginn: gilt für KV-Gagen und alle Sätze"),
        ("Überschreitungsreserve", "5%", "FISA+: bis 8 %; ÖFI: 5 % Pflicht, bis 8 % (10 % mit Bond)"),
        ("Honorar Produzent*in (Anteil der FK)", "5%",
         "auf die Fertigungskosten ohne Honorar; 0 % oder leer = kein Honorar; Höchstwert FISA+ 5 %, ÖFI Staffel (Prüfung in der Zusammenfassung)"),
        ("Completion Bond", "", ""),
        ("Finanzierungskosten", "", ""),
        ("Kostenstand per", "", ""),
    ]
    P = {}
    for i, (label, value, hint) in enumerate(fields):
        r = 3 + i
        value = inputs.get("Projekt", {}).get(label, value)
        pr.set(r, 1, label, {"bold": True})
        pr.set(r, 2, value, INPUT)
        pr.set(r, 3, hint, NOTE)
        P[label] = f"'Projekt'!$B${r}"
    FST, FORMAT, FOLGEN, FILMART = P["Förderstelle (Regelwerk)"], P["Format"], P["Anzahl Folgen"], P["Filmart"]
    WERKSTATT, BMWKMS = P["Werkstattprojekt (§ 19 KV)"], P["gefördert durch BMWKMS/if"]
    DEBUT_RE, DEBUT_DB = P["Debutfilm (Regie)"], P["Debut (Drehbuch)"]
    GREEN, GGF = P["Green Bonus beantragt"], P["Gender Gap Financing beantragt"]
    LAND, STICHTAG, UER = P["Bundesland der Betriebsstätte"], P["Stichtag für Sätze"], P["Überschreitungsreserve"]
    BOND, FIN = P["Completion Bond"], P["Finanzierungskosten"]
    PHP = f"N({P['Honorar Produzent*in (Anteil der FK)']})"
    ISFISA = f"{FST}=\"FISA+\""
    ISOEFI = f"{FST}=\"ÖFI\""
    rr = {label: 3 + i for i, (label, *_ ) in enumerate(fields)}
    pr.ops += [
        {"op": "checkbox", "range": f"B{rr['gefördert durch BMWKMS/if']}:B{rr['Gender Gap Financing beantragt']}"},
        {"op": "dropdown", "range": f"B{rr['Werkstattprojekt (§ 19 KV)']}", "options": ["nein", "Werkstatt-KV", "halber KV"]},
        {"op": "dropdown", "range": f"B{rr['Förderstelle (Regelwerk)']}", "options": FOERDERSTELLEN},
        {"op": "dropdown", "range": f"B{rr['Format']}", "options": ["Film", "Serie"]},
        {"op": "dropdown", "range": f"B{rr['Filmart']}", "options": ["Spielfilm", "Dokumentarfilm"]},
        {"op": "dropdown", "range": f"B{rr['Bundesland der Betriebsstätte']}", "options": [b for b, _ in BUNDESLAENDER]},
        {"op": "percent", "range": f"B{rr['Überschreitungsreserve']}:B{rr['Honorar Produzent*in (Anteil der FK)']}"},
        {"op": "currency", "range": f"B{rr['Completion Bond']}:B{rr['Finanzierungskosten']}"},
    ]

    # Sätze: (Schlüssel, Bezeichnung, Förderstelle, Bundesland, Wert, Einheit, gültig ab, Quelle)
    FI, OE = "FISA+", "ÖFI"
    RL_FISA = "FISA+-Richtlinie österr. Produktionen (aws, ab 01.09.2025)"
    S = [
        ("SV_HBGL", "SV-Höchstbeitragsgrundlage laufend (monatlich)", "", "", "6930", "€", "01.01.2026", "ÖFI-Vorlage 2027 (vorläufig)"),
        ("SV_LFD", "DG-Anteil SV laufend (AV 2,95 + KV 3,78 + UV 1,10 + PV 12,55 + IESG 0,10 + WF 0,50)", "", "", "20,98%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("SV_SZ", "DG-Anteil SV Sonderzahlungen (ohne WF)", "", "", "20,48%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("UEL", "Urlaubsersatzleistung", "", "", "10,41%", "", "01.01.2026", "ÖFI-Rechenlogik"),
        ("DB", "Dienstgeberbeitrag (FLAF)", "", "", "3,7%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
    ] + [("DZ", "Zuschlag zum DB", "", land, satz, "", "01.01.2026", "WKO, Hebesätze 2026") for land, satz in BUNDESLAENDER] + [
        ("KOMMST", "Kommunalsteuer", "", "", "3%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("MV", "Betriebliche Vorsorge (MV-Beitrag)", "", "", "1,53%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("UBAHN", "Dienstgeberabgabe Wien (U-Bahn-Abgabe) je Woche", "", "Wien", "2", "€", "01.06.2012", "Stadt Wien"),
        ("F_SZUEL", "Faktor SZ/UEL auf KV-Grundgage", "", "", "1,2881166", "", "01.01.2026", "ÖFI-Vorlage (Kollektivvertrag)"),
        ("F_60H", "Faktor Wochenpauschale 60 h (§ 7 KV)", "", "", "1,385", "", "01.01.2026", "ÖFI-Vorlage (Kollektivvertrag)"),
        ("WS1", "Werkstattprojekt: HK-Grenze Werkstatt-KV", "", "", "1220000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("WS2", "Werkstattprojekt: HK-Grenze halber KV", "", "", "1570000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("DIAET", "Tagesgeld Inland", "", "", "30", "€", "01.01.2025", "§ 26 Z 4 EStG"),
        ("NAECHT", "Nächtigungsgeld Inland (Pauschale)", "", "", "17", "€", "01.01.2025", "§ 26 Z 4 EStG"),
        ("KM", "Kilometergeld", "", "", "0,5", "€", "01.01.2025", "§ 26 Z 4 EStG"),
        # ---- Förderstelle ÖFI
        ("ILV", "ILV: max. Anteil der KV-Gage (KV × 1,1 × 0,8)", OE, "", "88%", "", "01.01.2026", "ÖFI-Richtlinie"),
        ("UEBER", "Überzahlung anerkannt bis (FRL Pkt. 6.2.3)", OE, "", "110%", "", "01.01.2026", "ÖFI-Richtlinie"),
        ("GFC_W", "Green Film Consultant: Höchstsatz je Woche", OE, "", "3000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("GFC_ILV", "Green Film Consultant: Höchstsatz je Woche bei ILV", OE, "", "2400", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("MD_MAX", "Producer for Marketing & Distribution: max. anerkannt", OE, "", "10000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("FREIBETRAG", "Drehbuch/Regie: Freibetrag Vorförderungen", OE, "", "15000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("DEBUT", "Richtsatz bei Debut", OE, "", "85%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("ILV_RS", "Richtsatz Drehbuch/Regie bei ILV", OE, "", "80%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("UER_MIN", "Überschreitungsreserve mindestens", OE, "", "5%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("UER_MAX", "Überschreitungsreserve höchstens (mit Completion Bond 10 %)", OE, "", "8%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("ANREIZ", "Anreizförderung (ÖFI+): Satz", OE, "", "30%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("GREEN", "Anreizförderung: Green Bonus", OE, "", "5%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("GGF", "Anreizförderung: Gender Gap Financing", OE, "", "25000", "€", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("BEM_MAX", "Bemessungsgrundlage höchstens (Anteil der HK gesamt)", OE, "", "80%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        ("EIGEN", "Eigenanteil (Barmittel) Anteil der HK", OE, "", "2,5%", "", "01.01.2026", "ÖFI-Vorlage 2027"),
        # ---- Förderstelle FISA+
        ("PH_MAX", "Honorar Produzent*in: max. Anteil der Fertigungskosten", FI, "", "5%", "", "01.09.2025", RL_FISA + " Pkt. 4.1 (12)"),
        ("FGK_MAX", "Fertigungsgemeinkosten: max. Anteil der Fertigungskosten", FI, "", "7,5%", "", "01.09.2025", RL_FISA + " Pkt. 4.1 (13)"),
        ("HL_MAX", "Herstellungsleitung: max. Anteil der Fertigungskosten", FI, "", "2,5%", "", "01.09.2025", RL_FISA + " Pkt. 4.1 (6)"),
        ("UER_MIN", "Überschreitungsreserve mindestens", FI, "", "0%", "", "01.09.2025", RL_FISA + " Pkt. 4.1 (15)"),
        ("UER_MAX", "Überschreitungsreserve höchstens", FI, "", "8%", "", "01.09.2025", RL_FISA + " Pkt. 4.1 (15)"),
        ("ANREIZ", "Anreizförderung (FISA+): Satz", FI, "", "30%", "", "01.09.2025", RL_FISA + " Pkt. 5 (4)"),
        ("GREEN", "Anreizförderung: Green Bonus", FI, "", "5%", "", "01.09.2025", RL_FISA + " Pkt. 5 (4)"),
        ("GGF", "Anreizförderung: Gender Gap Financing", FI, "", "25000", "€", "01.09.2025", RL_FISA + " Pkt. 5 (6)"),
        ("BEM_MAX", "Bemessungsgrundlage höchstens (Anteil der HK gesamt)", FI, "", "80%", "", "01.09.2025", RL_FISA + " Pkt. 5 (3)"),
        ("MAX_FILM", "Zuschuss höchstens je Film bzw. je Folge", FI, "", "5000000", "€", "01.09.2025", RL_FISA + " Pkt. 5 (7)"),
        ("MAX_SERIE", "Zuschuss höchstens je Serie bzw. Staffel", FI, "", "7500000", "€", "01.09.2025", RL_FISA + " Pkt. 5 (7)"),
        ("MIN_F_SPIEL", "förderbare Kosten mindestens (Spielfilm/fiktional)", FI, "", "150000", "€", "01.09.2025", RL_FISA + " Pkt. 4.3"),
        ("MIN_F_DOKU", "förderbare Kosten mindestens (Dokumentarfilm)", FI, "", "80000", "€", "01.09.2025", RL_FISA + " Pkt. 4.3"),
        ("MIN_HK_FILM", "Gesamt-HK mindestens je Film", FI, "", "1800000", "€", "01.09.2025", RL_FISA + " Pkt. 7.3"),
        ("MIN_HK_FOLGE", "Gesamt-HK mindestens je Folge", FI, "", "600000", "€", "01.09.2025", RL_FISA + " Pkt. 7.3"),
        ("MDA_MIN", "Mediendienstanbieter: Beteiligung mindestens (Anteil HK)", FI, "", "30%", "", "01.09.2025", RL_FISA + " Pkt. 7.6"),
        ("EIGEN", "Eigenanteil in der Regel (Anteil HK)", FI, "", "2,5%", "", "01.09.2025", RL_FISA + " Pkt. 7.5"),
        # ---- ergänzt (am Ende, damit bestehende Zeilen gleich bleiben)
        ("DIAET_W", "Diäten Ortsgebiet (lt. KV Filmschaffende)", "", "", "15,4", "€", "01.01.2025", "ÖFI-Vorlage 2027 (lt. KV 2025)"),
    ]
    keys = []
    for k, *_ in S:
        if k not in keys:
            keys.append(k)
    KEYROW = len(fields) + 5
    pr.set(KEYROW - 1, 1, "Aktuelle Sätze zum Stichtag und zur gewählten Förderstelle (automatisch; Werte im Blatt „Sätze“ pflegen)", {"bold": True})
    K = {}
    for i, k in enumerate(keys):
        r = KEYROW + i
        label = next(s[1] for s in S if s[0] == k)
        pr.set(r, 1, label)
        pr.set(r, 2, f"=SUMIFS('Sätze'!$E:$E;'Sätze'!$A:$A;$C{r};'Sätze'!$I:$I;TRUE())", CALC)
        pr.set(r, 3, k, NOTE)
        K[k] = f"'Projekt'!$B${r}"
    last_key = KEYROW + len(keys) - 1
    pr.ops += [{"op": "currency", "range": f"B{KEYROW}:B{last_key}"}]
    pct_keys = {s[0] for s in S if s[4].endswith("%")}
    pr.ops += [{"op": "percent", "range": f"B{KEYROW + keys.index(k)}"} for k in keys if k in pct_keys]
    pr.ops += [{"op": "plain", "range": f"B{KEYROW + keys.index(k)}"} for k in ("F_SZUEL", "F_60H")]

    # ================================================================ Sätze
    sa = Sheet("Sätze")
    head = ["Schlüssel", "Bezeichnung", "Förderstelle", "Bundesland", "Wert", "Einheit", "gültig ab", "Quelle",
            "gilt am Stichtag", "Hilfe: Gruppe"]
    for c, h in enumerate(head, 1):
        sa.set(1, c, h, HEAD)
    for i, row in enumerate(S):
        r = 2 + i
        for c, v in enumerate(row, 1):
            sa.set(r, c, v)
        sa.set(r, 10, f"=$A{r}&\"|\"&$C{r}&\"|\"&$D{r}", CALC)
        sa.set(r, 9, f"=AND($G{r}<={STICHTAG};OR($C{r}=\"\";$C{r}={FST});OR($D{r}=\"\";$D{r}={LAND});"
                     f"$G{r}=MAXIFS($G:$G;$J:$J;$J{r};$G:$G;\"<=\"&{STICHTAG}))", CALC)
    sa.ops += [{"op": "dropdown", "range": f"C2:C{1 + len(S)}", "options": FOERDERSTELLEN},
               {"op": "freeze", "rows": 1}, {"op": "hide", "range": "J:J"}]

    # ================================================================ KV
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

    # ================================================================ Positionen
    po = Sheet("Positionen")
    cols = ["Gr.", "Nr.", "Position", "Name / Beschreibung", "Partner", "KV-Funktion", "Menge", "Einheit", "KV-Gage",
            "Preis", "ILV", "Rechnung", "M/W", "Plan brutto", "LNK", "Plan gesamt", "Anteil förderbar", "förderbar", "Prüfung",
            "Ist", "noch zu erwarten", "Endkosten", "± zu Plan", "Abw. zu KV", "ohne U-Bahn-Abgabe",
            "Art", "Department", "Hilfe: U-Bahn-Abgabe", "Hilfe: Gage", "Hilfe: Wochen", "Hilfe: 60 h"]
    C = {name: i + 1 for i, name in enumerate(cols)}
    L = {name: col(i) for name, i in C.items()}
    for c, h in enumerate(cols, 1):
        po.set(1, c, h, HEAD)
    po.set(2, 3, "Stab: KV-Funktion und Einheit (Wochen 40 h, Wochen 60 h, Tage) wählen, Menge eintragen; Preis leer = KV-Gage, sonst eigene Gage; "
                 "„Rechnung“ = Rechnungssteller ohne Lohnnebenkosten (Preis netto Pflicht) · Cast: Tagesgage · sonst Einzelpreis (ohne Menge = Pauschale). "
                 "Weitere Person/Partner: Zeile darunter einfügen, die Zeile der Position hineinkopieren und Werte überschreiben.", NOTE)
    r = 3
    first_data = r
    group_rows, dept_rows, rows_meta = [], [], {}
    for g in range(1, 13):
        po.set(r, 1, str(g), GROUP)
        po.set(r, 3, f"{g}. {gnames[g]}", GROUP)
        po.set(r, C["Art"], "Gruppe", GROUP)
        for c in range(1, len(cols) + 1):
            po.style(r, c, GROUP)
        group_rows.append(r)
        r += 1
        cur_dept = None
        plist = []
        for p in by_group.get(g, []):
            plist.append(p)
            if p["nr"] == "82":
                plist += [p] * N_DARSTELLER   # vorbereitete Darsteller-Zeilen (gleiche Nr.)
        for p in plist:
            dep = department(p["nr"]) if g == 3 else None
            if dep and dep != cur_dept:
                # Department-Zwischenzeile mit Zwischensummen
                po.set(r, 3, dep, DEPT)
                po.set(r, C["Art"], "Department", DEPT)
                po.set(r, C["Department"], dep, DEPT)
                for c in range(1, len(cols) + 1):
                    po.style(r, c, DEPT)
                dept_rows.append(r)
                cur_dept = dep
                r += 1
            po.set(r, 2, "'" + p["nr"])
            po.set(r, 3, p["bezeichnung"])
            rows_meta[r] = p
            art = p["art"]
            if art in ("Kosten", "Stab", "Cast", "Info"):
                po.set(r, C["Partner"], "AT", INPUT)
            if art == "Stab":
                po.set(r, C["Einheit"], "Wochen 40 h", INPUT)
            elif art == "Cast":
                po.set(r, C["Einheit"], "Drehtage")
            elif art == "Kosten":
                po.set(r, C["Einheit"], p["einheit"])
            elif art == "Berechnet" and p["nr"] == "84":
                po.set(r, 4, "je Person in Spalte LNK enthalten; Aufteilung siehe Zusammenfassung", NOTE)
            elif art == "Berechnet":
                po.set(r, 4, f"=IF({ISFISA};\"FISA+: Honorar steht als Pkt. 14 in der Zusammenfassung\";"
                             f"\"wird automatisch berechnet (siehe Zusammenfassung)\")", NOTE)
            elif art == "Überschrift":
                po.set(r, 4, "Summe der Unterpositionen", NOTE)
            r += 1
    last = r - 1
    rng = lambda name: f"{L[name]}{first_data}:{L[name]}{last}"
    UEA = f"'{UE}'!$A:$A"
    UEB = f"'{UE}'!$B:$B"
    UE_ART = f"'{UE}'!$U:$U"
    UE_KV = f"'{UE}'!$V:$V"
    UE_DEP = f"'{UE}'!$W:$W"
    boundaries = sorted(group_rows + dept_rows)
    SUMCOLS = ("Plan brutto", "LNK", "Plan gesamt", "förderbar", "Ist", "noch zu erwarten", "Endkosten", "± zu Plan")
    NOSUM = "\"<>Gruppe\";{a};\"<>Department\";{a};\"<>Info\""

    for rr_ in range(first_data, last + 1):
        f = lambda name: f"{L[name]}{rr_}"
        art = f("Art")
        if rr_ in boundaries:
            # Gruppe: Block bis zur nächsten Gruppenzeile; Department: bis zur nächsten Gruppen- oder Department-Zeile
            stops = group_rows if rr_ in group_rows else boundaries
            nxt = next((b for b in stops if b > rr_), last + 1)
            style = GROUP if rr_ in group_rows else DEPT
            blk = lambda name: f"{L[name]}{rr_ + 1}:{L[name]}{nxt}"
            for name in SUMCOLS:
                po.set(rr_, C[name], f"=SUMIFS({blk(name)};{blk('Art')};" + NOSUM.format(a=blk('Art')) + ")", style)
            continue
        # Gr., Art, KV-Funktion und Department automatisch aus der Nr.
        frei94 = f"LEFT($B{rr_};3)=\"94.\""
        po.set(rr_, C["Gr."], f"=IFERROR(INDEX({UEA};MATCH($B{rr_};{UEB};0));IF({frei94};3;\"\"))", CALC)
        po.set(rr_, C["Art"], f"=IF($B{rr_}=\"\";\"\";IFERROR(INDEX({UE_ART};MATCH($B{rr_};{UEB};0));\"Kosten\"))", CALC)
        po.set(rr_, C["Department"], f"=IF($B{rr_}=\"\";\"\";IFERROR(INDEX({UE_DEP};MATCH($B{rr_};{UEB};0))&\"\";IF({frei94};\"Honorare & Dienstleister\";\"\")))", CALC)
        po.set(rr_, C["KV-Funktion"], f"=IFERROR(INDEX({UE_KV};MATCH($B{rr_};{UEB};0))&\"\";\"\")", CALC)
        einh = f"LOWER(SUBSTITUTE({f('Einheit')};\" \";\"\"))"
        po.set(rr_, C["KV-Gage"],
               f"=IF(OR({art}<>\"Stab\";{f('KV-Funktion')}=\"\");\"\";"
               f"LET(kv;SUMIFS('KV'!$E:$E;'KV'!$A:$A;{f('KV-Funktion')};'KV'!$H:$H;TRUE());"
               f"ws;SUMIFS('KV'!$E:$E;'KV'!$A:$A;\"Werkstattprojekt (§ 19 KV)\";'KV'!$H:$H;TRUE());"
               f"grund;IF({ISOEFI};IF({WERKSTATT}=\"Werkstatt-KV\";ws;IF({WERKSTATT}=\"halber KV\";kv/2;kv));kv);"
               f"einh;{einh};"
               f"faktor;IF(einh=\"wochen40h\";1;IF(einh=\"wochen60h\";{K['F_60H']};IF(einh=\"tage\";0,2;0)));"
               f"IF(OR(grund=0;faktor=0);\"\";grund*faktor)))", CALC)
        rech = f"{f('Rechnung')}=TRUE()"
        po.set(rr_, C["Hilfe: Gage"], f"=IF(OR({rech};{f('Preis')}<>\"\");N({f('Preis')});N({f('KV-Gage')}))", CALC)
        po.set(rr_, C["Hilfe: Wochen"], f"=IF(AND({art}=\"Stab\";{einh}=\"tage\");N({f('Menge')})/5;N({f('Menge')}))", CALC)
        po.set(rr_, C["Hilfe: 60 h"], f"={einh}=\"wochen60h\"", CALC)
        fd = feed(rows_meta[rr_]["nr"], f("Partner"))
        if fd:
            # Summe aus dem Blatt Reisen bzw. Motive plus eigene Eingabe (Menge × Preis)
            blatt, expr = fd
            po.set(rr_, C["Plan brutto"],
                   f"=LET(zu;{expr};eigen;IF({f('Preis')}=\"\";0;IF({f('Menge')}=\"\";N({f('Preis')});N({f('Menge')})*N({f('Preis')})));"
                   f"IF(zu+eigen=0;\"\";zu+eigen))", CALC)
            po.set(rr_, C["Name / Beschreibung"], f"=LET(zu;{expr};IF(zu=0;\"\";\"aus Blatt {blatt}: \"&TEXT(zu;\"#,##0 €\")))", NOTE)
        else:
            po.set(rr_, C["Plan brutto"],
                   f"=IF(OR({art}=\"Überschrift\";{art}=\"Berechnet\");\"\";"
                   f"IF(AND({f('Preis')}=\"\";OR({rech};{f('KV-Gage')}=\"\";N({f('Menge')})=0));\"\";"
                   f"IF({f('Menge')}=\"\";N({f('Preis')});N({f('Menge')})*{f('Hilfe: Gage')})))", CALC)
        ub = f"IF({f('ohne U-Bahn-Abgabe')}=TRUE();0;{K['UBAHN']})"
        po.set(rr_, C["Hilfe: U-Bahn-Abgabe"],
               f"=IF({rech};0;IF({art}=\"Stab\";N({f('Hilfe: Wochen')})*{ub};IF({art}=\"Cast\";N({f('Menge')})/5*{ub};0)))", CALC)
        stab = (f"LET(brutto;N({f('Plan brutto')});wochen;N({f('Hilfe: Wochen')});hbgl;{K['SV_HBGL']};"
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
        po.set(rr_, C["LNK"], f"=IF({rech};\"\";IF({art}=\"Stab\";{stab};IF({art}=\"Cast\";{cast};\"\")))", CALC)
        po.set(rr_, C["Plan gesamt"],
               f"=IF({art}=\"Berechnet\";IF(AND({f('Nr.')}=\"87\";{ISOEFI});'Zusammenfassung'!$B${ZR['ph_oefi']};\"\");"
               f"IF(AND({f('Plan brutto')}=\"\";{f('LNK')}=\"\");\"\";N({f('Plan brutto')})+N({f('LNK')})))", CALC)
        po.set(rr_, C["Anteil förderbar"], f"=IF({f('Partner')}=\"AT\";1;0)", INPUT)
        po.set(rr_, C["förderbar"], f"=IF({f('Plan gesamt')}=\"\";\"\";N({f('Plan gesamt')})*N({f('Anteil förderbar')}))", CALC)
        kvg = f"N({f('KV-Gage')})"
        gage = f"N({f('Hilfe: Gage')})"
        pr_ = f"N({f('Preis')})"
        po.set(rr_, C["Abw. zu KV"], f"=IF(OR({rech};{kvg}=0;N({f('Plan brutto')})=0);\"\";{gage}/{kvg}-1)", CALC)
        ilv = f"{f('ILV')}=TRUE()"
        checks = [
            # beide Förderstellen
            f"IF(AND({art}=\"Stab\";NOT({ilv});NOT({rech});{f('Preis')}<>\"\";{kvg}>0;{pr_}<{kvg});\"unter KV\";\"\")",
            f"IF(AND({art}=\"Stab\";NOT({rech});{f('Preis')}=\"\";N({f('Menge')})>0;{f('KV-Gage')}=\"\");\"keine KV-Gage: KV-Funktion/Einheit prüfen oder Preis eintragen\";\"\")",
            f"IF(AND({rech};{f('Preis')}=\"\";N({f('Menge')})>0);\"Rechnung: Preis (netto) eintragen\";\"\")",
            f"IF(AND({f('Gr.')}=12;N({f('Plan brutto')})>0);\"Erträge mit negativem Vorzeichen eintragen\";\"\")",
            f"IF(N({f('Anteil förderbar')})>1;\"Anteil förderbar über 100 %\";\"\")",
            f"IF(AND({f('Nr.')}<>\"\";COUNTIF({UEB};{f('Nr.')})=0;NOT({frei94}));\"Nr. nicht in der Positionsliste\";\"\")",
            f"IF(AND({f('Partner')}<>\"\";N({f('Plan brutto')})<>0;COUNTIF({{\"AT\";\"A\";\"B\";\"C\"}};{f('Partner')})=0);\"Partner AT, A, B oder C\";\"\")",
            # nur ÖFI
            f"IF(AND({ISOEFI};{art}=\"Stab\";NOT({rech});{ilv};{kvg}>0;N({f('Plan brutto')})>0;{gage}>{kvg}*{K['ILV']});\"ILV über 88 % KV\";\"\")",
            f"IF(AND({ISOEFI};{art}=\"Stab\";NOT({rech});{kvg}>0;{pr_}>{kvg}*{K['UEBER']});\"Überzahlung über 10 % (nur mit Begründung)\";\"\")",
            f"IF(AND({ISOEFI};{f('Nr.')}=\"81.0\";{pr_}>IF({ilv};{K['GFC_ILV']};{K['GFC_W']}));\"über Höchstsatz Green Film Consultant\";\"\")",
            f"IF(AND({ISOEFI};{f('Nr.')}=\"94.4\";N({f('Plan brutto')})>{K['MD_MAX']});\"max. 10.000 € anerkannt\";\"\")",
            f"IF(AND({ISOEFI};{f('Nr.')}=\"4\";N({f('Plan brutto')})>ROUND('Zusammenfassung'!$B${ZR['rs_db']}*IF({ilv};{K['ILV_RS']};1);-1)-'Zusammenfassung'!$B${ZR['abz_db']});"
            f"\"über Richtsatz Drehbuch\";\"\")",
            f"IF(AND({ISOEFI};{f('Nr.')}=\"24\";N({f('Plan brutto')})>ROUND('Zusammenfassung'!$B${ZR['rs_re']}*IF({ilv};{K['ILV_RS']};1);-1)-'Zusammenfassung'!$B${ZR['abz_re']});"
            f"\"über Richtsatz Regie\";\"\")",
            # nur FISA+
            f"IF(AND({ISFISA};{f('Nr.')}=\"13\";'Zusammenfassung'!$B${ZR['hl']}>{K['HL_MAX']});\"Herstellungsleitung über 2,5 % der Fertigungskosten\";\"\")",
        ]
        po.set(rr_, C["Prüfung"], "=TEXTJOIN(\", \";TRUE();" + ";".join(checks) + ")", CALC)
        po.set(rr_, C["Endkosten"], f"=IF(AND({f('Ist')}=\"\";{f('noch zu erwarten')}=\"\");\"\";N({f('Ist')})+N({f('noch zu erwarten')}))", CALC)
        po.set(rr_, C["± zu Plan"], f"=IF({f('Endkosten')}=\"\";\"\";N({f('Plan brutto')})-{f('Endkosten')})", CALC)

    # Eingaben aus der bestehenden Tabelle übernehmen (Schlüssel: Nr. und Vorkommen)
    nr_rows = {}
    for rr_ in range(first_data, last + 1):
        nr = po.cells.get((rr_, 2), "").lstrip("'")
        if nr:
            nr_rows.setdefault(nr, []).append(rr_)
    for key, vals in inputs.get("Positionen", {}).items():
        nr, idx = key.split("|")
        target = nr_rows[nr][int(idx)]
        for cname, v in vals.items():
            po.set(target, C[cname], v)

    def runs(rows_):
        """zusammenhängende Zeilenbereiche, z. B. [(13, 23), (25, 27)]"""
        out = []
        for x in sorted(rows_):
            if out and x == out[-1][1] + 1:
                out[-1][1] = x
            else:
                out.append([x, x])
        return out

    g3 = [r_ for r_, p in rows_meta.items() if p["gruppe"] == 3]
    g3a, g3b = min(g3), max(g3)
    stab_runs = runs(r_ for r_, p in rows_meta.items() if p["art"] == "Stab")
    money = [rng(n) for n in ("KV-Gage", "Preis", "Plan brutto", "LNK", "Plan gesamt", "förderbar", "Ist",
                              "noch zu erwarten", "Endkosten", "± zu Plan", "Hilfe: U-Bahn-Abgabe", "Hilfe: Gage")]
    po.ops += [{"op": "text", "range": rng("Nr.")}]
    # Partner-Auswahl je Gruppe (ohne Gruppenzeile); Department-Zeilen werden danach bereinigt
    bounds = group_rows + [last + 1]
    po.ops += [{"op": "dropdown", "range": f"{L['Partner']}{a + 1}:{L['Partner']}{b - 1}", "options": ["AT", "A", "B", "C"]}
               for a, b in zip(bounds, bounds[1:]) if b - 1 >= a + 1]
    po.ops += [{"op": "dropdown", "range": f"{L['Einheit']}{a}:{L['Einheit']}{b}", "options": ["Wochen 40 h", "Wochen 60 h", "Tage"]}
               for a, b in stab_runs]
    po.ops += [{"op": "currency", "range": m} for m in money]
    po.ops += [{"op": "percent", "range": rng("Abw. zu KV")}, {"op": "percent", "range": rng("Anteil förderbar")}]
    po.ops += [
        {"op": "checkbox", "range": f"{L['ILV']}{g3a}:{L['Rechnung']}{g3b}"},
        {"op": "checkbox", "range": f"{L['ohne U-Bahn-Abgabe']}{g3a}:{L['ohne U-Bahn-Abgabe']}{g3b}"},
        {"op": "dropdown", "range": f"{L['M/W']}{g3a}:{L['M/W']}{g3b}", "options": ["M", "W"]},
        {"op": "dropdown_range", "range": f"{L['KV-Funktion']}{g3a}:{L['KV-Funktion']}{g3b}", "source": f"='KV'!$A$2:$A${kv_last}"},
    ]
    # Department-Zeilen: Auswahlfelder und Kästchen wieder entfernen (leere Zelle im Department-Format darüber kopieren)
    for d in dept_rows:
        po.ops += [{"op": "copyfmt", "src": f"D{d}", "dst": f"{L['Partner']}{d}:{L['M/W']}{d}"},
                   {"op": "copyfmt", "src": f"D{d}", "dst": f"{L['ohne U-Bahn-Abgabe']}{d}"}]
    po.ops += [
        {"op": "cf_formula", "range": f"B{first_data}:{L['ohne U-Bahn-Abgabe']}{last}", "formula": f"=${L['Prüfung']}{first_data}<>\"\"", "color": "rot"},
        {"op": "freeze", "rows": 1, "cols": 3},
        {"op": "hide", "range": f"{L['Art']}:{L['Hilfe: 60 h']}"},
        {"op": "hide", "range": "A:A"},
    ]
    PO = {name: f"'Positionen'!${L[name]}${first_data}:${L[name]}" for name in cols}


    # ================================================================ Reisen
    rei = Sheet("Reisen")
    rei.set(1, 1, "Reisen, Aufenthalt, Catering", TITLE)
    rei.set(2, 1, "Fahrt = km × km-Geld + Flüge + Bahn → Pos. 160 · Aufenthalt = Diäten (Inland netto ÷ 1,1) + Nächte + Catering → Pos. 161 · "
                  "je Partner; Sätze im Blatt „Sätze“", NOTE)
    a, b = REI_CAT
    for c, h in enumerate(["Catering", "Partner", "Drehtage", "Personen", "€ pro Person und Tag", "Pauschale", "Summe"], 1):
        rei.set(a - 1, c, h, HEAD)
    for r in range(a, b + 1):
        rei.set(r, 1, "Catering" if r == a else "Catering (weitere)")
        rei.set(r, 2, "AT", INPUT)
        for c in range(3, 7):
            rei.set(r, c, "", INPUT)
        rei.set(r, 7, f"=IF(N(C{r})*N(D{r})*N(E{r})+N(F{r})=0;\"\";N(C{r})*N(D{r})*N(E{r})+N(F{r}))", CALC)
    rei.set(b + 1, 1, "Summe Catering", {"bold": True})
    rei.set(b + 1, 7, f"=SUM(G{a}:G{b})", {**CALC, "bold": True})
    head = ["Name", "Gruppe", "Department", "Partner", "km", "Flüge Anzahl", "Flug € je Stück", "Bahn Anzahl", "Bahn € je Fahrt",
            "Tage Ortsgebiet", "Tage Österreich", "Tage Ausland", "Satz Ausland €", "Nächte", "€ pro Nacht",
            "Fahrt", "Aufenthalt", "Gesamt"]
    f1, fl = REI_FIRST, REI_FIRST + N_REISEN - 1
    rei.set(f1 - 2, 1, "Summe Personen", {"bold": True})
    for c in (16, 17, 18):
        cl = col(c)
        rei.set(f1 - 2, c, f"=SUM({cl}{f1}:{cl})", {**CALC, "bold": True})
    for c, h in enumerate(head, 1):
        rei.set(f1 - 1, c, h, HEAD)
    for r in range(f1, fl + 1):
        for c in range(1, 16):
            rei.set(r, c, "", INPUT)
        rei.set(r, 3, f"=IF(B{r}=\"Cast\";\"Cast\";\"\")", INPUT)
        rei.set(r, 4, "AT", INPUT)
        fahrt = f"N(E{r})*{K['KM']}+N(F{r})*N(G{r})+N(H{r})*N(I{r})"
        aufenthalt = f"N(J{r})*{K['DIAET_W']}/1,1+N(K{r})*{K['DIAET']}/1,1+N(L{r})*N(M{r})+N(N{r})*N(O{r})"
        rei.set(r, 16, f"=IF({fahrt}=0;\"\";{fahrt})", CALC)
        rei.set(r, 17, f"=IF({aufenthalt}=0;\"\";{aufenthalt})", CALC)
        rei.set(r, 18, f"=IF(N(P{r})+N(Q{r})=0;\"\";N(P{r})+N(Q{r}))", CALC)
    rei.ops += [{"op": "dropdown", "range": f"B{a}:B{b}", "options": ["AT", "A", "B", "C"]},
                {"op": "currency", "range": f"E{a}:G{b + 1}"},
                {"op": "dropdown", "range": f"B{f1}:B{fl}", "options": ["Crew", "Cast", "Sonstige"]},
                {"op": "dropdown", "range": f"C{f1}:C{fl}", "options": [d for d, *_ in DEPARTMENTS]},
                {"op": "dropdown", "range": f"D{f1}:D{fl}", "options": ["AT", "A", "B", "C"]},
                {"op": "currency", "range": f"G{f1}:G{fl}"}, {"op": "currency", "range": f"I{f1}:I{fl}"},
                {"op": "currency", "range": f"M{f1}:M{fl}"}, {"op": "currency", "range": f"O{f1 - 2}:R{fl}"},
                {"op": "freeze", "rows": f1 - 1, "cols": 1}]

    # ================================================================ Motive
    mot = Sheet("Motive")
    mot.set(1, 1, "Motive", TITLE)
    mot.set(2, 1, "Eine Zeile je Motiv/Set; die Beträge fließen je Partner automatisch in die Positionen (Nr. im Spaltenkopf).", NOTE)
    m1, ml = MOT_FIRST, MOT_FIRST + N_MOTIVE - 1
    head = ["Motiv/Set", "innen/außen", "Partner", "Drehtage"] + [f"{name} ({nr})" for _, name, nr in MOTIV_COLS] + ["Summe", "Anmerkungen"]
    mot.set(m1 - 3, 1, "Summe", {"bold": True})
    for c in range(5, 17):
        cl = col(c)
        mot.set(m1 - 3, c, f"=SUM({cl}{m1}:{cl})", {**CALC, "bold": True})
    for c, h in enumerate(head, 1):
        mot.set(m1 - 1, c, h, HEAD)
    for r in range(m1, ml + 1):
        for c in list(range(1, 16)) + [17]:
            mot.set(r, c, "", INPUT)
        mot.set(r, 3, "AT", INPUT)
        mot.set(r, 16, f"=IF(SUM(E{r}:O{r})=0;\"\";SUM(E{r}:O{r}))", CALC)
    mot.ops += [{"op": "dropdown", "range": f"B{m1}:B{ml}", "options": ["innen", "außen"]},
                {"op": "dropdown", "range": f"C{m1}:C{ml}", "options": ["AT", "A", "B", "C"]},
                {"op": "currency", "range": f"E{m1 - 3}:P{ml}"},
                {"op": "freeze", "rows": m1 - 1, "cols": 1}]

    # ================================================================ Zusammenfassung
    zs = Sheet("Zusammenfassung")
    zs.set(1, 1, "Zusammenfassung", TITLE)
    zs.set(2, 1, "Positionen mit Hinweis (Spalte „Prüfung“)", {"bold": True})
    zs.set(2, 2, f"=COUNTIF({PO['Prüfung']};\"?*\")", {**CALC, "bold": True})
    zs.set(2, 3, f"=\"Förderstelle: \"&{FST}", NOTE)
    partners = ["AT", "A", "B", "C"]
    for c, h in enumerate(["Kostengruppe", "AT", "A", "B", "C", "Gesamt", "Endkosten (brutto)", "± zu Plan", "förderbar"], 1):
        zs.set(3, c, h, HEAD)
    for g in range(1, 13):
        r = 3 + g
        zs.set(r, 1, f"{g}. {gnames[g]}")
        for i, pt in enumerate(partners):
            base = (f"SUMIFS({PO['Plan brutto']};{PO['Gr.']};{g};{PO['Partner']};\"{pt}\";{PO['Art']};\"<>Info\";{PO['Art']};\"<>Gruppe\")"
                    f"+SUMIFS({PO['LNK']};{PO['Gr.']};{g};{PO['Partner']};\"{pt}\")")
            if g == 3 and pt == "AT":
                base += f"+IF({ISOEFI};$B${ZR['ph_oefi']};0)"
            zs.set(r, 2 + i, "=" + base, CALC)
        zs.set(r, 6, f"=SUM(B{r}:E{r})", CALC)
        zs.set(r, 7, f"=SUMIFS({PO['Endkosten']};{PO['Gr.']};{g};{PO['Art']};\"<>Info\";{PO['Art']};\"<>Gruppe\")", CALC)
        zs.set(r, 8, f"=SUMIFS({PO['± zu Plan']};{PO['Gr.']};{g};{PO['Art']};\"<>Info\";{PO['Art']};\"<>Gruppe\")", CALC)
        zs.set(r, 9, f"=SUMIFS({PO['förderbar']};{PO['Gr.']};{g};{PO['Art']};\"<>Info\";{PO['Art']};\"<>Gruppe\")"
                     + (f"+IF({ISOEFI};$B${ZR['ph_oefi']};0)" if g == 3 else ""), CALC)
    staffel_fgk = ("ROUND(IF(B16>=5000000;375000;IF(B16<=800000;B16*9%;IF(B16<=1200000;72000+(B16-800000)*30000/400000;"
                   "IF(B16<=1600000;102000+(B16-1200000)*26000/400000;IF(B16<=2500000;128000+(B16-1600000)*59500/900000;B16*7,5%)))));-1)")
    rows = [
        (16, "Fertigungskosten (FK)", ["=SUM(B4:B15)", "=SUM(C4:C15)", "=SUM(D4:D15)", "=SUM(E4:E15)"], "=SUM(I4:I15)"),
        (17, "Honorar Produzent*in (FISA+ Pkt. 14; ÖFI: in Gruppe 3)", [f"=IF({ISFISA};ROUND(B16*{PHP};2);0)", "", "", ""], "=B17"),
        (18, "Fertigungsgemeinkosten (FGK)", [f"=IF({ISFISA};ROUND(B16*{K['FGK_MAX']};2);IF({BMWKMS}=TRUE();B16*7,5%;{staffel_fgk}))", "", "", ""],
         f"=IF({ISFISA};B18;0)"),
        (19, "Completion Bond", [f"=N({BOND})", "", "", ""], f"=IF({ISFISA};B19;0)"),
        (20, "Finanzierungskosten", [f"=N({FIN})", "", "", ""], f"=IF({ISFISA};B20;0)"),
        (21, "Überschreitungsreserve", [f"=ROUND(B16*{UER};-1)", "", "", ""], f"=ROUND(I16*{UER};-1)"),
        (22, "Herstellungskosten (HK)", ["=SUM(B16:B21)", "=SUM(C16:C21)", "=SUM(D16:D21)", "=SUM(E16:E21)"], "=SUM(I16:I21)"),
    ]
    for r, label, vals, fz in rows:
        bold = r in (16, 22)
        zs.set(r, 1, label, {"bold": bold})
        for i, v in enumerate(vals):
            if v:
                zs.set(r, 2 + i, v, {**CALC, "bold": bold})
        zs.set(r, 6, f"=SUM(B{r}:E{r})", {**CALC, "bold": bold})
        zs.set(r, 9, fz, {**CALC, "bold": bold})
    zs.set(16, 7, "=SUM(G4:G15)", CALC)
    zs.set(16, 8, "=SUM(H4:H15)", CALC)

    # Anreizförderung
    zs.set(24, 1, "Anreizförderung", HEAD)
    zs.set(24, 2, "", HEAD)
    zs.set(24, 3, "Erklärung", HEAD)
    cap = f"IF({ISFISA};IF({FORMAT}=\"Serie\";MIN({K['MAX_SERIE']};{K['MAX_FILM']}*MAX(1;N({FOLGEN})));{K['MAX_FILM']});9,99E+15)"
    anreiz = [
        (25, "Bemessungsgrundlage (förderbar, max. 80 % der HK gesamt)", f"=MIN($I$22;{K['BEM_MAX']}*$F$22)",
         "Summe „förderbar“; FISA+: inkl. Honorar, FGK, Bond, Finanzierung; ÖFI+: Fertigungskosten und Überschreitungsreserve"),
        (26, "Höchstbetrag Zuschuss", f"=IF({ISFISA};{cap};\"\")", "FISA+: 5 Mio. € je Film/Folge, 7,5 Mio. € je Serie"),
        (27, "Basis", f"=IF(AND({ISOEFI};{BMWKMS}<>TRUE());0;IF({ISFISA};MIN($B$25*{K['ANREIZ']};N($B$26));ROUNDDOWN($B$25*{K['ANREIZ']};-2)))",
         "FISA+ 30 %; ÖFI+ 30 % (nur mit BMWKMS-Förderung)"),
        (28, "Green Bonus", f"=IF({GREEN}<>TRUE();0;IF(AND({ISOEFI};{BMWKMS}<>TRUE());0;IF({ISFISA};MAX(0;MIN($B$25*{K['GREEN']};N($B$26)-$B$27));"
                            f"ROUNDDOWN($B$25*{K['GREEN']};-2))))", "zusätzlich 5 Prozentpunkte, wenn beantragt"),
        (29, "Gender Gap Financing", f"=IF({GGF}<>TRUE();0;IF(AND({ISOEFI};{BMWKMS}<>TRUE());0;IF({ISFISA};MAX(0;MIN({K['GGF']};N($B$26)-$B$27-$B$28));{K['GGF']})))",
         "zusätzlich 25.000 €, wenn beantragt"),
        (30, "Anreizförderung gesamt", "=SUM(B27:B29)", "fließt automatisch ins Blatt Finanzierung"),
    ]
    for r, label, formula, note in anreiz:
        zs.set(r, 1, label, {"bold": r == 30})
        zs.set(r, 2, formula, {**CALC, "bold": r == 30})
        zs.set(r, 3, note, NOTE)

    # Prüfungen zur Förderstelle
    zs.set(32, 1, "Prüfungen Förderstelle", HEAD)
    zs.set(32, 2, "Ergebnis", HEAD)
    zs.set(32, 3, "", HEAD)
    FINS = "'Finanzierung'"
    ph_used = f"IF({ISFISA};$B$17;$B${ZR['ph_oefi']})"
    checks = [
        (33, "Überschreitungsreserve im Rahmen",
         f"=IF(OR({UER}<{K['UER_MIN']};{UER}>IF(AND({ISOEFI};N({BOND})>0);10%;{K['UER_MAX']}));\"außerhalb \"&TEXT({K['UER_MIN']};\"0 %\")&\" bis \"&TEXT({K['UER_MAX']};\"0 %\");\"OK\")"),
        (34, "Herstellungsleitung (FISA+ max. 2,5 % FK)", f"=IF({ISFISA};IF($B${ZR['hl']}>{K['HL_MAX']};\"zu hoch: \"&TEXT($B${ZR['hl']};\"0.00 %\");\"OK\");\"–\")"),
        (35, "Honorar Produzent*in höchstens laut Förderstelle",
         f"=IF({ph_used}>$B${ZR['ph_max']}+0,5;\"zu hoch: \"&TEXT({ph_used};\"#,##0 €\")&\" (höchstens \"&TEXT($B${ZR['ph_max']};\"#,##0 €\")&\")\";\"OK\")"),
        (36, "förderbare Kosten über Mindestbetrag (FISA+)",
         f"=IF({ISFISA};IF($B$25<IF({FILMART}=\"Dokumentarfilm\";{K['MIN_F_DOKU']};{K['MIN_F_SPIEL']});\"unter Mindestbetrag\";\"OK\");\"–\")"),
        (37, "Gesamt-HK über Mindestbetrag (FISA+)",
         f"=IF({ISFISA};IF(IF({FORMAT}=\"Serie\";$F$22/MAX(1;N({FOLGEN}))<{K['MIN_HK_FOLGE']};$F$22<{K['MIN_HK_FILM']});\"unter Mindest-HK\";\"OK\");\"–\")"),
        (38, "Mediendienstanbieter mind. 30 % der HK (FISA+)",
         f"=IF({ISFISA};IF({FINS}!$B$8<{K['MDA_MIN']};\"nur \"&TEXT({FINS}!$B$8;\"0 %\");\"OK\");\"–\")"),
        (39, "Eigenanteil mind. 2,5 % der HK",
         f"=IF({FINS}!$B$9<{K['EIGEN']};\"nur \"&TEXT({FINS}!$B$9;\"0.0 %\");\"OK\")"),
        (40, "Werkstattprojekt: Stufe passt zur HK (ÖFI)",
         f"=IF(OR({ISFISA};{WERKSTATT}=\"nein\";{WERKSTATT}=\"\");\"–\";IF($B$22>IF({WERKSTATT}=\"Werkstatt-KV\";{K['WS1']};{K['WS2']});"
         f"\"HK über \"&TEXT(IF({WERKSTATT}=\"Werkstatt-KV\";{K['WS1']};{K['WS2']});\"#,##0 €\");\"OK\"))"),
    ]
    for r, label, formula in checks:
        zs.set(r, 1, label)
        zs.set(r, 2, formula, CALC)

    zs.set(ZR["ph_oefi"] - 1, 1, "Berechnete Werte", HEAD)
    zs.set(ZR["ph_oefi"] - 1, 2, "", HEAD)
    zs.set(ZR["ph_oefi"] - 1, 3, "Erklärung", HEAD)
    basis = (f"SUMIFS({PO['Plan brutto']};{PO['Partner']};\"AT\";{PO['Art']};\"<>Info\";{PO['Art']};\"<>Gruppe\";{PO['Art']};\"<>Department\")"
             f"+SUMIFS({PO['LNK']};{PO['Partner']};\"AT\")")
    B_ = f"$B${ZR['basis']}"
    ph_max = (f"=IF({ISFISA};ROUND($B$16*{K['PH_MAX']};2);IF({BMWKMS}=TRUE();LET(basis;{B_};ueber;{UER};zusatz;N({BOND})+N({FIN});"
              f"schaetz;ROUND(0,05*(basis*(1,075+ueber)+zusatz)/(1-0,05*(1,075+ueber));-1);"
              f"erste;ROUND(0,05*((basis+schaetz)*1,075+ROUND((basis+schaetz)*ueber;-1)+zusatz);-1);"
              f"zweite;ROUND(0,05*((basis+erste)*1,075+ROUND((basis+erste)*ueber;-1)+zusatz);-1);zweite);"
              f"LET(basis;{B_};IF(basis<=180000;basis/9;IF(basis<=925000;(6250+basis*0,06875)/0,93125;"
              f"IF(basis<=1400000;(25000+basis*0,05)/0,95;IF(basis<=2375000;(62500+basis*0,025)/0,975;"
              f"IF(basis<=4750000;basis/19;250000))))))))")
    calc = [
        (ZR["ph_oefi"], "ÖFI: Produzent*innenhonorar (Pos. 87)", f"=IF({ISFISA};0;ROUND({B_}*{PHP};2))",
         "Anteil laut Projekt × Fertigungskosten AT ohne Honorar; bei FISA+ siehe Zeile 17"),
        (ZR["basis"], "Fertigungskosten AT ohne Honorar", "=" + basis, "Basis für das Honorar"),
        (ZR["ph_max"], "Honorar: Höchstwert der Förderstelle", ph_max,
         "FISA+ 5 % der FK; ÖFI-Staffel ohne Zirkelbezug gelöst (BMWKMS: 5 % der HK)"),
        (ZR["rs_db"], "ÖFI: Richtsatz Drehbuch", f"=IF({ISFISA};\"\";LET(n;$F$16;rs;ROUND(IF(n<=200000;n*12%;IF(n<=500000;24000+(n-200000)*5%;IF(n<=800000;39000+(n-500000)*3%;"
             f"IF(n<=2800000;48000+(n-800000)*2%;IF(n<=6000000;88000+(n-2800000)*1%;n*2%)))));-1);IF({DEBUT_DB}=TRUE();ROUND(rs*{K['DEBUT']};-1);rs)))",
         "nur ÖFI; auf die Fertigungskosten gesamt; Debut 85 %"),
        (ZR["rs_re"], "ÖFI: Richtsatz Regie", f"=IF({ISFISA};\"\";LET(n;$F$16;rs;ROUND(IF(n<=200000;n*18%;IF(n<=300000;36000+(n-200000)*6%;IF(n<=600000;42000+(n-300000)*4%;"
             f"IF(n<=2400000;54000+(n-600000)*3%;IF(n<=3000000;108000+(n-2400000)*2%;IF(n<=4500000;120000+(n-3000000)*1%;n*3%))))));-1);"
             f"IF({DEBUT_RE}=TRUE();ROUND(rs*{K['DEBUT']};-1);rs)))", "nur ÖFI; auf die Fertigungskosten gesamt; Debut 85 %"),
        (ZR["abz_db"], "ÖFI: Abzug Drehbuch (Vorförderung über Freibetrag)",
         f"=MAX(0;SUMIFS({PO['Plan brutto']};{PO['Nr.']};\"1.1a\")-{K['FREIBETRAG']})", "Pos. 1.1a abzüglich 15.000 €"),
        (ZR["abz_re"], "ÖFI: Abzug Regie (Vorförderung über Freibetrag)",
         f"=MAX(0;SUMIFS({PO['Plan brutto']};{PO['Nr.']};\"1.1b\")-{K['FREIBETRAG']})", "Pos. 1.1b abzüglich 15.000 €"),
        (ZR["ueber"], "ÖFI: Überzahlungen über 10 % (FRL Pkt. 6.2.3)",
         f"=IF({ISFISA};0;SUMPRODUCT(({PO['Art']}=\"Stab\")*({PO['KV-Gage']}<>\"\")*({PO['Rechnung']}<>TRUE)*"
         f"IFERROR(({PO['Hilfe: Gage']}-{PO['KV-Gage']}*{K['UEBER']})*{PO['Menge']}*(({PO['Hilfe: Gage']}-{PO['KV-Gage']}*{K['UEBER']})>0);0)))",
         "nur anerkannt, wenn im Finanzierungsplan gedeckt"),
        (ZR["hl"], "Herstellungsleitung: Anteil an den Fertigungskosten",
         f"=IF($B$16=0;0;SUMIFS({PO['Plan gesamt']};{PO['Nr.']};\"13\";{PO['Partner']};\"AT\")/$B$16)", "FISA+: höchstens 2,5 %"),
    ]
    for r, label, formula, note in calc:
        zs.set(r, 1, label)
        zs.set(r, 2, formula, CALC)
        zs.set(r, 3, note, NOTE)

    lh, l1, ls = ZR["lnk_head"], ZR["lnk_first"], ZR["lnk_sum"]
    zs.set(lh, 1, "Lohnnebenkosten (Pos. 84)", HEAD)
    for i, pt in enumerate(partners):
        zs.set(lh, 2 + i, pt, HEAD)
    zs.set(lh, 6, "Gesamt", HEAD)
    for i, pt in enumerate(partners):
        cc = col(2 + i)
        stab = f"SUMIFS({PO['Plan brutto']};{PO['Art']};\"Stab\";{PO['Partner']};\"{pt}\";{PO['Rechnung']};FALSE())"
        cast = f"SUMIFS({PO['Plan brutto']};{PO['Art']};\"Cast\";{PO['Partner']};\"{pt}\";{PO['Rechnung']};FALSE())"
        zs.set(l1 + 1, 2 + i, f"={stab}*{K['MV']}", CALC)
        zs.set(l1 + 2, 2 + i, f"=({stab}+{cast})*{K['DB']}", CALC)
        zs.set(l1 + 3, 2 + i, f"=({stab}+{cast})*{K['DZ']}", CALC)
        zs.set(l1 + 4, 2 + i, f"=({stab}+{cast})*{K['KOMMST']}", CALC)
        zs.set(l1 + 5, 2 + i, f"=SUMIFS({PO['Hilfe: U-Bahn-Abgabe']};{PO['Partner']};\"{pt}\")", CALC)
        zs.set(ls, 2 + i, f"=SUMIFS({PO['LNK']};{PO['Partner']};\"{pt}\")", {**CALC, "bold": True})
        zs.set(l1, 2 + i, f"={cc}{ls}-SUM({cc}{l1 + 1}:{cc}{l1 + 5})", CALC)
    for off, label in enumerate(["Sozialversicherung (inkl. SZ/UEL)", "MV-Beitrag", "DB", "DZ", "Kommunalsteuer",
                                 "U-Bahn-Abgabe", "Summe Lohnnebenkosten"]):
        r = l1 + off
        zs.set(r, 1, label, {"bold": r == ls})
        zs.set(r, 6, f"=SUM(B{r}:E{r})", {**CALC, "bold": r == ls})

    gh = ZR["gender_head"]
    zs.set(gh, 1, "Gender-Budgeting (Plan brutto Stab und Cast)", HEAD)
    zs.set(gh, 2, "Betrag", HEAD)
    zs.set(gh, 3, "Anteil", HEAD)
    for off, (label, crit) in enumerate([("weiblich (W)", "\"W\""), ("männlich (M)", "\"M\""), ("ohne Angabe", "\"\"")], 1):
        r = gh + off
        zs.set(r, 1, label)
        zs.set(r, 2, f"=SUMIFS({PO['Plan brutto']};{PO['Art']};\"Stab\";{PO['M/W']};{crit})"
                     f"+SUMIFS({PO['Plan brutto']};{PO['Art']};\"Cast\";{PO['M/W']};{crit})", CALC)
        zs.set(r, 3, f"=IF(SUM($B${gh + 1}:$B${gh + 3})=0;\"\";B{r}/SUM($B${gh + 1}:$B${gh + 3}))", CALC)

    # Departments (Gruppe 3): Summen über die Spalte „Department“ der Positionen
    dh = ZR["dep_head"]
    for c, h in enumerate(["Departments (3. Gagen, Löhne, Honorare)", "Plan brutto", "LNK", "Plan gesamt",
                           "Endkosten (brutto)", "± zu Plan", "Anteil an Gruppe 3", "Reise & Aufenthalt", "Department gesamt"], 1):
        zs.set(dh, c, h, HEAD)
    REI_R, REI_C = f"'Reisen'!$R${REI_FIRST}:$R", f"'Reisen'!$C${REI_FIRST}:$C"
    for i, (name, *_ ) in enumerate(DEPARTMENTS):
        r = dh + 1 + i
        zs.set(r, 1, name)
        for c, pc in [(2, "Plan brutto"), (3, "LNK"), (4, "Plan gesamt"), (5, "Endkosten"), (6, "± zu Plan")]:
            zs.set(r, c, f"=SUMIFS({PO[pc]};{PO['Department']};$A{r};{PO['Art']};\"<>Department\")", CALC)
        zs.set(r, 7, f"=IF($D${dh + len(DEPARTMENTS) + 1}=0;\"\";D{r}/$D${dh + len(DEPARTMENTS) + 1})", CALC)
        zs.set(r, 8, f"=SUMIFS({REI_R};{REI_C};$A{r})", CALC)
        zs.set(r, 9, f"=D{r}+H{r}", CALC)
    dsum = dh + len(DEPARTMENTS) + 1
    zs.set(dsum, 1, "Summe Gruppe 3", {"bold": True})
    for c in list(range(2, 7)) + [8, 9]:
        cl = col(c)
        zs.set(dsum, c, f"=SUM({cl}{dh + 1}:{cl}{dsum - 1})", {**CALC, "bold": True})
    zs.set(dsum + 1, 1, "Reise ohne Department-Zuordnung und Catering")
    zs.set(dsum + 1, 8, f"=SUM('Reisen'!$R${REI_FIRST}:$R)-H{dsum}+SUM('Reisen'!$G${REI_CAT[0]}:$G${REI_CAT[1]})", CALC)
    zs.set(dsum, 7, f"=IF(ROUND(D{dsum}-$F$6;2)=0;\"\";\"Abweichung zu Zeile 6: \"&TEXT(D{dsum}-$F$6;\"#,##0.00 €\"))", NOTE)

    zs.ops += [{"op": "currency", "range": "B4:I22"}, {"op": "currency", "range": "B25:B30"},
               {"op": "currency", "range": f"B{ZR['ph_oefi']}:B{ZR['ueber']}"}, {"op": "percent", "range": f"B{ZR['hl']}"},
               {"op": "currency", "range": f"B{l1}:F{ls}"}, {"op": "currency", "range": f"B{gh + 1}:B{gh + 3}"},
               {"op": "percent", "range": f"C{gh + 1}:C{gh + 3}"},
               {"op": "currency", "range": f"B{dh + 1}:F{dsum}"}, {"op": "percent", "range": f"G{dh + 1}:G{dsum - 1}"},
               {"op": "currency", "range": f"H{dh + 1}:I{dsum + 1}"},
               {"op": "cf_formula", "range": "A2:B2", "formula": "=$B$2>0", "color": "rot"},
               {"op": "cf_formula", "range": "B33:B40", "formula": "=AND(B33<>\"OK\";B33<>\"–\")", "color": "rot"},
               {"op": "freeze", "rows": 3}]

    # ================================================================ Finanzierung
    fi = Sheet("Finanzierung")
    fi.set(1, 1, "Finanzierung", TITLE)
    fi.set(2, 1, "Status: geplant, beantragt, zugesagt oder abgelehnt. Abgelehnte Quellen zählen nicht zur Finanzierung.", NOTE)
    HK_G = "'Zusammenfassung'!$F$22"
    HK_AT = "'Zusammenfassung'!$B$22"
    top = 12  # Kopfzeile der Tabelle
    D_ = f"$D${top + 1}:$D"
    A_ = f"$A${top + 1}:$A"
    E_ = f"$E${top + 1}:$E"
    summary = [
        (3, "Herstellungskosten gesamt", f"={HK_G}", "currency"),
        (4, "Finanzierung (ohne abgelehnt)", f"=SUMIFS({D_};{E_};\"<>abgelehnt\")", "currency"),
        (5, "davon zugesagt", f"=SUMIFS({D_};{E_};\"zugesagt\")", "currency"),
        (6, "Lücke zur HK", "=B3-B4", "currency"),
        (7, "Anreizförderung (automatisch)", "='Zusammenfassung'!$B$30", "currency"),
        (8, "Anteil Mediendienstanbieter an der HK", f"=IF(B3=0;0;SUMIFS({D_};{A_};\"Mediendienstanbieter\";{E_};\"<>abgelehnt\")/B3)", "percent"),
        (9, "Anteil Eigenanteil an der HK (AT)", f"=IF({HK_AT}=0;0;SUMIFS({D_};{A_};\"Eigenanteil\";{E_};\"<>abgelehnt\")/{HK_AT})", "percent"),
        (10, "Prüfung", f"=TEXTJOIN(\", \";TRUE();IF(ROUND(B6;0)>0;\"Lücke \"&TEXT(B6;\"#,##0 €\");\"\");IF(ROUND(B6;0)<0;\"Überfinanzierung \"&TEXT(-B6;\"#,##0 €\");\"\");"
                        f"IF(AND({ISFISA};B8<{K['MDA_MIN']});\"Mediendienstanbieter unter 30 %\";\"\");IF(B9<{K['EIGEN']};\"Eigenanteil unter 2,5 %\";\"\"))", None),
    ]
    for r, label, formula, fmt in summary:
        fi.set(r, 1, label, {"bold": r in (4, 6, 10)})
        fi.set(r, 2, formula, {**CALC, "bold": r in (4, 6, 10)})
    for c, h in enumerate(["Bereich", "Quelle", "Partner", "Betrag", "Status", "Anteil an HK", "Notiz"], 1):
        fi.set(top, c, h, HEAD)
    ilv = f"SUMIFS({PO['Plan gesamt']};{PO['ILV']};TRUE();{PO['Partner']};\"AT\")"
    anr = lambda zr: f"='Zusammenfassung'!$B${zr}"
    src = [
        ("Eigenanteil", "Eigenmittel (Barmittel)", "AT", f"=ROUNDUP({HK_AT}*{K['EIGEN']};-2)", "geplant", "Vorschlag 2,5 % der HK; überschreiben"),
        ("Eigenanteil", "Bewertete Eigenleistungen (ILV)", "AT", f"={ilv}", "geplant", "Summe der Positionen mit ILV-Häkchen"),
        ("Eigenanteil", "Rückstellungen (Honorar, FGK)", "AT", "", "geplant", ""),
        ("Eigenanteil", "Darlehen, zusätzliche Barmittel", "AT", "", "geplant", ""),
        ("Mediendienstanbieter", "Lizenzanteil (mind. 50 % des MDA-Betrags)", "AT", "", "geplant", "z. B. ORF, ServusTV, Streamer"),
        ("Mediendienstanbieter", "Koproduktionsanteil", "AT", "", "geplant", ""),
        ("Lizenzen, Garantien", "Verleihgarantie", "AT", "", "geplant", ""),
        ("Lizenzen, Garantien", "Weltvertrieb (MG)", "AT", "", "geplant", ""),
        ("Anreizförderung", f"=IF({ISFISA};\"FISA+ Zuschuss\";\"ÖFI+ Basisfinanzierung\")&\" (\"&TEXT({K['ANREIZ']};\"0 %\")&\")\"", "AT", anr(27), "geplant", "automatisch aus der Zusammenfassung"),
        ("Anreizförderung", "Green Bonus", "AT", anr(28), "geplant", "automatisch, wenn im Projekt beantragt"),
        ("Anreizförderung", "Gender Gap Financing", "AT", anr(29), "geplant", "automatisch, wenn im Projekt beantragt"),
        ("Förderung", "ÖFI Projektförderung (selektiv)", "AT", "", "geplant", ""),
        ("Förderung", "ÖFI Referenzfilmförderung", "AT", "", "geplant", ""),
        ("Förderung", "Fernsehfonds Austria", "AT", "", "geplant", "inkl. Exzellenzbonus"),
        ("Förderung", "Filmfonds Wien", "AT", "", "geplant", ""),
        ("Förderung", "Länderförderung (z. B. Cine Tirol)", "AT", "", "geplant", ""),
        ("Förderung", "EURIMAGES", "AT", "", "geplant", ""),
        ("Förderung", "Sonstige", "AT", "", "geplant", ""),
        ("Ko-Partner", "Finanzierung Partner A", "A", "", "geplant", ""),
        ("Ko-Partner", "Finanzierung Partner B", "B", "", "geplant", ""),
        ("Ko-Partner", "Finanzierung Partner C", "C", "", "geplant", ""),
    ]
    for i, (bereich, quelle, partner, betrag, status, notiz) in enumerate(src):
        r = top + 1 + i
        ov = inputs.get("Finanzierung", {}).get(quelle, {})
        quelle, partner, betrag, status = (ov.get("Quelle", quelle), ov.get("Partner", partner),
                                           ov.get("Betrag", betrag), ov.get("Status", status))
        fi.set(r, 1, bereich)
        fi.set(r, 2, quelle, CALC if quelle.startswith("=") else None)
        fi.set(r, 3, partner)
        fi.set(r, 4, betrag, CALC if betrag.startswith("=") else INPUT)
        fi.set(r, 5, status, INPUT)
        fi.set(r, 6, f"=IF(OR($B$3=0;D{r}=\"\");\"\";D{r}/$B$3)", CALC)
        fi.set(r, 7, notiz, NOTE)
    fi_last = top + len(src)
    fi.ops += [{"op": "currency", "range": "B3:B7"}, {"op": "percent", "range": "B8:B9"},
               {"op": "currency", "range": f"D{top + 1}:D{fi_last + 20}"}, {"op": "percent", "range": f"F{top + 1}:F{fi_last + 20}"},
               {"op": "dropdown", "range": f"E{top + 1}:E{fi_last}", "options": ["geplant", "beantragt", "zugesagt", "abgelehnt"]},
               {"op": "dropdown", "range": f"C{top + 1}:C{fi_last}", "options": ["AT", "A", "B", "C"]},
               {"op": "dropdown", "range": f"A{top + 1}:A{fi_last}", "options": ["Eigenanteil", "Mediendienstanbieter", "Lizenzen, Garantien",
                                                                                 "Anreizförderung", "Förderung", "Ko-Partner"]},
               {"op": "cf_formula", "range": "A10:B10", "formula": "=$B$10<>\"\"", "color": "rot"},
               {"op": "freeze", "rows": top}]

    # ================================================================ Übertrag Antrag
    ue = Sheet(UE)
    head = ["Gr.", "Nr.", "Position", "Plan AT", "Wochen 40 h / Menge AT", "Ø Preis AT", "Wochen 60 h AT",
            "Ø Wochengage 60 h AT", "Ist AT", "noch zu erwarten AT", "Plan A", "Ist A", "noch zu erwarten A",
            "Plan B", "Ist B", "noch zu erwarten B", "Plan C", "Ist C", "noch zu erwarten C", "Plan gesamt",
            "Art (Liste)", "KV-Funktion (Liste)", "Department (Liste)"]
    for c, h in enumerate(head, 1):
        ue.set(1, c, h, HEAD)
    ue.set(2, 3, "Summen je Position im österreichischen Kalkulationsschema (aws/FISA+ und ÖFI). Werte blockweise in den Antrag übertragen.", NOTE)
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
            ue.set(r, 21, p["art"])
            if g == 3 and department(nr):
                ue.set(r, 23, department(nr))
            ks, kd = p.get("kv_spielfilm", ""), p.get("kv_doku", "")
            if p["art"] == "Stab" and ks and kd and ks != kd:
                ue.set(r, 22, f"=IF({FILMART}=\"Dokumentarfilm\";\"{kd}\";\"{ks}\")")
            elif p["art"] == "Stab" and ks:
                ue.set(r, 22, ks)
            if p["art"] == "Berechnet" and nr == "87":
                ue.set(r, 4, f"=IF({ISFISA};'Zusammenfassung'!$B$17;'Zusammenfassung'!$B${ZR['ph_oefi']})", CALC)
                ue.set(r, 20, f"=D{r}", CALC)
            elif p["art"] == "Berechnet" and nr == "84":
                for cc, zc in [(4, "B"), (11, "C"), (14, "D"), (17, "E"), (20, "F")]:
                    ue.set(r, cc, f"='Zusammenfassung'!${zc}${ZR['lnk_sum']}", CALC)
                r += 1
                for zr, label in [(ZR["lnk_first"] + i, lbl) for i, lbl in enumerate(
                        ["Sozialversicherung", "MV-Beitrag", "DB", "DZ", "Kommunalsteuer", "U-Bahn-Abgabe"])]:
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
                    ue.set(r, 4, f"=SUMIFS({PO['Plan brutto']};{PO['Nr.']};$B{r}&\".*\";{PO['Partner']};\"AT\")", CALC)
                else:
                    ue.set(r, 4, s("Plan brutto", "AT"), CALC)
                    # 40 h bzw. Menge und 60 h getrennt nach Einheit; Stab-Tage zählen als 1/5 Woche
                    at = f"{crit};{PO['Partner']};\"AT\""
                    ue.set(r, 5, f"=SUMIFS({PO['Hilfe: Wochen']};{at};{PO['Hilfe: 60 h']};FALSE())", CALC)
                    ue.set(r, 6, f"=IF(N(E{r})=0;\"\";SUMIFS({PO['Plan brutto']};{at};{PO['Hilfe: 60 h']};FALSE();{PO['Menge']};\"<>\")/E{r})", CALC)
                    ue.set(r, 7, f"=SUMIFS({PO['Hilfe: Wochen']};{at};{PO['Hilfe: 60 h']};TRUE())", CALC)
                    ue.set(r, 8, f"=IF(N(G{r})=0;\"\";SUMIFS({PO['Plan brutto']};{at};{PO['Hilfe: 60 h']};TRUE())/G{r})", CALC)
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
    ue.set(r, 1, "Summenblatt (Kalkulationsschema Pkt. 13–19)", HEAD)
    for c in range(2, 5):
        ue.set(r, c, "", HEAD)
    r += 1
    for zr, label in [(16, "13. Fertigungskosten"), (17, "14. Honorar Produzent*in (FISA+)"), (18, "15. Fertigungsgemeinkosten"),
                      (19, "16. Fertigstellungsversicherung"), (20, "17. Kosten der Finanzierung"), (21, "18. Überschreitungsreserve"),
                      (22, "19. Herstellungskosten")]:
        ue.set(r, 3, label, {"bold": zr in (16, 22)})
        ue.set(r, 4, f"='Zusammenfassung'!$B${zr}", CALC)
        ue.set(r, 20, f"='Zusammenfassung'!$F${zr}", CALC)
        r += 1
    ue.ops += [{"op": "text", "range": f"B3:B{ue_last}"},
               {"op": "currency", "range": f"D3:D{r}"}, {"op": "currency", "range": f"F3:F{ue_last}"},
               {"op": "currency", "range": f"H3:T{r}"}, {"op": "freeze", "rows": 1, "cols": 3},
               {"op": "hide", "range": "U:W"}]

    return [pr, po, rei, mot, zs, fi, ue, sa, kv]


if __name__ == "__main__":
    inputs = json.load(open(sys.argv[3])) if len(sys.argv) > 3 else {}
    sheets = build(sys.argv[1], sys.argv[2], inputs)
    json.dump([s.dump() for s in sheets], sys.stdout, ensure_ascii=False)
