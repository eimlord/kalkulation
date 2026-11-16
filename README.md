# Kalkulationsvorlage 2027 als Claude Sheet

Umbau der ÖFI-Kalkulationsvorlage `20270826_EMPTY_TEMPLATE.xlsx` (18 Blätter, ~11.500 Formeln)
in ein Claude Sheet. Die Skripte in `tools/` sind so geschrieben, dass sich eine neue Vorlage
(z. B. 2028) auf dem gleichen Weg umwandeln lässt.

## Inhalt

| Pfad | Zweck |
| --- | --- |
| `vorlage/20270826_EMPTY_TEMPLATE.xlsx` | Original-Vorlage |
| `vorlage/20270826_namen_ersetzt.xlsx` | Importierte Fassung: benannte Bereiche durch Zellbezüge ersetzt |
| `tools/replace_names.py` | Ersetzt benannte Bereiche im Formel-XML (Claude Sheets kennt keine) |
| `tools/extract_controls.py` | Liest Excel-Checkboxen und Kombinationsfelder samt Position und Hilfszelle |
| `tools/build_controls.py` | Erzeugt Zell-Checkboxen/Dropdowns und Formeln für die Hilfszellen |
| `tools/build_groups.py` | Excel-Gliederung und Ausblendungen → zugeklappte Gruppen |
| `tools/image_tables.py` | Abgetippte Bildtabellen (Diäten, Richtsätze) |
| `tools/build_hinweise.py` | Kommentare, Eingabehilfen und Umbau-Notizen für das Blatt „Hinweise“ |
| `tools/extract_rules.py` | Liest bedingte Formatierungen inkl. x14-Erweiterungen |
| `tools/build_rules.py` | Übersetzt sie (und Eingabegrenzen) in Claude-Sheets-Regeln |

Die `build_*`-Skripte geben JSON aus: Edits für die `applyEdits`-Schnittstelle des
Claude-Docs-Connectors. Sie wurden in dieser Reihenfolge angewendet:

```bash
python3 -I tools/replace_names.py vorlage/20270826_EMPTY_TEMPLATE.xlsx vorlage/20270826_namen_ersetzt.xlsx
```

1. `20270826_namen_ersetzt.xlsx` als Datei ins Sheet importieren, Locale `de-AT`, Zeitzone `Europe/Vienna`.
2. Zirkelbezug auflösen (siehe unten).
3. `extract_controls.py` → `build_controls.py`
4. `build_groups.py`
5. `image_tables.py`
6. `build_hinweise.py`
7. `extract_rules.py` → `build_rules.py`

## Was anders ist als in Excel

- **Zirkelbezug Produzent\*innenhonorar:** Excel rechnet das Honorar iterativ aus Kosten,
  die das Honorar enthalten. Im Sheet steht in `BERECHNUNGSSÄTZE!C30` die NFK ohne Honorar;
  `C27` und `Kalkulation!F264` lösen die ÖFI-Staffel bzw. die 5-%-Regel direkt nach dem
  Fixpunkt auf. Getestet gegen die Excel-Formeln mit 12 Werten von 0 bis 6 Mio. €: Abweichung 0.
- **Checkboxen/Auswahllisten** sind Zell-Steuerelemente. Fünf ILV-Häkchen, die es in Excel
  doppelt gab (Stammblatt und Kalkulation), werden im Stammblatt gesetzt; die Kalkulation zeigt ☑/☐.
  „Originalton I/II“ (Kalkulation B82) ist eine Auswahlliste.
- **Prüfregeln:** 24 Claude-Sheets-Regeln statt 499 Excel-Regeln. Eingabegrenzen sperren nicht
  mehr, sondern markieren. Hilfsspalten: Kalkulation `BF` und `AW` (obere Zeile je Position),
  DarstellerInnen `AE`, Stammblatt `N33:N35`.
- **Bilder:** Diäten- und Richtsatztabellen abgetippt, Logos entfallen.
- **Kein Blattschutz** in Claude Sheets.

Alle Abweichungen stehen auch im Blatt „Hinweise“ des Sheets (Art = Umbau).

# Interne Kalkulation als Google-Tabelle

Neu aufgebautes internes Werkzeug (kein Excel): eine Positionstabelle statt 276 Spalten,
Sätze mit Gültig-ab-Datum und Förderstelle, Prüfungen im Klartext mit roter Zeilenmarkierung,
ein Finanzierungsplan und ein Blatt für den Übertrag ins Antragsformular. Befüllt wird eine
per Link bearbeitbare Google-Tabelle direkt im Browser.

Die Förderstelle (Regelwerk) wird zentral im Blatt „Projekt“ gewählt: **FISA+** (Standard)
oder **ÖFI**. Sie bestimmt Produzent\*innenhonorar, FGK, Prüfungen und Anreizförderung.

| Blatt | Inhalt |
| --- | --- |
| Projekt | Titel, Förderstelle, Format (Film/Serie), Folgen, Filmart, Werkstatt/BMWKMS/Debut, Green Bonus, Gender Gap Financing, Bundesland, Stichtag, Überschreitungsreserve, Bond, Finanzierung; darunter die zum Stichtag und zur Förderstelle gültigen Sätze |
| Positionen | alle 177 ÖFI-Positionen; je Zeile Partner (AT/A/B/C), Menge, Preis, 60-h-Wochen, ILV, M/W, Plan, LNK, Anteil förderbar, Prüfung, Kostenstand. Gr., Art und KV-Funktion ergeben sich aus der Nr.; Zeilen mit Prüfhinweis sind rot |
| Zusammenfassung | Anzahl Prüfhinweise; Kostengruppen je Partner; FK, Honorar, FGK, Bond, Finanzierung, Überschreitungsreserve, HK (gesamt und förderbar); Anreizförderung; Prüfungen der Förderstelle; ÖFI-Honorar ohne Zirkelbezug, Richtsätze; LNK-Aufteilung; Gender-Budgeting |
| Finanzierung | Quellen mit Partner, Betrag, Status (geplant/beantragt/zugesagt/abgelehnt) und Anteil an der HK; Anreizförderung automatisch; Prüfung von Lücke, Eigenanteil und Mediendienstanbieter-Anteil |
| Übertrag Antrag | Summen je Antragsposition in Formular-Reihenfolge (Wochen und Ø Gage für 40 h/60 h, Ist, noch zu erwarten je Partner); bei FISA+ Honorar und FGK als Pkt. 14/15 |
| Sätze | Schlüssel, Bezeichnung, Förderstelle, Bundesland, Wert, gültig ab, Quelle (SV, DB, DZ je Bundesland, KommSt, MV, U-Bahn-Abgabe nur Wien, Diäten, FISA+- und ÖFI-Grenzen) |
| KV | KV-Mindestgagen je Funktion mit gültig ab; 40 h/60 h/Tag inkl. SZ/UEL |

Erzeugen und befüllen (`google/eingaben.json` enthält Eingaben, die beim Neuaufbau erhalten bleiben):

```bash
python3 -I google/extract_positions.py vorlage/20270826_namen_ersetzt.xlsx > google/positions.json
```

```bash
python3 -I google/build_sheets.py vorlage/20270826_EMPTY_TEMPLATE.xlsx google/positions.json google/eingaben.json > google/sheets.json
```

```bash
GSHEET_ID=<ID der Tabelle> node google/fill.js google/sheets.json create paste:Projekt paste:Sätze paste:KV paste:Positionen paste:Zusammenfassung paste:Finanzierung "paste:Übertrag Antrag"
```

```bash
GSHEET_ID=<ID der Tabelle> node google/fill.js google/sheets.json ops:Projekt ops:Positionen ops:Zusammenfassung ops:Finanzierung "ops:Übertrag Antrag" ops:Sätze ops:KV
```

Weitere Schritte von `fill.js`: `clear:<Blatt>`, `plain:'<Blatt>'!<Bereich>` (Zahlenformat
„Automatisch“, falls Formeln als Text gespeichert wurden), `cells:<Blatt>:<Zellen>`,
`delete:<Blatt>`, `add:<Blatt>`. Spaltenbreiten setzt `google/widths.js`.

Voraussetzungen: Die Tabelle ist für „Jeder, der den Link hat“ als Mitbearbeiter freigegeben,
das Gebietsschema ist deutsch (Formeln mit `;`, Dezimalkomma), und die Umgebung darf
`docs.google.com`, `0.docs.google.com` … `9.docs.google.com` sowie die Google-Static-Hosts erreichen.

Geprüft: Lohnnebenkosten (Stab und Cast), Produzent\*innenhonorar (ÖFI und FISA+), FGK,
Überschreitungsreserve, Richtsätze, Überzahlungen und Anreizförderung stimmen mit einer
unabhängigen Nachrechnung überein.
