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
