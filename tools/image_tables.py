"""Abgetippte Tabellen, die in der Excel-Vorlage nur als Bilder vorliegen.

* Blatt "Tages- und Nächtigungsgelder": Diäten Österreich und Auslandsdiäten
  (Bilder image6-9.png, Stand September 2017).
* Blatt "Richt- und Höchstsätze": Produzent*innenpauschalen, Drehbuch- & Regie-Richtsätze
  (Bild image5.emf, Stand 01.01.2026).

Aufruf (gibt die Claude-Sheets-Edits als JSON aus):
    python3 -I tools/image_tables.py
"""

import json

STAND_DIAETEN = "Stand September 2017"
INLAND = [("Tagesgeldpauschale, Tagesdiätensatz", 26.40), ("Nächtigungsgeld, Pauschale", 15.00)]

# (Region, Land, Tagesgebühr, Nächtigungsgebühr); "- Grenzorte" mit dem Land davor ergänzt
AUSLAND_TEXT = """
I. Europa|Albanien|27,90|20,90
I. Europa|Belarus|36,80|31,00
I. Europa|Belgien|35,30|22,70
I. Europa|Brüssel|41,40|32,00
I. Europa|Bosnien und Herzegowina|31,00|23,30
I. Europa|Bulgarien|31,00|22,70
I. Europa|Dänemark|41,40|41,40
I. Europa|Deutschland|35,30|27,90
I. Europa|Deutschland – Grenzorte|30,70|18,10
I. Europa|Estland|36,80|31,00
I. Europa|Finnland|41,40|41,40
I. Europa|Frankreich|32,70|24,00
I. Europa|Paris und Straßburg|35,80|32,70
I. Europa|Griechenland|28,60|23,30
I. Europa|Großbritannien und Nordirland|36,80|36,40
I. Europa|London|41,40|41,40
I. Europa|Irland|36,80|33,10
I. Europa|Island|37,90|31,40
I. Europa|Italien|35,80|27,90
I. Europa|Rom und Mailand|40,60|36,40
I. Europa|Italien – Grenzorte|30,70|18,10
I. Europa|Jugoslawien|31,00|23,30
I. Europa|Kroatien|31,00|23,30
I. Europa|Lettland|36,80|31,00
I. Europa|Liechtenstein|30,70|18,10
I. Europa|Litauen|36,80|31,00
I. Europa|Luxemburg|35,30|22,70
I. Europa|Malta|30,10|30,10
I. Europa|Moldau|36,80|31,00
I. Europa|Niederlande|35,30|27,90
I. Europa|Norwegen|42,90|41,40
I. Europa|Polen|32,70|25,10
I. Europa|Portugal|27,90|22,70
I. Europa|Rumänien|36,80|27,30
I. Europa|Russische Föderation|36,80|31,00
I. Europa|Moskau|40,60|31,00
I. Europa|Schweden|42,90|41,40
I. Europa|Schweiz|36,80|32,70
I. Europa|Schweiz – Grenzorte|30,70|18,10
I. Europa|Slowakei|27,90|15,90
I. Europa|Preßburg|31,00|24,40
I. Europa|Slowenien|31,00|23,30
I. Europa|Slowenien – Grenzorte|27,90|15,90
I. Europa|Spanien|34,20|30,50
I. Europa|Tschechien|31,00|24,40
I. Europa|Tschechien – Grenzorte|27,90|15,90
I. Europa|Türkei|31,00|36,40
I. Europa|Ukraine|36,80|31,00
I. Europa|Ungarn|26,60|26,60
I. Europa|Budapest|31,00|26,60
I. Europa|Ungarn – Grenzorte|26,60|18,10
I. Europa|Zypern|28,60|30,50
II. Afrika|Ägypten|37,90|41,40
II. Afrika|Algerien|41,40|27,00
II. Afrika|Angola|43,60|41,40
II. Afrika|Äthiopien|37,90|41,40
II. Afrika|Benin|36,20|26,60
II. Afrika|Burkina Faso|39,20|21,10
II. Afrika|Burundi|37,90|37,90
II. Afrika|Côte d'Ivoire|39,20|32,00
II. Afrika|Demokratische Rep. Kongo|47,30|33,10
II. Afrika|Dschibuti|45,80|47,30
II. Afrika|Gabun|45,80|39,90
II. Afrika|Gambia|43,60|30,10
II. Afrika|Ghana|43,60|30,10
II. Afrika|Guinea|43,60|30,10
II. Afrika|Kamerun|45,80|25,30
II. Afrika|Kap Verde|27,90|19,60
II. Afrika|Kenia|34,90|32,00
II. Afrika|Liberia|39,20|41,40
II. Afrika|Libyen|43,60|36,40
II. Afrika|Madagaskar|36,40|36,40
II. Afrika|Malawi|32,70|32,70
II. Afrika|Mali|39,20|31,20
II. Afrika|Marokko|32,70|21,80
II. Afrika|Mauretanien|33,80|31,20
II. Afrika|Mauritius|36,40|36,40
II. Afrika|Mosambik|43,60|41,40
II. Afrika|Namibia|34,90|34,00
II. Afrika|Niger|39,20|21,10
II. Afrika|Nigeria|39,20|34,20
II. Afrika|Republik Kongo|39,20|26,80
II. Afrika|Ruanda|37,90|37,90
II. Afrika|Sambia|37,10|34,00
II. Afrika|Senegal|49,30|31,20
II. Afrika|Seychellen|36,40|36,40
II. Afrika|Sierra Leone|43,60|34,20
II. Afrika|Simbabwe|37,10|34,00
II. Afrika|Somalia|32,70|29,00
II. Afrika|Südafrika|34,90|34,00
II. Afrika|Sudan|43,60|41,40
II. Afrika|Tansania|43,60|32,00
II. Afrika|Togo|36,20|26,60
II. Afrika|Tschad|36,20|26,60
II. Afrika|Tunesien|36,20|29,20
II. Afrika|Uganda|41,40|32,00
II. Afrika|Zentralafrikanische Republik|39,20|29,00
III. Amerika|Argentinien|33,10|47,30
III. Amerika|Bahamas|48,00|30,50
III. Amerika|Barbados|51,00|43,60
III. Amerika|Bolivien|26,60|25,10
III. Amerika|Brasilien|33,10|36,40
III. Amerika|Chile|37,50|36,40
III. Amerika|Costa Rica|31,80|31,80
III. Amerika|Dominikanische Republik|39,20|43,60
III. Amerika|Ecuador|26,60|21,60
III. Amerika|El Salvador|31,80|26,20
III. Amerika|Guatemala|31,80|31,80
III. Amerika|Guyana|39,20|34,20
III. Amerika|Haiti|39,20|27,70
III. Amerika|Honduras|31,80|27,00
III. Amerika|Jamaika|47,10|47,10
III. Amerika|Kanada|41,00|34,20
III. Amerika|Kolumbien|33,10|35,10
III. Amerika|Kuba|54,10|27,70
III. Amerika|Mexiko|41,00|36,40
III. Amerika|Nicaragua|31,80|36,40
III. Amerika|Niederländische Antillen|43,60|27,70
III. Amerika|Panama|43,60|36,40
III. Amerika|Paraguay|33,10|25,10
III. Amerika|Peru|33,10|25,10
III. Amerika|Suriname|39,20|25,10
III. Amerika|Trinidad und Tobago|51,00|43,60
III. Amerika|Uruguay|33,10|25,10
III. Amerika|USA|52,30|42,90
III. Amerika|New York und Washington|65,40|51,00
III. Amerika|Venezuela|39,20|35,10
IV. Asien|Afghanistan|31,80|27,70
IV. Asien|Armenien|36,80|31,00
IV. Asien|Aserbaidschan|36,80|31,00
IV. Asien|Bahrain|54,10|37,50
IV. Asien|Bangladesch|31,80|34,20
IV. Asien|Brunei|33,10|42,10
IV. Asien|China|35,10|30,50
IV. Asien|Georgien|36,80|31,00
IV. Asien|Hongkong|46,40|37,90
IV. Asien|Indien|31,80|39,90
IV. Asien|Indonesien|39,20|32,00
IV. Asien|Irak|54,10|36,40
IV. Asien|Iran|37,10|29,00
IV. Asien|Israel|37,10|32,50
IV. Asien|Japan|65,60|42,90
IV. Asien|Jemen|54,10|37,50
IV. Asien|Jordanien|37,10|32,50
IV. Asien|Kambodscha|31,40|31,40
IV. Asien|Kasachstan|36,80|31,00
IV. Asien|Katar|54,10|37,50
IV. Asien|Kirgisistan|36,80|31,00
IV. Asien|Korea, Demokratische Volksrepublik|32,50|32,50
IV. Asien|Korea, Republik|45,30|32,50
IV. Asien|Kuwait|54,10|37,50
IV. Asien|Laos|31,40|31,40
IV. Asien|Libanon|31,80|35,10
IV. Asien|Malaysia|43,60|45,10
IV. Asien|Mongolei|29,40|29,40
IV. Asien|Myanmar|29,40|29,40
IV. Asien|Nepal|31,80|34,20
IV. Asien|Oman|54,10|37,50
IV. Asien|Pakistan|27,70|25,10
IV. Asien|Philippinen|32,50|32,50
IV. Asien|Saudi-Arabien|54,10|37,50
IV. Asien|Singapur|43,60|44,70
IV. Asien|Sri Lanka|31,80|32,70
IV. Asien|Syrien|32,70|29,00
IV. Asien|Tadschikistan|36,80|31,00
IV. Asien|Taiwan|39,20|37,50
IV. Asien|Thailand|39,20|42,10
IV. Asien|Turkmenistan|36,80|31,00
IV. Asien|Usbekistan|36,80|31,00
IV. Asien|Vereinigte Arabische Emirate|54,10|37,50
IV. Asien|Vietnam|31,40|31,40
V. Australien|Australien|47,30|39,90
V. Australien|Neuseeland|32,50|36,40
"""

# Fertigungskosten, Prod.-Honorar, FGK, Drehbuch, Regie (Stand 01.01.2026)
RICHTSAETZE = [
    (200000, 20000, 18000, 24000, 36000), (300000, 26875, 27000, 29000, 42000),
    (400000, 33750, 36000, 34000, 46000), (500000, 40625, 45000, 39000, 50000),
    (600000, 47500, 54000, 42000, 54000), (700000, 54375, 63000, 45000, 57000),
    (800000, 61250, 72000, 48000, 60000), (900000, 68125, 79500, 50000, 63000),
    (1000000, 75000, 87000, 52000, 66000), (1100000, 80000, 94500, 54000, 69000),
    (1200000, 85000, 102000, 56000, 72000), (1300000, 90000, 108500, 58000, 75000),
    (1400000, 95000, 115000, 60000, 78000), (1500000, 100000, 121500, 62000, 81000),
    (1600000, 102500, 128000, 64000, 84000), (1700000, 105000, 134611, 66000, 87000),
    (1800000, 107500, 141222, 68000, 90000), (1900000, 110000, 147833, 70000, 93000),
    (2000000, 112500, 154444, 72000, 96000), (2100000, 115000, 161056, 74000, 99000),
    (2200000, 117500, 167667, 76000, 102000), (2300000, 120000, 174278, 78000, 105000),
    (2400000, 122500, 180889, 80000, 108000), (2500000, 125000, 187500, 82000, 110000),
    (2600000, 130000, 195000, 84000, 112000), (2700000, 135000, 202500, 86000, 114000),
    (2800000, 140000, 210000, 88000, 116000), (2900000, 145000, 217500, 89000, 118000),
    (3000000, 150000, 225000, 90000, 120000), (3100000, 155000, 232500, 91000, 121000),
    (3200000, 160000, 240000, 92000, 122000), (3300000, 165000, 247500, 93000, 123000),
    (3400000, 170000, 255000, 94000, 124000), (3500000, 175000, 262500, 95000, 125000),
    (3600000, 180000, 270000, 96000, 126000), (3700000, 185000, 277500, 97000, 127000),
    (3800000, 190000, 285000, 98000, 128000), (3900000, 195000, 292500, 99000, 129000),
    (4000000, 200000, 300000, 100000, 130000), (4500000, 225000, 337500, 105000, 135000),
    (5000000, 250000, 375000, 110000, 150000), (5500000, 250000, 375000, 115000, 165000),
    (6000000, 250000, 375000, 120000, 180000), (7000000, 250000, 375000, 140000, 210000),
    (8000000, 250000, 375000, 160000, 240000), (9000000, 250000, 375000, 180000, 270000),
    (10000000, 250000, 375000, 200000, 300000),
]


def num(s):
    return float(s.replace(".", "").replace(",", "."))


def diaeten_edits():
    sheet = "Tages- und Nächtigungsgelder"
    rows = [line.split("|") for line in AUSLAND_TEXT.strip().splitlines()]
    top = 9  # Tabellenkopf
    end = top + len(rows)
    edits = [
        {"kind": "writeCells", "sheet": sheet, "topLeft": "A1", "rows": [
            ["Diäten Österreich", None, None, STAND_DIAETEN],
            [None, None, None, None],
            [INLAND[0][0], None, INLAND[0][1], None],
            [INLAND[1][0], None, INLAND[1][1], None],
            [None, None, None, None],
            ["Tages- und Nächtigungsgelder bei Auslandsdienstreisen", None, None, None],
            ["Abgetippt aus den Bildern der Excel-Vorlage.", None, None, None],
            [None, None, None, None],
            ["Region", "Land", "Tagesgebühr", "Nächtigungsgebühr"],
        ]},
        {"kind": "writeCells", "sheet": sheet, "topLeft": f"A{top + 1}",
         "rows": [[r[0], r[1], num(r[2]), num(r[3])] for r in rows]},
        {"kind": "formatCells", "sheet": sheet, "ranges": ["A1", "A6"], "format": {"font": {"bold": True, "size": 12}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": ["A7", "D1"], "format": {"font": {"italic": True, "color": {"kind": "standard", "hue": "gray", "shade": 600}}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": ["A9:D9"], "format": {"fillColor": {"kind": "rgb", "hex": "#f3f3f3"}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": ["C3:C4", f"C{top + 1}:D{end}"], "format": {"numberFormat": {"pattern": "#,##0.00 \"€\""}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": ["C9:D9"], "format": {"alignment": {"horizontal": "right"}}},
        {"kind": "setCellTable", "sheet": sheet, "cellTable": {"columns": "A:D", "headerRows": "9", "footerRows": None}},
        {"kind": "setSheetSettings", "sheet": sheet, "settings": {"lastFrozenRow": 9}},
        {"kind": "setColumnWidths", "sheet": sheet, "columns": "A", "widthInPixels": 120},
        {"kind": "setColumnWidths", "sheet": sheet, "columns": "B", "widthInPixels": 260},
        {"kind": "setColumnWidths", "sheet": sheet, "columns": "C:D", "widthInPixels": 140},
    ]
    return edits, len(rows)


def richtsatz_edits():
    sheet = "Richt- und Höchstsätze"
    top = 16  # Kopfzeile
    data = []
    for i, (fk, ph, fgk, db, rg) in enumerate(RICHTSAETZE):
        r = top + 1 + i
        data.append([fk, ph, f"=B{r}/A{r}", fgk, f"=D{r}/A{r}", db, f"=F{r}/A{r}", rg, f"=H{r}/A{r}"])
    end = top + len(RICHTSAETZE)
    edits = [
        {"kind": "writeCells", "sheet": sheet, "topLeft": "A14", "rows": [
            ["Produzent*innenpauschalen, Drehbuch- & Regie-Richtsätze (01.01.2026)"] + [None] * 8,
            ["Abgetippt aus dem Bild der Excel-Vorlage; die %-Spalten sind Formeln."] + [None] * 8,
            ["Fertigungskosten", "Prod.-Honorar", "%", "FGK", "%", "Drehbuch", "%", "Regie", "%"],
        ]},
        {"kind": "writeCells", "sheet": sheet, "topLeft": f"A{top + 1}", "rows": data},
        {"kind": "writeCells", "sheet": sheet, "topLeft": f"A{end + 2}", "rows": [
            ["Bemessungsgrundlage bei internationalen Koproduktionen:", None, None, None, None, None, None, None, None],
            ["Prod.-Honorar und FGK: österreichische Fertigungskosten", None, None, None, None, None, None, None, None],
            ["Drehbuch und Regie: gesamte Fertigungskosten", None, None, None, None, None, None, None, None],
        ]},
        {"kind": "formatCells", "sheet": sheet, "ranges": ["A14", f"A{end + 2}"], "format": {"font": {"bold": True}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": ["A15"], "format": {"font": {"italic": True, "color": {"kind": "standard", "hue": "gray", "shade": 600}}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": [f"A{top}:I{top}"], "format": {"fillColor": {"kind": "rgb", "hex": "#f3f3f3"}, "alignment": {"horizontal": "right"}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": [f"A{top + 1}:B{end}", f"D{top + 1}:D{end}", f"F{top + 1}:F{end}", f"H{top + 1}:H{end}"], "format": {"numberFormat": {"pattern": "#,##0"}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": [f"C{top + 1}:C{end}", f"E{top + 1}:E{end}", f"G{top + 1}:G{end}", f"I{top + 1}:I{end}"], "format": {"numberFormat": {"pattern": "0.00%"}}},
        {"kind": "formatCells", "sheet": sheet, "ranges": [f"F{top}:I{end}"], "format": {"fillColor": {"kind": "standard", "hue": "blue", "shade": 100}}},
    ]
    return edits


if __name__ == "__main__":
    d, n = diaeten_edits()
    print(json.dumps({"diaeten": d, "richtsaetze": richtsatz_edits()}, ensure_ascii=False, separators=(",", ":")))
