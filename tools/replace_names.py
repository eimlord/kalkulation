"""Ersetzt benannte Bereiche (Defined Names) in einer .xlsx durch direkte Zellbezüge.

Claude Sheets kennt keine benannten Bereiche: eine Formel wie =NFK*ÜR würde dort
#NAME? zeigen. Dieses Skript schreibt jede Formel im Arbeitsmappen-XML so um, dass
an Stelle des Namens der Bezug steht, auf den er zeigt (z. B. 'Zusammenfassung'!$D$17).

Bearbeitet wird nur der Formeltext im XML; alles andere (Formate, Steuerelemente,
Bilder, Erweiterungen) wird Byte für Byte übernommen.

Aufruf:
    python3 -I tools/replace_names.py EINGABE.xlsx AUSGABE.xlsx
"""

import html
import re
import sys
import zipfile

from openpyxl.formula.tokenizer import Token, Tokenizer

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def quote_sheet(ref):
    """'Blatt'!$A$1 – Blattnamen immer in Hochkommas, das ist in jedem Fall gültig."""
    if "!" not in ref:
        return ref
    sheet, cell = ref.rsplit("!", 1)
    if not (sheet.startswith("'") and sheet.endswith("'")):
        sheet = "'" + sheet.replace("'", "''") + "'"
    return f"{sheet}!{cell}"


def read_workbook(zf):
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    sheets = re.findall(r'<sheet [^>]*name="([^"]*)"[^>]*r:id="([^"]*)"', wb)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target = dict(re.findall(r'Id="([^"]*)"[^>]*Target="([^"]*)"', rels))
    target.update({k: v for v, k in re.findall(r'Target="([^"]*)"[^>]*Id="([^"]*)"', rels)})
    sheet_files = {}
    for idx, (name, rid) in enumerate(sheets):
        t = target[rid]
        path = t.lstrip("/") if t.startswith("/") else "xl/" + t
        sheet_files[path] = (html.unescape(name), idx)
    global_names, local_names = {}, {}
    for attrs, text in re.findall(r"<definedName ([^>]*)>([^<]*)</definedName>", wb):
        name = re.search(r'name="([^"]*)"', attrs).group(1)
        value = html.unescape(text)
        local = re.search(r'localSheetId="(\d+)"', attrs)
        if name.startswith("_xlnm."):
            continue
        if local:
            local_names.setdefault(int(local.group(1)), {})[name.lower()] = value
        else:
            global_names[name.lower()] = value
    return sheet_files, global_names, local_names


def rewrite_formula(text, names, stats):
    tok = Tokenizer("=" + text)
    changed = False
    for t in tok.items:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            key = t.value.lower()
            if key in names:
                target = names[key]
                if "#REF!" in target:
                    stats["broken"].append(t.value)
                    continue
                t.value = quote_sheet(target)
                stats["names"][key] = stats["names"].get(key, 0) + 1
                changed = True
    if not changed:
        return text, False
    out = "".join(t.value for t in tok.items)
    return out, True


def check_roundtrip(text):
    tok = Tokenizer("=" + text)
    return "".join(t.value for t in tok.items) == text


FORMULA_RE = re.compile(r"(<(f|formula|formula1|formula2|xm:f)(\s[^>]*)?>)([^<]*)(</\2>)")


def process_sheet(xml, names, stats):
    count = 0

    def repl(m):
        nonlocal count
        raw = html.unescape(m.group(4))
        if not raw:
            return m.group(0)
        if not check_roundtrip(raw):
            stats["roundtrip_fail"].append(raw)
            return m.group(0)
        new, changed = rewrite_formula(raw, names, stats)
        if not changed:
            return m.group(0)
        count += 1
        esc = new.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return m.group(1) + esc + m.group(5)

    return FORMULA_RE.sub(repl, xml), count


def main(src, dst):
    stats = {"names": {}, "broken": [], "roundtrip_fail": []}
    with zipfile.ZipFile(src) as zin:
        sheet_files, global_names, local_names = read_workbook(zin)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            total = 0
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in sheet_files:
                    sheet_name, idx = sheet_files[item.filename]
                    names = dict(global_names)
                    names.update(local_names.get(idx, {}))
                    xml, n = process_sheet(data.decode("utf-8"), names, stats)
                    data = xml.encode("utf-8")
                    total += n
                    if n:
                        print(f"{sheet_name:40s} {n:5d} Formeln umgeschrieben")
                zout.writestr(item, data)
    print(f"Gesamt: {total} Formeln, {sum(stats['names'].values())} Namensverwendungen ersetzt")
    if stats["broken"]:
        print("WARNUNG: Namen mit #REF! verwendet:", sorted(set(stats["broken"])))
    if stats["roundtrip_fail"]:
        print("WARNUNG: Formeln nicht verlustfrei tokenisierbar:", len(stats["roundtrip_fail"]))
        for f in stats["roundtrip_fail"][:10]:
            print("   ", f)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
