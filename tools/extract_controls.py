"""Liest alle Formular-Steuerelemente (Checkboxen, Auswahllisten) aus einer .xlsx.

Ausgabe (JSON): je Steuerelement Blatt, Typ, Position (von/bis-Zelle), verknüpfte Zelle,
Listenbereich, aktueller Zustand und Beschriftung.

Aufruf:
    python3 -I tools/extract_controls.py VORLAGE.xlsx > controls.json
"""

import html
import json
import re
import sys
import zipfile

from openpyxl.utils import get_column_letter


def rels_of(zf, path):
    folder, name = path.rsplit("/", 1)
    rel_path = f"{folder}/_rels/{name}.rels"
    if rel_path not in zf.namelist():
        return {}
    xml = zf.read(rel_path).decode("utf-8")
    out = {}
    for m in re.finditer(r"<Relationship ([^>]*)/>", xml):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        target = attrs["Target"]
        if target.startswith("/"):
            target = target.lstrip("/")
        else:
            parts = folder.split("/")
            for seg in target.split("/"):
                if seg == "..":
                    parts.pop()
                else:
                    parts.append(seg)
            target = "/".join(parts)
        out[attrs["Id"]] = target
    return out


def main(src):
    zf = zipfile.ZipFile(src)
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    wb_rels = rels_of(zf, "xl/workbook.xml")
    result = []
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]*)"[^>]*r:id="([^"]*)"', wb):
        sheet = html.unescape(name)
        path = wb_rels[rid]
        xml = zf.read(path).decode("utf-8")
        rels = rels_of(zf, path)
        # VML-Beschriftungen nach shapeId
        labels = {}
        for target in rels.values():
            if target.endswith(".vml"):
                vml = zf.read(target).decode("utf-8", "replace")
                for shp in re.finditer(r'<v:shape id="_x0000_s(\d+)"(.*?)</v:shape>', vml, re.S):
                    box = re.search(r"<v:textbox[^>]*>(.*?)</v:textbox>", shp.group(2), re.S)
                    text = re.sub(r"<[^>]+>", "", box.group(1)) if box else ""
                    text = html.unescape(re.sub(r"\s+", " ", text)).strip()
                    labels[shp.group(1)] = text
        for m in re.finditer(r"<control ([^>]*)>(.*?)</control>", xml, re.S):
            attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
            body = m.group(2)
            frm = re.search(r"<from><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff></from>", body)
            to = re.search(r"<to><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff><xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff></to>", body)
            ctrl_path = rels[attrs["r:id"]]
            cp = zf.read(ctrl_path).decode("utf-8")
            cpa = dict(re.findall(r'(\w+)="([^"]*)"', cp))
            fc, fco, fr, fro = map(int, frm.groups())
            tc, tco, tr, tro = map(int, to.groups())
            result.append({
                "sheet": sheet,
                "name": attrs.get("name"),
                "type": cpa.get("objectType"),
                "from": f"{get_column_letter(fc + 1)}{fr + 1}",
                "fromOffsetPx": [round(fco / 9525), round(fro / 9525)],
                "to": f"{get_column_letter(tc + 1)}{tr + 1}",
                "toOffsetPx": [round(tco / 9525), round(tro / 9525)],
                "link": html.unescape(cpa.get("fmlaLink", "")) or None,
                "listRange": html.unescape(cpa.get("fmlaRange", "")) or None,
                "checked": cpa.get("checked"),
                "sel": cpa.get("sel"),
                "label": labels.get(attrs.get("shapeId"), ""),
            })
    json.dump(result, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(sys.argv[1])
