"""Liest alle bedingten Formatierungen (inkl. x14-Erweiterungen) und Datenüberprüfungen
mit Wertgrenzen aus einer .xlsx und gibt sie als JSON aus.

Aufruf:
    python3 -I tools/extract_rules.py VORLAGE.xlsx > rules.json
"""

import html
import json
import re
import sys
import zipfile


def dxf_list(zf):
    styles = zf.read("xl/styles.xml").decode("utf-8")
    m = re.search(r"<dxfs[^>]*>(.*?)</dxfs>", styles, re.S)
    out = []
    if not m:
        return out
    for d in re.findall(r"<dxf>(.*?)</dxf>", m.group(1), re.S):
        out.append(parse_dxf(d))
    return out


def parse_dxf(d):
    res = {}
    font = re.search(r"<(?:x14:)?font>(.*?)</(?:x14:)?font>", d, re.S)
    if font:
        c = re.search(r'<color ([^>]*)/>', font.group(1))
        if c:
            res["fontColor"] = dict(re.findall(r'(\w+)="([^"]*)"', c.group(1)))
        if "<b/>" in font.group(1) or "<b " in font.group(1):
            res["bold"] = True
    fill = re.search(r"<fill>(.*?)</fill>", d, re.S)
    if fill:
        c = re.search(r'<bgColor ([^>]*)/>', fill.group(1)) or re.search(r'<fgColor ([^>]*)/>', fill.group(1))
        if c:
            res["fill"] = dict(re.findall(r'(\w+)="([^"]*)"', c.group(1)))
    return res


def main(src):
    zf = zipfile.ZipFile(src)
    dxfs = dxf_list(zf)
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target = {}
    for m in re.finditer(r"<Relationship ([^>]*)/>", rels):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        target[a["Id"]] = "xl/" + a["Target"].lstrip("/").replace("xl/", "")
    out = []
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]*)"[^>]*r:id="([^"]*)"', wb):
        sheet = html.unescape(name)
        xml = zf.read(target[rid]).decode("utf-8")
        # klassische Regeln
        for m in re.finditer(r'<conditionalFormatting sqref="([^"]*)"[^>]*>(.*?)</conditionalFormatting>', xml, re.S):
            sqref = m.group(1)
            for r in re.finditer(r"<cfRule ([^>]*?)(?:/>|>(.*?)</cfRule>)", m.group(2), re.S):
                a = dict(re.findall(r'(\w+)="([^"]*)"', r.group(1)))
                body = r.group(2) or ""
                formulas = [html.unescape(f) for f in re.findall(r"<formula>(.*?)</formula>", body, re.S)]
                rule = {"sheet": sheet, "sqref": sqref, "type": a.get("type"), "operator": a.get("operator"),
                        "priority": int(a.get("priority", 0)), "stopIfTrue": a.get("stopIfTrue"),
                        "formulas": formulas, "format": dxfs[int(a["dxfId"])] if "dxfId" in a else None}
                if a.get("type") == "iconSet":
                    rule["iconSet"] = re.search(r'iconSet="?([^" >]*)', body).group(1) if 'iconSet="' in body else None
                    rule["cfvo"] = re.findall(r'<cfvo ([^>]*)/>', body)
                out.append(rule)
        # x14-Erweiterungen
        for m in re.finditer(r"<x14:conditionalFormatting[^>]*>(.*?)</x14:conditionalFormatting>", xml, re.S):
            sqref = re.search(r"<xm:sqref>(.*?)</xm:sqref>", m.group(1)).group(1)
            for r in re.finditer(r"<x14:cfRule ([^>]*)>(.*?)</x14:cfRule>", m.group(1), re.S):
                a = dict(re.findall(r'(\w+)="([^"]*)"', r.group(1)))
                body = r.group(2)
                formulas = [html.unescape(f) for f in re.findall(r"<xm:f>(.*?)</xm:f>", body, re.S)]
                dxf = re.search(r"<x14:dxf>(.*?)</x14:dxf>", body, re.S)
                rule = {"sheet": sheet, "sqref": sqref, "type": a.get("type"), "operator": a.get("operator"),
                        "priority": int(a.get("priority", 0)), "stopIfTrue": a.get("stopIfTrue"),
                        "formulas": formulas, "format": parse_dxf(dxf.group(1)) if dxf else None, "x14": True}
                if a.get("type") == "iconSet":
                    iset = re.search(r"<x14:iconSet ([^>]*)>", body)
                    rule["iconSet"] = iset.group(1) if iset else None
                    rule["cfvo"] = re.findall(r"<x14:cfvo ([^>]*)>(?:<xm:f>(.*?)</xm:f>)?", body)
                out.append(rule)
    json.dump(out, sys.stdout, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main(sys.argv[1])
