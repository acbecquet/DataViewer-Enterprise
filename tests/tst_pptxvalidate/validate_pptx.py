"""Bespoke OOXML semantic validator for DataViewer-generated .pptx decks.

Checks the constraint classes PowerPoint's loader enforces that plain
well-formedness scans cannot see: child-element ORDER (xsd:sequence), shape-id
uniqueness, table geometry, plus the structural basics (zip headers, rels
integrity, content-type coverage, media magic, control chars).

W3 root cause this exists to gate (2026-08-27): makeTableCell emitted
<a:solidFill> before the <a:lnL/R/T/B> elements inside <a:tcPr>, violating the
CT_TableCellProperties xsd:sequence - every deck contained a table, so every
report triggered PowerPoint's repair prompt since v0.1.0-alpha.

Usage: python validate_pptx.py deck1.pptx [deck2.pptx ...]
Exit 0 when every deck is clean; exit 1 with one "VIOLATION: ..." line each.
"""
import re
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
CTRL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")

# Schema child order (subsequence rule): emitted children must appear in this
# relative order; children not listed are violations for the parents we emit.
ORDERS = {
    A + "tcPr": [A + "lnL", A + "lnR", A + "lnT", A + "lnB", A + "lnTlToBr",
                 A + "lnBlToTr", A + "cell3D", A + "noFill", A + "solidFill",
                 A + "gradFill", A + "blipFill", A + "pattFill", A + "grpFill",
                 A + "headers"],
    P + "spPr": [A + "xfrm", A + "custGeom", A + "prstGeom", A + "noFill",
                 A + "solidFill", A + "gradFill", A + "blipFill", A + "pattFill",
                 A + "grpFill", A + "ln", A + "effectLst", A + "effectDag",
                 A + "scene3d", A + "sp3d", A + "extLst"],
    P + "pic": [P + "nvPicPr", P + "blipFill", P + "spPr", P + "style"],
    P + "sp": [P + "nvSpPr", P + "spPr", P + "style", P + "txBody"],
    P + "graphicFrame": [P + "nvGraphicFramePr", P + "xfrm", A + "graphic"],
    A + "tbl": [A + "tblPr", A + "tblGrid", A + "tr"],
    A + "tc": [A + "txBody", A + "tcPr", A + "extLst"],
    A + "txBody": [A + "bodyPr", A + "lstStyle", A + "p"],
    P + "txBody": [A + "bodyPr", A + "lstStyle", A + "p"],
    A + "p": [A + "pPr", A + "r", A + "br", A + "fld", A + "endParaRPr"],
    A + "r": [A + "rPr", A + "t"],
    P + "sld": [P + "cSld", P + "clrMapOvr", P + "transition", P + "timing"],
}
# spPr appears in both p: and a: namespaces (table cells use a:); mirror it.
ORDERS[A + "spPr"] = ORDERS[P + "spPr"]


def check_order(part, elem, out):
    order = ORDERS.get(elem.tag)
    if order is not None:
        last = -1
        for child in elem:
            if child.tag not in order:
                out.append(f"{part}: <{elem.tag.split('}')[1]}> has unexpected "
                           f"child <{child.tag.split('}')[1]}>")
                continue
            idx = order.index(child.tag)
            if idx < last:
                out.append(f"{part}: <{elem.tag.split('}')[1]}> child "
                           f"<{child.tag.split('}')[1]}> out of schema order")
            # repeated same-tag children (a:p, a:tr, a:r...) keep last, not idx+1
            last = max(last, idx)
    for child in elem:
        check_order(part, child, out)


def validate(path):
    out = []
    z = zipfile.ZipFile(path)
    names = z.namelist()
    with open(path, "rb") as f:
        blob = f.read()

    # -- zip structure ------------------------------------------------------
    if z.testzip() is not None:
        out.append(f"zip: CRC failure at {z.testzip()}")
    for info in z.infolist():
        sig, ver, flags, method = struct.unpack_from("<IHHH", blob, info.header_offset)
        if sig != 0x04034b50:
            out.append(f"zip:{info.filename}: bad local header signature")
        if flags & 0x8:
            out.append(f"zip:{info.filename}: data-descriptor flag set")
        if method not in (0, 8):
            out.append(f"zip:{info.filename}: compression method {method}")
        if "\\" in info.filename:
            out.append(f"zip:{info.filename}: backslash in entry name")

    # -- XML parts: well-formed, control chars, semantic order --------------
    for n in names:
        if not n.endswith((".xml", ".rels")):
            continue
        data = z.read(n)
        if CTRL.search(data.decode("utf-8", "replace")):
            out.append(f"{n}: XML-illegal control character")
        try:
            root = ET.fromstring(data)
        except ET.ParseError as e:
            out.append(f"{n}: malformed XML: {e}")
            continue
        check_order(n, root, out)
        if n.startswith("ppt/slides/") and not n.endswith(".rels"):
            # shape-id uniqueness
            ids = [e.get("id") for e in root.iter(P + "cNvPr")]
            dups = {i for i in ids if ids.count(i) > 1}
            if dups:
                out.append(f"{n}: duplicate cNvPr ids {sorted(dups)}")
            # table geometry: tc count per tr == gridCol count
            for tbl in root.iter(A + "tbl"):
                grid = tbl.find(A + "tblGrid")
                ncols = len(grid) if grid is not None else 0
                for k, tr in enumerate(tbl.iter(A + "tr")):
                    ntc = len([c for c in tr if c.tag == A + "tc"])
                    if ntc != ncols:
                        out.append(f"{n}: table row {k} has {ntc} cells for "
                                   f"{ncols} gridCols")

    # -- content-type coverage ---------------------------------------------
    ct = z.read("[Content_Types].xml").decode()
    defaults = set(re.findall(r'Default Extension="([^"]+)"', ct))
    overrides = set(re.findall(r'Override PartName="([^"]+)"', ct))
    for n in names:
        if n == "[Content_Types].xml":
            continue
        if ("/" + n) not in overrides and n.rsplit(".", 1)[-1].lower() not in defaults:
            out.append(f"content-types: {n} uncovered")

    # -- rels integrity ----------------------------------------------------
    for n in [x for x in names if x.endswith(".rels")]:
        base = n.rsplit("_rels/", 1)[0]
        ids = {}
        for rel in ET.fromstring(z.read(n)):
            rid, tgt = rel.get("Id"), rel.get("Target")
            if rid in ids:
                out.append(f"{n}: duplicate relationship id {rid}")
            ids[rid] = tgt
            if rel.get("TargetMode") == "External":
                continue
            parts, res = (base + tgt).split("/"), []
            for p_ in parts:
                if p_ == "..":
                    res.pop()
                elif p_ != ".":
                    res.append(p_)
            if "/".join(res) not in names:
                out.append(f"{n}: {rid} targets missing part {'/'.join(res)}")
        src = base + n.split("_rels/")[-1][:-len(".rels")]
        if src in names:
            body = z.read(src).decode("utf-8", "replace")
            for rid in set(re.findall(r'r:(?:id|embed|link)="([^"]+)"', body)):
                if rid not in ids:
                    out.append(f"{src}: references {rid} absent from rels")

    # -- media sanity ------------------------------------------------------
    for n in [x for x in names if x.startswith("ppt/media/")]:
        d = z.read(n)
        if not d:
            out.append(f"{n}: zero-byte media part")
        elif n.endswith(".png") and d[:4] != b"\x89PNG":
            out.append(f"{n}: .png without PNG magic")
        elif n.endswith((".jpg", ".jpeg")) and d[:3] != b"\xff\xd8\xff":
            out.append(f"{n}: .jpg without JPEG magic")
    return out


if __name__ == "__main__":
    failed = False
    for deck in sys.argv[1:]:
        for v in validate(deck):
            print(f"VIOLATION: [{deck}] {v}")
            failed = True
    sys.exit(1 if failed else 0)
