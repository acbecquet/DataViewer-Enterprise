"""Copy an xlsx fixture, converting the puffs-column LITERALS of data rows 2+
into UNCACHED formulas (<f>prev+10</f> followed by an empty <v/>).

That is the exact on-disk state of app-template-lineage workbooks (the
bundled New File template is openpyxl-born, so its formula cells carry no
caches by construction) and equally of workbooks the OLD openpyxl write-back
destroyed. Which of the two it "is" depends only on the sheet layout the
pipeline resolves - which is precisely what the W3b fork exemption decides.

Data rows start at sheet row 5; row 5 keeps its literal seed (the template
seeds row-5 puffs). Puffs live at block col offset+1: columns A (block 0)
and M (block 1).

Usage: make_stripped_variant.py in.xlsx out.xlsx
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = "{%s}" % MAIN
ET.register_namespace("", MAIN)

src, dst = sys.argv[1], sys.argv[2]

with zipfile.ZipFile(src) as zin, \
        zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", item.filename):
            root = ET.fromstring(data)
            for row in root.iter(NS + "row"):
                r = int(row.get("r"))
                if r < 6:          # rows 1-4 headers, row 5 keeps its seed
                    continue
                for c in row.findall(NS + "c"):
                    ref = c.get("r") or ""
                    m = re.match(r"([A-Z]+)\d+", ref)
                    if not m or m.group(1) not in ("A", "M"):
                        continue   # puffs columns only (block 0 / block 1)
                    v = c.find(NS + "v")
                    if v is None:
                        continue
                    c.remove(v)
                    f = ET.SubElement(c, NS + "f")
                    f.text = "%s%d+10" % (m.group(1), r - 1)
                    ET.SubElement(c, NS + "v")   # EMPTY <v/> - stripped shape
            data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        zout.writestr(item, data)

print("OK")
