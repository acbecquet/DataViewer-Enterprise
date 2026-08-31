"""Builds a minimal formula-rich xlsx with CACHED values plus a foreign part
(xl/metadata.xml stand-in) so the surgery tests can assert preservation.

openpyxl cannot WRITE cached formula values (that is the whole bug), so the
fixture is assembled from raw parts for total control over <f> + <v> pairs.

D2 reproduces Excel's cached EMPTY-STRING result (t="str" with an empty <v/>):
a healthy Excel save of an IF(...,"",...) cell looks exactly like this, and
openpyxl's data_only read returns None for it - the W3b false positive. The
stripped-cache detector must treat it as CACHED (the workbook carries other
non-empty caches), not as destroyed data.

Usage: python make_formula_fixture.py out.xlsx [--strip]
  --strip: additionally round-trip the file through openpyxl load+save, i.e.
           reproduce the OLD write-back's destruction (caches gone, foreign
           part dropped) for the reader poka-yoke tests.
"""
import sys
import zipfile

SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>1.5</v></c></row>
<row r="2"><c r="A2"><f>A1+10</f><v>20</v></c><c r="B2"><f>B1*2</f><v>3</v></c><c r="D2" t="str"><f>IF(A1=99,"x","")</f><v/></c></row>
<row r="3"><c r="A3"><f>A2+10</f><v>30</v></c><c r="B3" t="inlineStr"><is><t>note</t></is></c></row>
</sheetData>
</worksheet>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""

WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>"""

CALCCHAIN = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A2" i="1"/><c r="A3" i="1"/><c r="B2" i="1"/></calcChain>"""

FOREIGN = b"<foreign>must survive byte-identical</foreign>"

out = sys.argv[1]
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WORKBOOK)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", SHEET)
    z.writestr("xl/calcChain.xml", CALCCHAIN)
    z.writestr("xl/metadata.xml", FOREIGN)

if "--strip" in sys.argv[2:]:
    from openpyxl import load_workbook
    wb = load_workbook(out)
    wb.save(out)   # the OLD write-back's effect: caches stripped, parts dropped

print("OK")
