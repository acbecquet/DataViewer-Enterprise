"""Post-surgery assertions on the formula fixture workbook.

Usage: python check_surgery.py edited.xlsx [mode]
  (no mode)      after writing A1 -> 99: caches + foreign part survive
  --textclear    after writing C3 -> "hello world" and clearing A1
  --after-delete after deleting row 2: renumbering + caches survive

Prints OK on success; raises AssertionError (nonzero exit) with the reason
otherwise - the C++ test surfaces that text directly.
"""
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

path = sys.argv[1]
mode = sys.argv[2] if len(sys.argv) > 2 else ""

z = zipfile.ZipFile(path)
names = z.namelist()

# Foreign-part preservation + calcChain purge hold in EVERY mode.
assert "xl/metadata.xml" in names, "foreign part dropped"
assert z.read("xl/metadata.xml") == b"<foreign>must survive byte-identical</foreign>", \
    "foreign part mutated"
assert "xl/calcChain.xml" not in names, "calcChain must be dropped after an edit"
assert "calcChain" not in z.read("xl/_rels/workbook.xml.rels").decode(), \
    "dangling calcChain relationship"
assert "calcChain" not in z.read("[Content_Types].xml").decode(), \
    "dangling calcChain content-type override"
assert 'fullCalcOnLoad="1"' in z.read("xl/workbook.xml").decode(), \
    "fullCalcOnLoad not set"

sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
cells = {c.get("r"): c for c in sheet.iter(NS + "c")}


def assert_formula_cache(ref, cache):
    assert ref in cells, f"{ref} missing"
    assert cells[ref].find(NS + "f") is not None, f"{ref} lost its formula"
    v = cells[ref].find(NS + "v")
    assert v is not None and v.text == cache, f"{ref} lost its cached value"


if mode == "":
    v = cells["A1"].find(NS + "v")
    assert v is not None and v.text == "99", "edited cell value wrong"
    assert cells["A1"].find(NS + "f") is None, "edited cell kept a formula"
    for ref, cache in (("A2", "20"), ("A3", "30"), ("B2", "3")):
        assert_formula_cache(ref, cache)
    assert cells["B3"].get("t") == "inlineStr", "untouched inline string mutated"
elif mode == "--textclear":
    c3 = cells["C3"]
    assert c3.get("t") == "inlineStr", "text cell not written as inline string"
    t = c3.find(NS + "is/" + NS + "t")
    assert t is not None and t.text == "hello world", "text cell content wrong"
    a1 = cells["A1"]
    assert a1.find(NS + "v") is None and a1.find(NS + "f") is None, \
        "cleared cell still holds content"
    for ref, cache in (("A2", "20"), ("A3", "30"), ("B2", "3")):
        assert_formula_cache(ref, cache)
elif mode == "--after-delete":
    rows = [r.get("r") for r in sheet.iter(NS + "row")]
    assert rows == ["1", "2"], f"row renumbering wrong: {rows}"
    # old row 3 (A3 =A2+10 cache 30) is now row 2 with refs renumbered
    assert "A2" in cells and "A3" not in cells, "cell refs not renumbered"
    f = cells["A2"].find(NS + "f")
    assert f is not None and f.text == "A2+10", \
        f"formula text must be left untouched (parity with openpyxl): {None if f is None else f.text}"
    v = cells["A2"].find(NS + "v")
    assert v is not None and v.text == "30", "moved row lost its cached value"
    assert cells["B2"].get("t") == "inlineStr", "moved inline string mutated"
else:
    raise SystemExit(f"unknown mode {mode}")

print("OK")
