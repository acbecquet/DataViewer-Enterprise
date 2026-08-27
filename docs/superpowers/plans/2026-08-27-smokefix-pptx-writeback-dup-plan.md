# Smoke-Fix Batch: PPTX repair prompt, Excel write-back corruption, duplicate file rows

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix the three product bugs the v2.10.6 smoke surfaced - (W3/W4) every generated PPTX report triggers PowerPoint's repair prompt and the sensory report crashes the repairer, (W1) the openpyxl Excel write-back destroys formula-rich workbooks (strips every cached formula value + drops modern parts, then the app's own reader silently fabricates data from the wreck), (W2) a reopen-from-disk save mints a duplicate `files` row instead of adopting the existing one - all with permanent closed-loop harnesses.

**Architecture:** PPTX fixes are schema-order corrections in `PptxWriter` + control-char stripping in `XmlBuilder`, gated by a new bespoke OOXML semantic validator (Python, stdlib-only) run from a new test suite. The write-back fix replaces the three openpyxl load/save scripts with zip-surgery scripts that rewrite ONLY the touched sheet parts and copy every other part through byte-identical (caches, metadata.xml, webextensions all survive), plus a reader-side poka-yoke that detects cache-stripped formula cells and refuses to fabricate. The dup fix adds path-adoption at the single choke point every save goes through (`persistFileCore`).

**Tech Stack:** C++17/Qt 6.10 (qmake, `-Werror`), Python 3 stdlib (`zipfile`, `xml.etree`) for scripts and validators, openpyxl only for test fixture verification, Qt Test.

**Root-cause evidence (all confirmed on 2026-08-27, see session):**
- `makeTableCell` emits `<a:solidFill>` BEFORE `<a:lnL/R/T/B>` inside `<a:tcPr>`; ECMA-376 `CT_TableCellProperties` is an `xsd:sequence` requiring lines first, fill after. Every report contains a table, so every deck violates the schema -> chronic repair prompt. A real production deck (`second report.pptx`, fingerprint-matched) passes every STRUCTURAL check (zip, rels, content types, well-formedness), so schema order is the remaining class.
- `XmlBuilder::escapeXml` passes XML-1.0-illegal control chars (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F) through; tester-entered notes text with such a char produces fatally invalid XML - the standing hypothesis for the sensory-deck repairer crash.
- `excelWriteCellsScript` / `excelDeleteRowScript` (`src/utils/ExcelWritePayload.cpp`) and `kWriteHeaders` (`src/MainWindow.cpp:2077`) all do openpyxl `load_workbook` + `wb.save` = full package rewrite: every cached formula value is dropped (openpyxl never computes), and parts openpyxl does not model (`xl/metadata.xml`, webextensions, calcChain) are dropped -> Excel repair prompt on the workbook + `data_only=True` reads None forever after. Confirmed against the destroyed `Sunday D1320` workbook (all formula caches None) vs its pristine sibling.
- `SheetProcessors.cpp:148-168` then FABRICATES puffs (`prev + guessed interval`, where the guess computed a diff across an unfilled zero row = the observed +30 chains); recomputed TPM from None weights = 0.0. Present since v0.1.0-alpha.
- The reader (`src/ExcelReader.cpp`) tries COM first; a repair-needed workbook makes COM `Workbooks.Open` throw, landing exactly these files on the poisoned openpyxl fallback.
- Duplicate rows: `files` unique key is `(file_path, added_at)`; `persistFileCore` branch (b) INSERTs whenever `result.id == -1` with no path lookup, so any save whose model never learned its id mints a new row (observed: rows 2+3, identical path, 2.5 min apart).

**Ground rules for every task:**
- Repo is PUBLIC: no real workbooks, no real data in fixtures. Synthetic only.
- The corpus byte-identity gates (`tst_v3shadow`, `tst_v3roundtrip` over `DVE_TEST_CORPUS_DIR`) must stay green: parse behavior may only change for files that are DETECTED as cache-stripped (no corpus file is).
- Test env: Qt/MinGW on PATH (`export PATH="/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH"`), run suites with `-o file.txt,txt` (exit code == failure count). Build a suite: `cd tests/<suite> && qmake && mingw32-make -j8` (or drive via `tests/tests.pro`).
- Python fixture/validator scripts are committed test assets; write them with the Write tool (does not MIP-label). Before any C++ build: `python tools/decrypt_via_copy.py --apply`.
- Commit after every green step; message prefix `fix(pptx):`, `fix(excel):`, `fix(db):` per workstream. No co-author trailers.

---

### Task 1: PPTX semantic validator harness (RED)

**Files:**
- Create: `tests/tst_pptxvalidate/tst_pptxvalidate.pro`
- Create: `tests/tst_pptxvalidate/tst_pptxvalidate.cpp`
- Create: `tests/tst_pptxvalidate/validate_pptx.py`
- Modify: `tests/tests.pro` (add `tst_pptxvalidate` to `SUBDIRS`, alphabetical position)

- [ ] **Step 1: Write the validator script** `tests/tst_pptxvalidate/validate_pptx.py`. Stdlib only. Exit 0 = clean, exit 1 = violations (one per line on stdout, prefixed `VIOLATION:`).

```python
"""Bespoke OOXML semantic validator for DataViewer-generated .pptx decks.

Checks the constraint classes PowerPoint's loader enforces that plain
well-formedness scans cannot see: child-element ORDER (xsd:sequence), shape-id
uniqueness, table geometry, plus the structural basics (zip headers, rels
integrity, content-type coverage, media magic, control chars).

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
    blob = open(path, "rb").read()

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
            # shape-id uniqueness (group root id=1 excluded once)
            ids = [e.get("id") for e in root.iter(P + "cNvPr")]
            dups = {i for i in ids if ids.count(i) > 1}
            if dups:
                out.append(f"{n}: duplicate cNvPr ids {sorted(dups)}")
            # table geometry: tc count per tr == gridCol count
            for tbl in root.iter(A + "tbl"):
                ncols = len(tbl.find(A + "tblGrid") or [])
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
```

- [ ] **Step 2: Write the C++ harness** `tests/tst_pptxvalidate/tst_pptxvalidate.cpp`. It builds a representative deck straight through `PptxWriter` (cover + content slide with a table containing multi-line text, the five XML-special chars, and a control char), saves to a `QTemporaryDir`, and runs the validator via `QProcess`. Mirror `tst_reportgenerator`'s fixture usage for a second, integration-level deck via `ReportGenerator::generateFullReport(m_formatE, config)` (copy that suite's `m_formatE` setup verbatim). Resolve python as plain `"python"` from PATH and `QSKIP` when `QProcess` fails to start it (same policy as the corpus harnesses' python-absent skip).

Key slot (deck built at PptxWriter level):

```cpp
void TestPptxValidate::pptxWriterDeck_passesSemanticValidation()
{
    DVE::PptxWriter w;
    w.setResourcePath(QStringLiteral(SRCDIR) + "/../../resources");  // match tst_reportgenerator's resource resolution
    w.addCoverSlide(QStringLiteral("Validator Deck"), QStringLiteral("2026-08-27"));

    DVE::SlideTable t;
    t.x = 0.5; t.y = 1.0; t.w = 9.0; t.h = 4.0;
    t.headers = { "Puffs", "Notes & <specials>" };
    t.rows = { { "10", QStringLiteral("line1\nline2 \"quoted\" '&'") },
               { "20", QStringLiteral("ctrl:\x01:char") } };
    w.addContentSlide(QStringLiteral("Sheet A"), t, {}, QString());

    const QString deck = m_tmp.path() + "/pptxwriter_deck.pptx";
    QVERIFY2(w.save(deck), qPrintable(w.lastError()));
    runValidator({ deck });   // helper: QProcess python validate_pptx.py <decks>, QVERIFY2(exit==0, output)
}
```

(`runValidator` prints the validator's stdout into the failure message so the RED run enumerates the real defects. Exact `SlideTable` member names: confirm against `src/reporting/PptxWriter.h` before compiling - the suite `tst_reportgenerator` uses them already.)

- [ ] **Step 3: Register the suite** in `tests/tests.pro` `SUBDIRS`, `.pro` file mirrors a neighbor suite (`tst_reportgenerator/tst_reportgenerator.pro`) with `SOURCES += tst_pptxvalidate.cpp` plus the reporting/utils sources that suite links.

- [ ] **Step 4: Build + run, verify RED for the RIGHT reason.**

Run: `cd tests/tst_pptxvalidate && qmake && mingw32-make -j8 && ./release/tst_pptxvalidate.exe -o out.txt,txt; cat out.txt`
Expected: FAIL with `VIOLATION: ... <tcPr> child <lnL> out of schema order` (solidFill emitted first) and `VIOLATION: ... control character` from the `\x01` fixture cell. If additional violations print, record them - they are Task 2's worklist.

- [ ] **Step 5: Commit the RED harness** (`git add tests/tst_pptxvalidate tests/tests.pro && git commit -m "test(pptx): semantic OOXML validator harness - RED on tcPr order + control chars"`).

### Task 2: Fix the schema violations (GREEN)

**Files:**
- Modify: `src/reporting/PptxWriter.cpp:1028-1043` (`makeTableCell` tail)
- Modify: any exact-string assertions in `tests/tst_pptxwriter/tst_pptxwriter.cpp` that encode the old order

- [ ] **Step 1: Reorder `tcPr` children** - lines first, fill after:

```cpp
    return QStringLiteral(
        R"(<a:tc>)"
        R"(<a:txBody>)"
        R"(<a:bodyPr wrap="square"/>)"
        R"(<a:lstStyle/>)"
        R"(%1)"
        R"(</a:txBody>)"
        R"(<a:tcPr anchor="ctr">)"
        R"(<a:lnL w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:lnL>)"
        R"(<a:lnR w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:lnR>)"
        R"(<a:lnT w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:lnT>)"
        R"(<a:lnB w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:lnB>)"
        R"(<a:solidFill><a:srgbClr val="%2"/></a:solidFill>)"
        R"(</a:tcPr>)"
        R"(</a:tc>)")
        .arg(paras, bgColorHex);
```

- [ ] **Step 2: Re-run the harness.** Fix every remaining `VIOLATION:` it prints the same way (each is a precise pointer: part + element + rule). Known candidates from the code read, if reported: duplicate `cNvPr` ids across shapes on one slide (fix = single per-slide id counter in the slide builders), unexpected children the ORDERS table flags. Do NOT touch the control-char violation yet (Task 3 owns it) - temporarily drop the `\x01` row from the Task 1 fixture if it is the only red left, then restore it in Task 3.
- [ ] **Step 3: Re-run `tst_pptxwriter`, `tst_reportgenerator`, `tst_reportlayout`** - update any assertions that byte-match the old cell XML.
- [ ] **Step 4: Commit** (`fix(pptx): emit tcPr children in schema order - clears the chronic PowerPoint repair prompt`).

### Task 3: Control-char stripping + sensory deck coverage

**Files:**
- Modify: `src/utils/XmlBuilder.cpp:72-87` (`escapeXml`)
- Test: `tests/tst_xmlbuilder/tst_xmlbuilder.cpp`
- Modify: `tests/tst_pptxvalidate/tst_pptxvalidate.cpp` (restore the `\x01` fixture; add a sensory-shaped deck)

- [ ] **Step 1: Failing unit test** in `tst_xmlbuilder`:

```cpp
void TestXmlBuilder::escapeXml_stripsIllegalControlChars()
{
    QString in = QStringLiteral("a") + QChar(0x01) + QStringLiteral("b\t\n c")
               + QChar(0x0B) + QChar(0x1F) + QStringLiteral("d");
    const QString out = DVE::XmlBuilder::escapeXml(in);
    QCOMPARE(out, QStringLiteral("ab\t\n cd"));   // tab/newline survive, C0 controls do not
}
```

Run: expect FAIL (controls pass through today).

- [ ] **Step 2: Implement** - extend the switch's default branch:

```cpp
        default:
            // XML 1.0 legal chars: #x9 #xA #xD, #x20-#xD7FF, #xE000-#xFFFD.
            // Anything else (C0 controls, #xFFFE/#xFFFF) is fatal to every
            // conforming parser - PowerPoint refuses the whole part - so strip
            // rather than emit. Surrogate halves pass through: Qt stores
            // non-BMP text as valid pairs and QString iteration keeps them
            // adjacent.
            if (const char16_t u = c.unicode();
                u == 0x9 || u == 0xA || u == 0xD
                || (u >= 0x20 && u <= 0xFFFD && u != 0xFFFE)) {
                out += c;
            }
            break;
```

Run the unit test: PASS. Run the whole `tst_xmlbuilder`: green.

- [ ] **Step 3: Restore/extend the harness fixture** - the `\x01` table cell from Task 1 must now produce a VALID deck (validator green). Add a sensory-shaped deck: read `src/reporting/ReportGenerator.h` for the sensory entry point and `tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp` for how a `SensorySession` fixture is built; generate that deck with a note containing `\x01` and a newline, validate it. Radar-chart images render through `PlotEngine` - if the test environment cannot render (offscreen), pass `-platform offscreen` in the suite's `.pro` test config as the plotting suites already do (check `tests/tst_plotengine` for the pattern).
- [ ] **Step 4: Full harness green; commit** (`fix(pptx): escapeXml strips XML-illegal control chars - sensory decks can no longer emit fatal XML`).

### Task 4: Duplicate file rows - choke-point adoption

**Files:**
- Test: `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp` (new scenario; reuse `makeFileResult(filePath, marker)` at line 85)
- Modify: `src/database/DatabaseOps.cpp` (`persistFileCore`, the branch selection around line 212-222)

- [ ] **Step 1: Failing E2E** (exact save API: whatever `scenario17` calls - `m_db->saveFile(...)` - copy its call shape):

```cpp
// W2 (2026-08-27 smoke): a save whose model never learned its id (fresh parse
// of an already-saved path - reopen-from-disk, recovery restore, ...) must
// ADOPT the existing files row, never mint a second one. files' unique key is
// (file_path, added_at), so nothing at the schema level stops the dup.
void TestSaveIntegrityE2E::scenario26_reopenFromDiskAdoptsExistingFileRow()
{
    const QString path = QStringLiteral("C:/e2e/scenario26/reopen.xlsx");
    DVE::FileResult first = makeFileResult(path, QStringLiteral("s26a"));
    QVERIFY(saveWholeFile(first));          // helper name per scenario17

    DVE::FileResult again = makeFileResult(path, QStringLiteral("s26b"));
    QCOMPARE(again.id, qint64(-1));         // fresh model, id unknown - the bug's trigger
    QVERIFY(saveWholeFile(again));

    QSqlQuery q(db());
    q.prepare("SELECT count(*) FROM files WHERE file_path = ?");
    q.addBindValue(path);
    QVERIFY(q.exec() && q.next());
    QCOMPARE(q.value(0).toInt(), 1);        // RED today: 2
    QVERIFY(again.id == first.id);          // adopted, not re-minted
}
```

Run: expect FAIL with count 2. (Cleanup: add the path to the suite's wipe helper so residue never trips the e2e gates - follow how scenario17's fixture rows are wiped.)

- [ ] **Step 2: Implement adoption** in `persistFileCore`, immediately before the `if (result.id != -1 && result.version > 0)` branch:

```cpp
    // W2 (2026-08-27): a fresh struct for an ALREADY-SAVED path must adopt the
    // newest existing row instead of inserting a sibling - files' unique key
    // is (file_path, added_at), so the INSERT below would happily mint a
    // duplicate. Newest row by (added_at, id), matching loadFileByPath's
    // resolution rule, so every load path and every save path converge on the
    // same row. A genuinely new path finds nothing and INSERTs as before.
    if (result.id == -1) {
        QSqlQuery adopt(db);
        adopt.prepare(
            "SELECT id, version FROM files WHERE file_path = ? "
            "ORDER BY added_at DESC, id DESC LIMIT 1");
        adopt.addBindValue(result.filePath);
        if (adopt.exec() && adopt.next()) {
            result.id      = adopt.value(0).toLongLong();
            result.version = adopt.value(1).toInt();
        }
    }
```

The existing branch (a) then runs (it re-reads the current committed version inside the transaction, so the adopted `version` only needs to be `> 0` - which any committed row satisfies).

- [ ] **Step 3: GREEN + regression.** Run `tst_saveintegrity_e2e` fully, then `tst_databasemanager` (its destructive/OCC slots exercise the same branch selection). All green.
- [ ] **Step 4: Commit** (`fix(db): adopt the existing files row by path on id-less saves - no more duplicate file rows on reopen`).

### Task 5: Surgical Excel write-back + reader poka-yoke

**Files:**
- Modify: `src/utils/ExcelWritePayload.cpp` (both scripts replaced; argv contracts UNCHANGED)
- Modify: `src/MainWindow.cpp:2077-2103` (`kWriteHeaders` replaced; JSON contract UNCHANGED)
- Modify: `src/ExcelReader.cpp` (`read_with_openpyxl` reports stripped-formula counts), `src/ExcelReader.h` (per-sheet count in the sheet struct)
- Modify: `src/pipeline/SheetProcessors.cpp:148-168` (gate the fabrication), `src/pipeline/ReportData.h` (transient flag - NOT serialized), `src/pipeline/DataProcessor.cpp` (thread the flag)
- Modify: `src/MainWindow.cpp` (warning dialog + DB-save block for poisoned files)
- Create: `tests/tst_excelsurgery/` (`.pro`, `.cpp`, `make_formula_fixture.py`, `check_surgery.py`), register in `tests/tests.pro`

**Design (locked):** the new scripts open the source xlsx as a zip, parse ONLY the parts they must touch (`xl/workbook.xml`, its `.rels`, the one target sheet part, `[Content_Types].xml`), copy every other entry through with content byte-identical, and write tmp + `os.replace` (keeping the v2.0.2 atomic tail). Touched-cell semantics match today: numeric-looking values become numbers, blanks clear the cell, everything else is text (written as `t="inlineStr"`); an existing `<f>` on an edited cell is removed (user override); the cell's `s=` style attribute is preserved. `xl/calcChain.xml` is dropped (plus its content-type override and workbook.xml.rels entry) and `<calcPr ... fullCalcOnLoad="1"/>` is ensured in workbook.xml so Excel recalculates dependents on next open. Untouched formula cells keep `<f>` AND their cached `<v>` - the entire point. Delete-row surgically removes the `<row>` and renumbers `r=` attributes of later rows and their cells (formula TEXT is left untouched - identical to openpyxl's existing behavior, minus the destruction).

- [ ] **Step 1: Fixture generator** `tests/tst_excelsurgery/make_formula_fixture.py` - builds `fixture.xlsx` from raw parts (total control over caches; openpyxl cannot write cached values):

```python
"""Builds a minimal formula-rich xlsx with CACHED values plus a foreign part
(xl/metadata.xml stand-in) so the surgery tests can assert preservation.
Usage: python make_formula_fixture.py out.xlsx"""
import sys
import zipfile

SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>10</v></c><c r="B1"><v>1.5</v></c></row>
<row r="2"><c r="A2"><f>A1+10</f><v>20</v></c><c r="B2"><f>B1*2</f><v>3</v></c></row>
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

with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT)
    z.writestr("_rels/.rels", ROOT_RELS)
    z.writestr("xl/workbook.xml", WORKBOOK)
    z.writestr("xl/_rels/workbook.xml.rels", WB_RELS)
    z.writestr("xl/worksheets/sheet1.xml", SHEET)
    z.writestr("xl/calcChain.xml", CALCCHAIN)
    z.writestr("xl/metadata.xml", FOREIGN)
print("OK")
```

(Note: the foreign `xl/metadata.xml` is deliberately NOT in [Content_Types] - real modern workbooks cover it, but the surgery must pass unknown entries through regardless; openpyxl in the checker never opens it.)

- [ ] **Step 2: Checker script** `tests/tst_excelsurgery/check_surgery.py` - asserts post-surgery invariants, exits nonzero with a message on any failure:

```python
"""Post-surgery assertions on a fixture workbook. Usage:
python check_surgery.py edited.xlsx  -> prints OK or raises."""
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
z = zipfile.ZipFile(sys.argv[1])
names = z.namelist()

assert "xl/metadata.xml" in names, "foreign part dropped"
assert z.read("xl/metadata.xml") == b"<foreign>must survive byte-identical</foreign>", \
    "foreign part mutated"
assert "xl/calcChain.xml" not in names, "calcChain must be dropped after an edit"
assert "calcChain" not in z.read("xl/_rels/workbook.xml.rels").decode(), \
    "dangling calcChain relationship"
assert "calcChain" not in z.read("[Content_Types].xml").decode(), \
    "dangling calcChain content-type override"
wb = z.read("xl/workbook.xml").decode()
assert 'fullCalcOnLoad="1"' in wb, "fullCalcOnLoad not set"

sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
cells = {c.get("r"): c for c in sheet.iter(NS + "c")}
# A1 was edited to 99 (test drives this): value replaced, no formula
assert cells["A1"].find(NS + "v").text == "99", "edited cell value wrong"
assert cells["A1"].find(NS + "f") is None
# A2/A3/B2 untouched: formula AND cache both survive
for ref, cache in (("A2", "20"), ("A3", "30"), ("B2", "3")):
    assert cells[ref].find(NS + "f") is not None, f"{ref} lost its formula"
    assert cells[ref].find(NS + "v").text == cache, f"{ref} lost its cached value"
# B3 (inline string) untouched
assert cells["B3"].get("t") == "inlineStr"
print("OK")
```

- [ ] **Step 3: Failing C++ test** `tests/tst_excelsurgery/tst_excelsurgery.cpp` - runs `make_formula_fixture.py` into a temp dir, invokes the CURRENT `excelWriteCellsScript()` via `QProcess` with `buildWriteCellsArgs(path, "Data", {{1, 1, "99"}})`, then runs `check_surgery.py`. Today openpyxl destroys caches, so the checker FAILS at "A2 lost its cached value". (`QSKIP` when python is absent, same policy as Task 1.) Also add slots: text-edit case (`{{3, 3, "hello world"}}` writes an inline string at C3), blank-clears case (`{{1, 1, ""}}` empties A1), delete-row case (`excelDeleteRowScript` + `buildDeleteRowArgs(path, "Data", 2)` then assert with a second checker mode: row 2 gone, old row 3 now `r="2"` with cells `A2`/`B2`, formulas' TEXT unchanged), and an atomicity case (truncate the fixture to 10 bytes, run the script, assert nonzero exit AND the file still 10 bytes).
- [ ] **Step 4: Run - RED** (`A2 lost its cached value`). Commit the RED suite (`test(excel): surgery invariants harness - RED, openpyxl write-back strips caches`).
- [ ] **Step 5: Implement the surgical `kWriteCells`** in `ExcelWritePayload.cpp` (same function name, same argv contract - `buildWriteCellsArgs` is untouched). Full replacement script:

```python
import os, re, shutil, sys, zipfile
import xml.etree.ElementTree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS)
ET.register_namespace("r", NSR)
Q = "{%s}" % NS

def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def cell_sort_key(c):
    m = re.match(r"([A-Z]+)([0-9]+)", c.get("r"))
    letters = m.group(1)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n

path, sheet_name = sys.argv[1], sys.argv[2]
edits = []
args = sys.argv[3:]
i = 0
while i + 2 < len(args):
    edits.append((int(args[i]), int(args[i + 1]), args[i + 2]))
    i += 3

zin = zipfile.ZipFile(path)

# sheet name -> part path via workbook.xml + its rels
wb_root = ET.fromstring(zin.read("xl/workbook.xml"))
rid = None
for sh in wb_root.iter(Q + "sheet"):
    if sh.get("name") == sheet_name:
        rid = sh.get("{%s}id" % NSR)
if rid is None:
    print("ERROR: sheet not found: " + sheet_name, file=sys.stderr)
    sys.exit(1)
rels_root = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
RNS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
target = None
for rel in rels_root.iter(RNS + "Relationship"):
    if rel.get("Id") == rid:
        target = rel.get("Target")
if target is None:
    print("ERROR: no rel for " + rid, file=sys.stderr)
    sys.exit(1)
sheet_part = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target

# apply edits to the ONE sheet part
sroot = ET.fromstring(zin.read(sheet_part))
sdata = sroot.find(Q + "sheetData")
rows = {int(r.get("r")): r for r in sdata.findall(Q + "row")}
for (r1, c1, val) in edits:
    if r1 not in rows:
        row = ET.Element(Q + "row", {"r": str(r1)})
        idx = sum(1 for x in sdata if int(x.get("r")) < r1)
        sdata.insert(idx, row)
        rows[r1] = row
    row = rows[r1]
    ref = col_letter(c1) + str(r1)
    cell = None
    for c in row.findall(Q + "c"):
        if c.get("r") == ref:
            cell = c
    if cell is None:
        cell = ET.Element(Q + "c", {"r": ref})
        cells = row.findall(Q + "c")
        pos = sum(1 for c in cells if cell_sort_key(c) < c1)
        row.insert(pos, cell)
    style = cell.get("s")
    cell.clear()
    cell.set("r", ref)
    if style is not None:
        cell.set("s", style)
    v = val.strip()
    if not v:
        continue                       # cleared cell: keep it empty
    try:
        num = float(v)
        el = ET.SubElement(cell, Q + "v")
        el.text = repr(num) if num != int(num) else str(int(num))
    except ValueError:
        cell.set("t", "inlineStr")
        is_el = ET.SubElement(cell, Q + "is")
        t_el = ET.SubElement(is_el, Q + "t")
        t_el.text = val
        if val != val.strip():
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
new_sheet = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(sroot)

# workbook.xml: ensure fullCalcOnLoad (dependent-formula caches are now stale)
calc = wb_root.find(Q + "calcPr")
if calc is None:
    calc = ET.Element(Q + "calcPr")
    wb_root.append(calc)               # calcPr is legal after sheets/definedNames
calc.set("fullCalcOnLoad", "1")
new_wb = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(wb_root)

# workbook rels + content types: purge calcChain references (part is dropped)
drop = set()
for rel in list(rels_root.iter(RNS + "Relationship")):
    if rel.get("Type", "").endswith("/calcChain"):
        drop.add("xl/" + rel.get("Target").lstrip("/"))
        for parent in rels_root.iter():
            if rel in list(parent):
                parent.remove(rel)
new_rels = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(rels_root)
ct_raw = zin.read("[Content_Types].xml").decode("utf-8")
ct_raw = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct_raw)

tmp = path + ".dve_tmp"
with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for info in zin.infolist():
        n = info.filename
        if n in drop or n == "xl/calcChain.xml":
            continue
        if n == sheet_part:
            zout.writestr(info, new_sheet)
        elif n == "xl/workbook.xml":
            zout.writestr(info, new_wb)
        elif n == "xl/_rels/workbook.xml.rels":
            zout.writestr(info, new_rels)
        elif n == "[Content_Types].xml":
            zout.writestr(info, ct_raw)
        else:
            zout.writestr(info, zin.read(n))
zin.close()
os.replace(tmp, path)
print("OK")
```

Wrap the whole body in `try/except` printing the error to stderr and `sys.exit(1)` BEFORE the `os.replace`, so any surprise leaves the original untouched (the atomicity test proves it). Note `ET.register_namespace("", NS)` keeps the default namespace unprefixed - required, or every element gains `ns0:`.

- [ ] **Step 6: GREEN** on the write slots; iterate on the delete-row script the same way (surgical `<row>` removal + `r=` renumbering of subsequent rows and their cell refs - regex on the serialized rows is NOT acceptable, do it in the tree). Then replace `kWriteHeaders` in `MainWindow.cpp` with the same surgical core driven by its JSON payload (share nothing textually - it is a distinct script - but keep the tmp+replace tail and `print("OK")` contract). Add a header-write slot to the suite (fixture gains the 3-row header band cells the script targets).
- [ ] **Step 7: E2E through the REAL app path:** extend the suite with a slot that drives `DVE::ExcelReader` (COM will fail on the synthetic fixture only if Excel is absent - either way the openpyxl fallback must read it) BEFORE and AFTER a surgical edit and asserts the untouched formula cells still read their cached values after. This is the direct regression test for the user's bug.
- [ ] **Step 8: Commit** (`fix(excel): surgical zip-level write-back - formula caches, foreign parts and styles survive app edits`).
- [ ] **Step 9: Reader poka-yoke.** In `src/ExcelReader.cpp`'s `read_with_openpyxl`, detect stripped caches (second pass, formulas visible):

```python
def read_with_openpyxl(path):
    """Fallback: openpyxl with data_only=True (cached formula values)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    wf = load_workbook(path)          # formulas visible; None caches detectable
    result = []
    for name in wb.sheetnames:
        ws, wsf = wb[name], wf[name]
        rows, stripped = [], 0
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0
        if max_row > 0 and max_col > 0:
            for row, frow in zip(ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col),
                                 wsf.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col)):
                vals = []
                for cell, fcell in zip(row, frow):
                    if cell.value is None and isinstance(fcell.value, str) and fcell.value.startswith("="):
                        stripped += 1
                    vals.append(to_val(cell.value))
                rows.append(vals)
        result.append({"name": name, "rows": rows, "stripped_formulas": stripped})
    return result
```

`read_with_com` sheets get `"stripped_formulas": 0` (COM computes live). C++ side: add `int strippedFormulaCells = 0;` to the reader's per-sheet struct in `src/ExcelReader.h`, populate from the JSON key (absent -> 0, so COM output parses unchanged), and thread it onto the corresponding sheet through `DataProcessor` into a new TRANSIENT field `int strippedFormulaCells = 0;` on `SheetResult` in `src/pipeline/ReportData.h`. **Do NOT serialize it in `ReportDataJson.cpp`** - the legacy referee never emits it and the shadow byte-identity gate must stay green (add a one-line comment in ReportDataJson saying exactly that).

- [ ] **Step 10: Gate the fabrication.** `SheetProcessors.cpp:148` - wrap the existing fill with the detection (behavior byte-identical for healthy files):

```cpp
    // --- fill in missing Puffs values ----------------------------------
    // (comment as today) ... 2026-08-27: when the READER has detected that this
    // sheet's formula caches were stripped (an openpyxl-saved workbook read via
    // the openpyxl fallback), the zeros are DESTROYED DATA, not sparse input -
    // extrapolating fabricates plausible garbage (the +30 chains that reached
    // the DB during the v2.10.6 smoke). Leave the zeros in place; the flag
    // makes MainWindow warn and refuse to persist.
    if (sheet.strippedFormulaCells == 0
        && !s.rows.isEmpty() && s.rows[0].puffs > 0.0) {
        ... existing body unchanged ...
    }
```

(Exact variable spelling for the sheet-level flag depends on how `processSheet` receives the sheet - read the function signature and thread accordingly.)

- [ ] **Step 11: Surface + block.** In `MainWindow::loadFile`'s completion path (find where the parsed `FileResult` lands and the status bar is set): when any sheet has `strippedFormulaCells > 0`, show a `QMessageBox::warning` - "This workbook's computed values are missing (it was last saved by a tool that strips them). Open it in Excel, let it recalculate, save, then reload. Values shown may be incomplete; saving this file to the database is disabled." - and record the file as save-blocked (member set keyed by file path; the DB-save entry points skip blocked files with a status-bar notice). Search every save entry point per the cover-every-load-path lesson: the manual save action, autosave/close-time save, and the sync indicator path.
- [ ] **Step 12: Tests for the gate:** in the surgery suite, run the ORIGINAL openpyxl-destroyed fixture (make one by running the OLD behavior: openpyxl load+save in the fixture generator with `--strip` flag) through `DVE::ExcelReader` + `DataProcessor::processFile` and assert (a) `strippedFormulaCells > 0` on the sheet, (b) puffs zeros were NOT extrapolated. Then run the corpus gates (`tst_v3shadow`, `tst_v3roundtrip` with `DVE_TEST_CORPUS_DIR="C:/Users/S1134987/Documents/Python/DataViewer Dev/DB Data"`) - green, because no corpus file trips the detector (~20 min each, expected).
- [ ] **Step 13: Commit** (`fix(excel): reader detects cache-stripped workbooks - warn + block DB save instead of fabricating puffs`).

### Task 6: Wrap - full suite, build, tracker, re-smoke

- [ ] **Step 1:** `python tools/decrypt_via_copy.py --apply`, then the FULL test suite via `tests/run-tests.ps1` equivalent (build all changed suites + run in sequence; DB suites need the container - `tests/start-test-postgres.ps1` provisioning is current). Expect green across the board.
- [ ] **Step 2:** In-tree release rebuild + `MSYS_NO_PATHCONV=1 cmd /c '.\build_installer.bat' < /dev/null` after bumping VERSION to 2.10.7 (clean rebuild per the VERSION rule; runtime DLLs already staged in the worktree release tree from the v2.10.6 build). Stage as `dist/DataViewer-setup-v2.10.7.exe` in the primary repo alongside the preserved v2.10.5.
- [ ] **Step 3:** Update `docs/sprint-tracker.html` (new workstream V3-SMOKEFIX with these tasks + outcome) and the memory topic file.
- [ ] **Step 4:** Deliver the re-smoke guide IN SESSION (no file): regenerate every report type and open in PowerPoint - NO repair prompt is the acceptance; sensory deck opens clean; edit a formula-rich workbook copy (NOT a OneDrive original - use a copy of the recovered Sunday file), reopen it in the app + Excel - values intact, no repair prompt, caches present; reopen-save a file twice - one `files` row. Then the merge question per the standing rule.

## Self-review notes

- The validator's subsequence rule uses `last = max(last, idx)` so repeated children (`a:p`, `a:tr`) don't false-positive; children not in the ORDERS list for a listed parent are flagged - if a legitimate emitted child surfaces (e.g. `a:hlinkClick` inside `rPr` is NOT checked because `rPr` has no ORDERS entry), extend the table rather than silence the rule.
- Task 5's float formatting (`repr(num)`/`int` collapse) must match what openpyxl produced for the same argv so downstream readers see identical values; the E2E slot in Step 7 is the arbiter, adjust formatting there if it surfaces a mismatch.
- If the sensory deck path in Task 3 needs DB fixtures, prefer constructing `SensorySession` structs directly (as `tst_sensoryreportsource` does) - no container dependency in the pptx suite.
- W4 confirmation is empirical-by-owner: the control-char fix + schema fixes are the evidence-backed candidates; if the owner's regenerated sensory deck STILL crashes PowerPoint, capture that exact deck file and extend the validator with whatever it reveals before touching more code.
