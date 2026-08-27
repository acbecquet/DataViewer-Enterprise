#include "utils/ExcelWritePayload.h"

namespace DVE {

// ─── The write-back scripts (single source of truth) ─────────────────────────
//
// These were inline `static const char*` literals in MainWindow::writeCellsToExcel
// and ::deleteRowFromExcel before SP3-T4. Moving them here means the synchronous
// path, the off-thread worker, and the equivalence test all reference the SAME
// bytes — there is no second copy to drift out of sync. The leading/trailing
// newlines of the raw string are part of the payload and MUST NOT change.
//
// W1 (2026-08-27 smoke): both scripts were openpyxl load_workbook + wb.save —
// a FULL package rewrite. openpyxl never computes formulas, so every cached
// formula value in the workbook was stripped on every edit, and parts openpyxl
// does not model (xl/metadata.xml, webextensions, calcChain) were dropped.
// Excel then demanded repair on the workbook, and the app's own data_only
// reader saw None for every formula cell forever after — SheetProcessors'
// gap-fill heuristic fabricated puff chains from the wreck and the garbage
// reached the database (the owner's Sunday D1320 workbook).
//
// They are now ZIP-SURGICAL: only the parts that MUST change are re-serialized
// (the one target sheet, workbook.xml, its rels, [Content_Types].xml); every
// other entry is copied through with content byte-identical. Untouched formula
// cells keep BOTH <f> and their cached <v>; foreign parts survive; styles
// (s= attributes) survive; calcChain is dropped together with its rels +
// content-type references (its entries go stale the moment a formula cell is
// overwritten); and <calcPr fullCalcOnLoad="1"/> makes Excel refresh the now-
// stale DEPENDENT caches on its next open. Two ElementTree gotchas are handled
// deliberately: every namespace present in the document is re-registered
// before parsing (so x14ac: etc. keep their prefixes), and the serialized root
// start-tag is REPLACED with the original raw one, because ET drops xmlns
// declarations that only appear in attribute VALUES (mc:Ignorable="x14ac").
//
// Both keep the v2.0.2 atomic tail: write tmp, os.replace on success only — a
// crash or any exception leaves the source workbook untouched (guarded by
// tst_excelsurgery::writeCells_atomicOnBadInput).

// Shared surgical core, embedded at the top of both scripts. Kept as one C++
// literal so the two scripts cannot drift (the SP3-T4 lesson).
#define DVE_SURGERY_PRELUDE \
    "import io\n" \
    "import os\n" \
    "import re\n" \
    "import sys\n" \
    "import zipfile\n" \
    "import xml.etree.ElementTree as ET\n" \
    "\n" \
    "NS  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'\n" \
    "NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'\n" \
    "RNS = '{http://schemas.openxmlformats.org/package/2006/relationships}'\n" \
    "Q   = '{%s}' % NS\n" \
    "\n" \
    "def col_letter(n):\n" \
    "    s = ''\n" \
    "    while n:\n" \
    "        n, r = divmod(n - 1, 26)\n" \
    "        s = chr(65 + r) + s\n" \
    "    return s\n" \
    "\n" \
    "def col_index(ref):\n" \
    "    n = 0\n" \
    "    for ch in ref:\n" \
    "        if ch.isdigit():\n" \
    "            break\n" \
    "        n = n * 26 + (ord(ch) - 64)\n" \
    "    return n\n" \
    "\n" \
    "def register_all_namespaces(data):\n" \
    "    for _ev, (prefix, uri) in ET.iterparse(io.BytesIO(data), events=('start-ns',)):\n" \
    "        try:\n" \
    "            ET.register_namespace(prefix, uri)\n" \
    "        except ValueError:\n" \
    "            pass\n" \
    "\n" \
    "def restore_root_tag(serialized, original_raw, tagname):\n" \
    "    pat = ('<' + tagname + '[^>]*>').encode()\n" \
    "    m_orig = re.search(pat, original_raw)\n" \
    "    m_new  = re.search(pat, serialized)\n" \
    "    if m_orig and m_new:\n" \
    "        return serialized[:m_new.start()] + m_orig.group(0) + serialized[m_new.end():]\n" \
    "    return serialized\n" \
    "\n" \
    "DECL = b'<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>'\n" \
    "\n" \
    "def serialize(root, original_raw, tagname):\n" \
    "    return restore_root_tag(DECL + ET.tostring(root), original_raw, tagname)\n" \
    "\n" \
    "def resolve_sheet_part(zin, sheet_name):\n" \
    "    wb_raw = zin.read('xl/workbook.xml')\n" \
    "    register_all_namespaces(wb_raw)\n" \
    "    wb_root = ET.fromstring(wb_raw)\n" \
    "    rid = None\n" \
    "    for sh in wb_root.iter(Q + 'sheet'):\n" \
    "        if sh.get('name') == sheet_name:\n" \
    "            rid = sh.get('{%s}id' % NSR)\n" \
    "    if rid is None:\n" \
    "        raise RuntimeError('sheet not found: ' + sheet_name)\n" \
    "    rels_raw = zin.read('xl/_rels/workbook.xml.rels')\n" \
    "    rels_root = ET.fromstring(rels_raw)\n" \
    "    target = None\n" \
    "    for rel in rels_root.iter(RNS + 'Relationship'):\n" \
    "        if rel.get('Id') == rid:\n" \
    "            target = rel.get('Target')\n" \
    "    if target is None:\n" \
    "        raise RuntimeError('no relationship for ' + rid)\n" \
    "    part = target if target.startswith('xl/') else 'xl/' + target.lstrip('/')\n" \
    "    return part, wb_raw, wb_root, rels_root\n" \
    "\n" \
    "def finish(path, zin, sheet_part, new_sheet, wb_raw, wb_root, rels_root):\n" \
    "    # workbook.xml: dependent-formula caches are stale after any edit -\n" \
    "    # have Excel recalculate everything on its next open.\n" \
    "    calc = wb_root.find(Q + 'calcPr')\n" \
    "    if calc is None:\n" \
    "        calc = ET.Element(Q + 'calcPr')\n" \
    "        anchor = -1\n" \
    "        for i, child in enumerate(list(wb_root)):\n" \
    "            if child.tag in (Q + 'sheets', Q + 'externalReferences', Q + 'definedNames'):\n" \
    "                anchor = i\n" \
    "        wb_root.insert(anchor + 1, calc)\n" \
    "    calc.set('fullCalcOnLoad', '1')\n" \
    "    new_wb = serialize(wb_root, wb_raw, 'workbook')\n" \
    "\n" \
    "    # calcChain refers to formula cells by address; edits/deletes make it\n" \
    "    # stale, and a stale calcChain is itself an Excel repair trigger. Drop\n" \
    "    # the part WITH its relationship and content-type override (a dangling\n" \
    "    # reference to a missing part is another repair trigger).\n" \
    "    drop = {'xl/calcChain.xml'}\n" \
    "    changed_rels = False\n" \
    "    for rel in list(rels_root.iter(RNS + 'Relationship')):\n" \
    "        if rel.get('Type', '').endswith('/calcChain'):\n" \
    "            t = rel.get('Target', '')\n" \
    "            drop.add(t if t.startswith('xl/') else 'xl/' + t.lstrip('/'))\n" \
    "            rels_root.remove(rel)\n" \
    "            changed_rels = True\n" \
    "    new_rels = DECL + ET.tostring(rels_root) if changed_rels else None\n" \
    "    ct_raw = zin.read('[Content_Types].xml').decode('utf-8')\n" \
    "    ct_new = re.sub(r'<Override PartName=\"/xl/calcChain\\.xml\"[^>]*/>', '', ct_raw)\n" \
    "\n" \
    "    tmp = path + '.dve_tmp'\n" \
    "    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:\n" \
    "        for info in zin.infolist():\n" \
    "            n = info.filename\n" \
    "            if n in drop:\n" \
    "                continue\n" \
    "            if n == sheet_part:\n" \
    "                zout.writestr(info, new_sheet)\n" \
    "            elif n == 'xl/workbook.xml':\n" \
    "                zout.writestr(info, new_wb)\n" \
    "            elif n == 'xl/_rels/workbook.xml.rels' and new_rels is not None:\n" \
    "                zout.writestr(info, new_rels)\n" \
    "            elif n == '[Content_Types].xml':\n" \
    "                zout.writestr(info, ct_new)\n" \
    "            else:\n" \
    "                zout.writestr(info, zin.read(n))\n" \
    "    zin.close()\n" \
    "    os.replace(tmp, path)\n"

const char* excelWriteCellsScript()
{
    static const char* kWriteCells = "\n"
        DVE_SURGERY_PRELUDE
        "\n"
        "def main():\n"
        "    path, sheet_name = sys.argv[1], sys.argv[2]\n"
        "    edits = []\n"
        "    args = sys.argv[3:]\n"
        "    i = 0\n"
        "    while i + 2 < len(args):\n"
        "        edits.append((int(args[i]), int(args[i + 1]), args[i + 2]))\n"
        "        i += 3\n"
        "\n"
        "    zin = zipfile.ZipFile(path)\n"
        "    sheet_part, wb_raw, wb_root, rels_root = resolve_sheet_part(zin, sheet_name)\n"
        "\n"
        "    sraw = zin.read(sheet_part)\n"
        "    register_all_namespaces(sraw)\n"
        "    sroot = ET.fromstring(sraw)\n"
        "    sdata = sroot.find(Q + 'sheetData')\n"
        "    if sdata is None:\n"
        "        raise RuntimeError('sheet has no sheetData')\n"
        "    rows = {int(r.get('r')): r for r in sdata.findall(Q + 'row')}\n"
        "    for (r1, c1, val) in edits:\n"
        "        if r1 not in rows:\n"
        "            row = ET.Element(Q + 'row', {'r': str(r1)})\n"
        "            idx = sum(1 for x in sdata.findall(Q + 'row') if int(x.get('r')) < r1)\n"
        "            sdata.insert(idx, row)\n"
        "            rows[r1] = row\n"
        "        row = rows[r1]\n"
        "        ref = col_letter(c1) + str(r1)\n"
        "        cell = None\n"
        "        for c in row.findall(Q + 'c'):\n"
        "            if c.get('r') == ref:\n"
        "                cell = c\n"
        "        if cell is None:\n"
        "            cell = ET.Element(Q + 'c', {'r': ref})\n"
        "            cells = row.findall(Q + 'c')\n"
        "            pos = sum(1 for c in cells if col_index(c.get('r', '')) < c1)\n"
        "            row.insert(pos, cell)\n"
        "        style = cell.get('s')\n"
        "        cell.clear()\n"
        "        cell.set('r', ref)\n"
        "        if style is not None:\n"
        "            cell.set('s', style)\n"
        "        v = val.strip()\n"
        "        if not v:\n"
        "            continue                     # cleared cell stays empty\n"
        "        try:\n"
        "            num = float(v)\n"
        "            el = ET.SubElement(cell, Q + 'v')\n"
        "            el.text = str(int(num)) if num == int(num) else repr(num)\n"
        "        except (ValueError, OverflowError):\n"
        "            cell.set('t', 'inlineStr')\n"
        "            is_el = ET.SubElement(cell, Q + 'is')\n"
        "            t_el = ET.SubElement(is_el, Q + 't')\n"
        "            t_el.text = val\n"
        "            if val != val.strip():\n"
        "                t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')\n"
        "    new_sheet = serialize(sroot, sraw, sroot.tag.split('}')[-1])\n"
        "    finish(path, zin, sheet_part, new_sheet, wb_raw, wb_root, rels_root)\n"
        "    print('OK')\n"
        "\n"
        "try:\n"
        "    main()\n"
        "except Exception as e:\n"
        "    import traceback\n"
        "    print('ERROR: ' + str(e), file=sys.stderr)\n"
        "    traceback.print_exc()\n"
        "    sys.exit(1)\n";
    return kWriteCells;
}

const char* excelDeleteRowScript()
{
    // Surgical row delete: remove the <row>, renumber the r= of every later
    // row and of each of its cells. Formula TEXT is deliberately left
    // untouched — exact parity with the openpyxl delete_rows this replaces,
    // which never translated references either; fullCalcOnLoad then has Excel
    // recompute whatever those formulas now mean, same as before, minus the
    // wholesale cache destruction.
    static const char* kDeleteRow = "\n"
        DVE_SURGERY_PRELUDE
        "\n"
        "def main():\n"
        "    path, sheet_name, row_s = sys.argv[1], sys.argv[2], sys.argv[3]\n"
        "    n = int(row_s)\n"
        "\n"
        "    zin = zipfile.ZipFile(path)\n"
        "    sheet_part, wb_raw, wb_root, rels_root = resolve_sheet_part(zin, sheet_name)\n"
        "\n"
        "    sraw = zin.read(sheet_part)\n"
        "    register_all_namespaces(sraw)\n"
        "    sroot = ET.fromstring(sraw)\n"
        "    sdata = sroot.find(Q + 'sheetData')\n"
        "    if sdata is None:\n"
        "        raise RuntimeError('sheet has no sheetData')\n"
        "    for row in list(sdata.findall(Q + 'row')):\n"
        "        r = int(row.get('r'))\n"
        "        if r == n:\n"
        "            sdata.remove(row)\n"
        "        elif r > n:\n"
        "            row.set('r', str(r - 1))\n"
        "            for c in row.findall(Q + 'c'):\n"
        "                ref = c.get('r', '')\n"
        "                letters = ref.rstrip('0123456789')\n"
        "                c.set('r', letters + str(r - 1))\n"
        "    new_sheet = serialize(sroot, sraw, sroot.tag.split('}')[-1])\n"
        "    finish(path, zin, sheet_part, new_sheet, wb_raw, wb_root, rels_root)\n"
        "    print('OK')\n"
        "\n"
        "try:\n"
        "    main()\n"
        "except Exception as e:\n"
        "    import traceback\n"
        "    print('ERROR: ' + str(e), file=sys.stderr)\n"
        "    traceback.print_exc()\n"
        "    sys.exit(1)\n";
    return kDeleteRow;
}

QStringList buildWriteCellsArgs(const QString& filePath,
                                const QString& sheetName,
                                const QVector<ExcelCellWrite>& cells)
{
    QStringList args = { filePath, sheetName };
    for (const ExcelCellWrite& cw : cells) {
        args << QString::number(cw.row) << QString::number(cw.col) << cw.value;
    }
    return args;
}

QStringList buildDeleteRowArgs(const QString& filePath,
                               const QString& sheetName,
                               int excelRow1)
{
    return { filePath, sheetName, QString::number(excelRow1) };
}

QVector<ExcelCellWrite> mergePendingWithInFlight(
    const QVector<ExcelCellWrite>& inFlight,
    const QVector<ExcelCellWrite>& pending)
{
    // Seed with the OLDER in-flight cells (order preserved), then overlay each
    // NEWER pending cell: same (row,col) overwrites in place (newest wins), a
    // fresh cell is appended. Behaviour-identical to the loop this replaced.
    QVector<ExcelCellWrite> merged = inFlight;
    for (const ExcelCellWrite& nw : pending) {
        bool replaced = false;
        for (ExcelCellWrite& cw : merged) {
            if (cw.row == nw.row && cw.col == nw.col) {
                cw.value = nw.value;
                replaced = true;
                break;
            }
        }
        if (!replaced)
            merged.append(nw);
    }
    return merged;
}

} // namespace DVE
