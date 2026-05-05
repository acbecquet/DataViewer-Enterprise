"""
Read-only diagnostic against the production DataViewer DB.

Surfaces:
  - files with template_version='unknown' (would be auto-purged by deduplicateFiles)
  - files with 0 sheets or 0 samples (likely from a failed/partial save)
  - orphaned tests/samples/data_rows (parent file/test/sample missing)
  - duplicate file_path values
  - distribution of templateVersion across files
  - SQLite integrity_check + foreign_key_check
  - WAL/journal status
"""
import os
import sqlite3
import sys
from collections import Counter
from pathlib import Path

DEFAULT_DB = Path(os.environ.get("USERPROFILE", "")) / "SynologyDrive/SDR/Device Group/DVE_Database/dataviewer.db"


def main():
    db_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_DB
    if not db_path.exists():
        print(f"ERROR: DB not found at {db_path}")
        return 2

    print(f"DB: {db_path}  ({db_path.stat().st_size / 1024 / 1024:.1f} MB)")

    # Open read-only via URI to avoid any chance of mutation
    uri = f"file:{db_path.as_posix()}?mode=ro"
    con = sqlite3.connect(uri, uri=True)
    cur = con.cursor()

    print("\n=== SQLite integrity checks ===")
    try:
        cur.execute("PRAGMA quick_check")
        qc = cur.fetchone()[0]
        print(f"quick_check     : {qc[:200] + ' ...' if len(qc) > 200 else qc}")
    except sqlite3.DatabaseError as e:
        print(f"quick_check     : DatabaseError ({e})")
    try:
        cur.execute("PRAGMA foreign_key_check")
        fk = cur.fetchall()
        print(f"foreign_key_check : {len(fk)} violations" + ((" — " + str(fk[:5])) if fk else ""))
    except sqlite3.DatabaseError as e:
        print(f"foreign_key_check : DatabaseError ({e})")
    cur.execute("PRAGMA journal_mode")
    print(f"journal_mode    : {cur.fetchone()[0]}")

    print("\n=== Table row counts ===")
    for tbl in ("files", "tests", "samples", "data_rows", "images",
                "sensory_sessions", "sensory_images",
                "detailed_sensory_sessions"):
        try:
            cur.execute(f"SELECT COUNT(*) FROM {tbl}")
            print(f"  {tbl:32s} {cur.fetchone()[0]:>8d}")
        except sqlite3.OperationalError as e:
            print(f"  {tbl:32s} (missing: {e})")

    print("\n=== TPM files breakdown ===")
    cur.execute("""SELECT id, file_path, file_name, loaded_at, template_version,
                          sheet_count, sample_count
                   FROM files ORDER BY loaded_at DESC""")
    rows = cur.fetchall()
    print(f"  Total files: {len(rows)}")
    tv_counter = Counter(r[4] for r in rows)
    print(f"  templateVersion distribution: {dict(tv_counter)}")

    unknown_or_empty = [r for r in rows if r[4] in ("unknown", None, "")]
    if unknown_or_empty:
        print(f"\n  *** {len(unknown_or_empty)} files would be DELETED on next deduplicateFiles() call:")
        for r in unknown_or_empty[:20]:
            print(f"      id={r[0]:5d}  tv={r[4]!r:12s}  sheets={r[5]:3d}  samples={r[6]:3d}  {r[2]}")
        if len(unknown_or_empty) > 20:
            print(f"      ... and {len(unknown_or_empty) - 20} more")

    zero_sheets = [r for r in rows if r[5] == 0]
    if zero_sheets:
        print(f"\n  {len(zero_sheets)} files have sheet_count=0 (cosmetic? or save failed):")
        for r in zero_sheets[:10]:
            print(f"      id={r[0]:5d}  tv={r[4]!r:12s}  {r[2]}")

    print("\n=== Duplicate file_path detection ===")
    cur.execute("""SELECT file_path, COUNT(*) c FROM files
                   GROUP BY file_path HAVING c > 1 ORDER BY c DESC""")
    dups = cur.fetchall()
    if dups:
        print(f"  {len(dups)} duplicated file_path entries (UNIQUE index should prevent this!):")
        for path, c in dups[:10]:
            print(f"      {c}x  {path}")
    else:
        print("  none")

    print("\n=== Orphan check (children whose parent rows are missing) ===")
    queries = [
        ("tests with missing file_id",
         "SELECT t.id, t.file_id, t.sheet_name FROM tests t LEFT JOIN files f ON f.id=t.file_id WHERE f.id IS NULL"),
        ("samples with missing test_id",
         "SELECT s.id, s.test_id, s.sample_name FROM samples s LEFT JOIN tests t ON t.id=s.test_id WHERE t.id IS NULL"),
        ("data_rows with missing sample_id",
         "SELECT d.id, d.sample_id FROM data_rows d LEFT JOIN samples s ON s.id=d.sample_id WHERE s.id IS NULL"),
        ("images with missing sample_id",
         "SELECT i.id, i.sample_id, i.file_name FROM images i LEFT JOIN samples s ON s.id=i.sample_id WHERE s.id IS NULL"),
    ]
    for label, sql in queries:
        cur.execute(sql)
        orphans = cur.fetchall()
        if orphans:
            print(f"  *** {len(orphans):5d} {label}")
            for o in orphans[:5]:
                print(f"        {o}")
        else:
            print(f"  {0:5d} {label}")

    print("\n=== Files with 0 actual sample rows (vs claimed sample_count) ===")
    cur.execute("""
        SELECT f.id, f.file_name, f.sample_count,
               (SELECT COUNT(*) FROM samples s
                JOIN tests t ON t.id=s.test_id
                WHERE t.file_id=f.id) AS actual
        FROM files f
        ORDER BY f.id
    """)
    mismatches = []
    for fid, fname, claimed, actual in cur.fetchall():
        if claimed != actual:
            mismatches.append((fid, fname, claimed, actual))
    if mismatches:
        print(f"  *** {len(mismatches)} files where sample_count != actual rows:")
        for fid, fname, c, a in mismatches[:15]:
            print(f"      id={fid:5d}  claimed={c:3d}  actual={a:3d}  {fname}")
        if len(mismatches) > 15:
            print(f"      ... and {len(mismatches) - 15} more")
    else:
        print("  all match")

    print("\n=== Sensory side (control — user reports this works) ===")
    cur.execute("SELECT COUNT(*) FROM sensory_sessions")
    print(f"  sensory_sessions     : {cur.fetchone()[0]}")
    cur.execute("SELECT COUNT(*) FROM detailed_sensory_sessions")
    print(f"  detailed_sensory_sessions : {cur.fetchone()[0]}")

    con.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
