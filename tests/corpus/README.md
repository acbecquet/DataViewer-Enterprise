# v3 Golden-File Corpus

Local-only stash of REAL historical workbooks used by tst_v3shadow (and, from
Phase 2, the round-trip harness).
Point the harnesses at it with:

    $env:DVE_TEST_CORPUS_DIR = "<this directory or any dir of .xlsx>"

Populate it via the DB Data collection workflow (see the db-data-collection
memory topic / tools collect_db_data.py): source .xlsx resolved from the DB
files table under Weekly_Reports_Transfer.
Live-DB queries need owner approval.

RULES
- This repo is PUBLIC: real workbooks must NEVER be committed.
  Everything in this directory except this README is gitignored.
- Harnesses run on the synthetic tests/data fixtures alone when
  DVE_TEST_CORPUS_DIR is unset - corpus presence widens coverage, never gates.
