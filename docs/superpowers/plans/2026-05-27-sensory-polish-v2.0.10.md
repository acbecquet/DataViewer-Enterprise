# Sensory Polish v2.0.10 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Two small UX fixes for sensory mode that surfaced during smoke-testing the v2.0.10 build — (1) make the per-sample comments box outline noticeably thicker so users can actually see where the typing area is, and (2) propagate Test Title edits to the database, including a natural-key collision dialog when the rename would clobber an existing session.

**Architecture:**
- Task 1 is a single QSS edit in `SampleCard` constructor in `src/ui/SensoryPanel.cpp`. No new files, no test surface (purely visual).
- Task 2 modifies `SensoryPanel::buildSession()` to always sync `sessionName = testTitle` and adds a `UniqueViolation` branch in `MainWindow::onUpdateDatabase` that prompts with a confirm/cancel and, on confirm, deletes the conflicting row via the existing `DatabaseManager::removeSensorySession(id)` + retries the save. Uses the already-bundled `findSensorySessionsByKeys` to resolve the conflicting row id.

**Tech Stack:** C++17 / Qt 6.10 / qmake + MinGW, Postgres via `QPSQL`. No new dependencies.

---

## File Structure

**Modified files:**
- `src/ui/SensoryPanel.cpp` — `SampleCard::SampleCard` constructor (comments QSS); `SensoryPanel::buildSession()` (sessionName ↔ testTitle sync)
- `src/MainWindow.cpp` — `onUpdateDatabase()` sensory save loop (add UniqueViolation branch)
- `CHANGELOG.md` — entries under existing `[2.0.10]` section

**No new files. No new tests.** The sensory save path has no automated coverage today; verification is by manual smoke test as we've been doing this cycle.

---

## Task 1: Triple-thick comment box outline

**Files:**
- Modify: `src/ui/SensoryPanel.cpp` (the stylesheet block currently at the `m_commentsEdit->setStyleSheet(...)` call inside `SampleCard::SampleCard`)

- [ ] **Step 1: Locate the current stylesheet**

The current block reads:

```cpp
    m_commentsEdit->setStyleSheet(
        "QTextEdit { border: 1px solid #A0A6AE; border-radius: 4px; "
        "background: white; padding: 4px; }"
        "QTextEdit:focus { border: 1px solid #0066CC; }");
```

Confirm by running:

```bash
grep -n 'border: 1px solid #A0A6AE' src/ui/SensoryPanel.cpp
```

Expected: single match in `SampleCard::SampleCard`.

- [ ] **Step 2: Bump the border to 3px in both rules**

Replace with:

```cpp
    m_commentsEdit->setStyleSheet(
        "QTextEdit { border: 3px solid #A0A6AE; border-radius: 4px; "
        "background: white; padding: 4px; }"
        "QTextEdit:focus { border: 3px solid #0066CC; }");
```

Keep border-radius at 4px so the corners stay rounded. Keep the focus color the existing accent blue. Padding stays at 4px — the 3px border eats no inner space because Qt's box model is content-box.

- [ ] **Step 3: Incremental rebuild**

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: builds clean, no errors. Only `SensoryPanel.o` should recompile.

- [ ] **Step 4: Copy exe to release/ and rebuild installer**

```bash
cd "/c/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
cp build/release/DataViewer.exe release/DataViewer.exe
cmd //c '.\\build_installer.bat'
```

Expected: `dist\DataViewer-setup.exe` regenerates. Final line should read `SUCCESS! Installer created at: dist\DataViewer-setup.exe`.

- [ ] **Step 5: Stop and hand to user for smoke test**

Tell the user the comments box border has been tripled to 3px (both at rest and on focus). Do not commit yet — Task 2 will be in the same commit.

---

## Task 2: Test title propagates to database, with override warning on collision

**Files:**
- Modify: `src/ui/SensoryPanel.cpp` — `SensoryPanel::buildSession()` near the existing `// "New Session" is the placeholder…` comment block
- Modify: `src/MainWindow.cpp` — the sensory save loop inside `onUpdateDatabase()` (currently handles `Success` / `VersionMismatch` / `RowDeleted` / fallthrough-failed; add a `UniqueViolation` branch)

### Background — what's broken today

`SensoryPanel::buildSession()` currently preserves the existing `sessionName` for already-saved sessions:

```cpp
const bool isPlaceholder = existing.isEmpty()
                        || existing == QLatin1String("New Session");
if (!isPlaceholder) {
    sess.sessionName = existing;              // ← keeps OLD name
} else if (!testTitle.isEmpty()) {
    sess.sessionName = testTitle;
} else {
    sess.sessionName = QString("Session_%1").arg(...);
}
sess.testTitle = testTitle;
```

So when the user changes Test Title on an existing session, `testTitle` updates in the JSON payload but `sessionName` (the natural-key column) stays at the old value. Result: no DB row rename, the session reference list still shows the old title, and there's no way to "rename" a saved session through the UI.

`removeSensorySession(int id)` already exists at `src/database/DatabaseManager.cpp:1970` and is what we'll call to delete the conflicting row on override.

`findSensorySessionsByKeys(...)` (callable from MainWindow via `m_db->findSensorySessionsByKeys(keys)`) resolves a natural key → existing row id. Already used by `SensoryPanel::inheritExistingIdsAndVersions()`.

The `tryWriteSensoryCore` UPDATE branch uses `WHERE id = ? AND version = ?` (OCC on, server-side). On a name collision the UPDATE itself doesn't trigger the unique violation — the violation surfaces during INSERT or when the UPDATE actually changes session_name. PostgreSQL's `kSqlStateUniqueViolation` (`'23505'`) is already mapped to `WriteResult::UniqueViolation` in `tryWriteSensoryCore` (line 1629 and 1678).

- [ ] **Step 1: Sync sessionName to testTitle in buildSession()**

In `src/ui/SensoryPanel.cpp`, locate the existing block in `buildSession()`:

```cpp
    SensorySession sess;
    const QString testTitle = m_testTitleEdit->text().trimmed();
    QString existing;
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size())
        existing = m_sessions[m_currentTesterIdx].sessionName;

    // "New Session" is the placeholder newSession()/init() seed before the
    // user types a Test Title. Keeping it would let two same-day sessions for
    // the same tester collide on idx_sensory_sessions_key (session_name,
    // tester_name, date). Promote testTitle to sessionName the moment we
    // have one; otherwise fall back to a unique timestamp.
    const bool isPlaceholder = existing.isEmpty()
                            || existing == QLatin1String("New Session");
    if (!isPlaceholder) {
        sess.sessionName = existing;
    } else if (!testTitle.isEmpty()) {
        sess.sessionName = testTitle;
    } else {
        sess.sessionName = QString("Session_%1").arg(
            QDateTime::currentDateTime().toString("yyyyMMdd_HHmmss"));
    }
    sess.testTitle    = testTitle;
```

Replace with:

```cpp
    SensorySession sess;
    const QString testTitle = m_testTitleEdit->text().trimmed();
    QString existing;
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size())
        existing = m_sessions[m_currentTesterIdx].sessionName;

    // Sync sessionName to testTitle whenever the user has provided one.
    // Renaming a saved session changes its natural-key column; if the new
    // name collides with another existing row, the save loop in
    // MainWindow::onUpdateDatabase will surface a UniqueViolation and
    // prompt the user to override or cancel.
    //
    // Fallbacks (in order): keep an existing non-placeholder sessionName
    // if testTitle is blank, otherwise generate a timestamped unique name
    // so a "New Session" with no Test Title still gets a writable key.
    if (!testTitle.isEmpty()) {
        sess.sessionName = testTitle;
    } else if (!existing.isEmpty() && existing != QLatin1String("New Session")) {
        sess.sessionName = existing;
    } else {
        sess.sessionName = QString("Session_%1").arg(
            QDateTime::currentDateTime().toString("yyyyMMdd_HHmmss"));
    }
    sess.testTitle    = testTitle;
```

Why this is safe:
- For brand-new sessions (placeholder name) with a Test Title filled in → identical to today's behavior.
- For brand-new sessions with no Test Title → identical fallback to `Session_<timestamp>`.
- For previously-saved sessions where the user changes Test Title → `sess.sessionName` now reflects the new title. `tryWriteSensoryCore` updates the row by `id`, so the SQL itself is fine. If the new name collides with another row's natural key, server returns `UniqueViolation` → handled in Step 2.
- For previously-saved sessions where the user clears Test Title → falls back to the old name (avoids accidentally renaming to a timestamped scratch value).

- [ ] **Step 2: Handle UniqueViolation in onUpdateDatabase's sensory loop**

In `src/MainWindow.cpp`, locate the sensory save loop inside `onUpdateDatabase()`. The current body of the loop reads:

```cpp
    if (m_sensoryPanel) {
        auto sessions = m_sensoryPanel->allSessions();
        for (SensorySession& sess : sessions) {
            if (DVE::isPlaceholderSession(sess)) continue;
            if (!m_db) { ++failed; continue; }
            // v2.0.5: use the granular tryWriteSensorySession so we can
            // distinguish VersionMismatch (benign — LiveSync already
            // wrote this session per-cell, the file-level UPDATE is a
            // no-op against a freshly-bumped row) from a real
            // OtherError. Reporting every per-cell-edited session as a
            // failed save was a false positive.
            const DVE::WriteResult r = m_db->tryWriteSensorySession(sess);
            if (r == DVE::WriteResult::Success) {
                ++sensSaved;
            } else if (r == DVE::WriteResult::VersionMismatch
                    || r == DVE::WriteResult::RowDeleted) {
                ++sensSkipped;
                qInfo().noquote()
                    << "[onUpdateDatabase] sensory session"
                    << sess.sessionName
                    << "skipped — already up to date via LiveSync (result="
                    << static_cast<int>(r) << ")";
            } else {
                ++failed;
            }
        }
        if (sensSaved > 0)
            m_sensorySessionsDirty = false;

        // v2.0.10: allSessions() returned a copy, so tryWriteSensorySession's
        // byRef id/version back-fill only landed on local `sessions` — not on
        // SensoryPanel::m_sessions. Without this, every just-INSERTed session
        // stays at id=-1 in panel state and the next Ctrl+U / auto-save tick
        // attempts INSERT again, tripping idx_sensory_sessions_key. Bulk
        // SELECT, idempotent once every session has a positive id.
        m_sensoryPanel->inheritExistingIdsAndVersions();
    }
```

Replace the inner if/else if/else chain with one that branches on `UniqueViolation`. Final body:

```cpp
    if (m_sensoryPanel) {
        auto sessions = m_sensoryPanel->allSessions();
        for (SensorySession& sess : sessions) {
            if (DVE::isPlaceholderSession(sess)) continue;
            if (!m_db) { ++failed; continue; }
            DVE::WriteResult r = m_db->tryWriteSensorySession(sess);

            // v2.0.10: rename collision. Test Title changes propagate to
            // sessionName (the natural-key column), so a Test Title rename
            // to a name another session already holds trips
            // idx_sensory_sessions_key. Prompt the user; on confirm, delete
            // the conflicting row and retry the save.
            if (r == DVE::WriteResult::UniqueViolation) {
                QVector<DatabaseManager::NaturalKey> keys;
                DatabaseManager::NaturalKey k;
                k.sessionName = sess.sessionName.trimmed();
                k.testerName  = sess.testerName.trimmed();
                k.date        = sess.date.trimmed();
                keys.append(k);
                const auto matches = m_db->findSensorySessionsByKeys(keys);
                int conflictId = -1;
                for (const auto& m : matches) {
                    if (m.id != sess.id) { conflictId = m.id; break; }
                }
                if (conflictId <= 0) {
                    // Couldn't resolve the conflicting row — fall through
                    // to the generic failed bucket so the existing error
                    // dialog surfaces the raw DB message.
                    ++failed;
                    continue;
                }
                const auto choice = QMessageBox::warning(
                    this,
                    tr("Override Existing Sensory Session"),
                    tr("A sensory session named \"%1\" already exists for "
                       "tester \"%2\" on %3.\n\n"
                       "Overriding will permanently delete the existing "
                       "session and replace it with your current edits.\n\n"
                       "Override?")
                        .arg(sess.sessionName, sess.testerName, sess.date),
                    QMessageBox::Yes | QMessageBox::Cancel,
                    QMessageBox::Cancel);
                if (choice != QMessageBox::Yes) {
                    ++failed;
                    continue;
                }
                if (!m_db->removeSensorySession(conflictId)) {
                    qWarning() << "[onUpdateDatabase] removeSensorySession("
                               << conflictId << ") failed:"
                               << m_db->lastError();
                    ++failed;
                    continue;
                }
                r = m_db->tryWriteSensorySession(sess);
            }

            if (r == DVE::WriteResult::Success) {
                ++sensSaved;
            } else if (r == DVE::WriteResult::VersionMismatch
                    || r == DVE::WriteResult::RowDeleted) {
                ++sensSkipped;
                qInfo().noquote()
                    << "[onUpdateDatabase] sensory session"
                    << sess.sessionName
                    << "skipped — already up to date via LiveSync (result="
                    << static_cast<int>(r) << ")";
            } else {
                ++failed;
            }
        }
        if (sensSaved > 0)
            m_sensorySessionsDirty = false;

        m_sensoryPanel->inheritExistingIdsAndVersions();
    }
```

Why this shape:
- `tryWriteSensorySession` is the granular API and already returns `UniqueViolation` distinctly from the others.
- `findSensorySessionsByKeys` is the cheap bulk lookup we already trust for inherit; one extra round-trip on the rare collision path is fine.
- `removeSensorySession(int id)` already exists at `src/database/DatabaseManager.cpp:1970`. It just does `DELETE FROM sensory_sessions WHERE id = ?`. The schema has `ON DELETE CASCADE` for the sensory_images FK, so the related image rows go with it.
- Filtering `m.id != sess.id` in the match loop is the trip-wire that keeps "I just saved this same session" from looking like a collision against itself. In practice `findSensorySessionsByKeys` only returns rows matching the natural key, so if sess.id is the only match we know there's no real conflict — fall to generic failure.
- The retry assigns back to `r`, so the success/skip/fail tally below still tallies the retry's outcome.

- [ ] **Step 3: Verify it compiles**

```bash
cd build && PATH="/c/Qt/Tools/mingw1310_64/bin:$PATH" mingw32-make -j8
```

Expected: builds clean. Both `SensoryPanel.o` and `MainWindow.o` recompile. No warnings (build is `-Werror`).

If `DatabaseManager::NaturalKey` isn't visible from MainWindow.cpp, add `#include "DatabaseManager.h"` if not already there. (It already is — line ~30 of MainWindow.cpp.)

- [ ] **Step 4: Update CHANGELOG**

Add two bullets under the existing `### Fixes` block in the `[2.0.10]` section of `CHANGELOG.md`:

```markdown
- Test Title edits in sensory mode now rename the underlying database session (the `session_name` natural-key column follows the displayed Test Title). Renaming into a name another session already holds prompts an override dialog — accepting deletes the existing row and writes the current edits in its place; declining leaves the rename unsaved.
- Sample-card comments box outline tripled from 1px to 3px so it's clearly visible against the white card background, at rest and on focus.
```

- [ ] **Step 5: Copy exe, rebuild installer, hand to user**

```bash
cd "/c/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise"
cp build/release/DataViewer.exe release/DataViewer.exe
cmd //c '.\\build_installer.bat'
```

Expected: `SUCCESS! Installer created at: dist\DataViewer-setup.exe`.

Smoke test the user should run:
1. Open an existing saved session. Change Test Title. Press Ctrl+U. Reopen — the navigator should display the new title, and the DB row's `session_name` should match.
2. Rename Test Title to a name another session in the DB already holds for the same tester+date. Press Ctrl+U. The override dialog should appear. Click Cancel → no change in DB, dirty indicator stays lit. Press Ctrl+U again, this time click Yes → the other session's row gets deleted and the current edits land in its place.
3. Comments box outline should look about 3× thicker than before, visible against the white card without straining to see.

- [ ] **Step 6: Commit (after user confirms both tasks pass smoke test)**

```bash
git add src/ui/SensoryPanel.cpp src/MainWindow.cpp CHANGELOG.md
git commit -m "$(cat <<'EOF'
feat(sensory): test-title rename propagates to DB + thicker comments outline

- Test Title edits now sync session_name (the natural-key column) so a
  rename actually persists. Collisions (rename into another session's
  name on the same tester+date) surface an override dialog; accepting
  deletes the conflicting row and writes the current edits in its place.
- Triple-thick (1px → 3px) outline on the per-sample comments box,
  visible at rest and on focus, since the previous 1px border was hard
  to see against the white card background.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

The version bump to 2.0.11 + tag + push happens after the user confirms; that's tracked separately in our session.

---

## Self-Review Notes

- Spec coverage: Task 1 covers the comments-outline ask; Task 2 covers both the rename-propagates-to-DB ask and the override-warning ask.
- Placeholder scan: every code block contains real code, every shell command is executable as written. No "implement appropriate error handling" hand-waves.
- Type consistency: `DatabaseManager::NaturalKey`, `SensoryPanel::inheritExistingIdsAndVersions()`, `removeSensorySession(int)` all match the existing codebase signatures verified in the header files.
- The override path deletes the conflicting row outright. An alternative would be to keep the conflicting row and rename `sess.sessionName` to a uniquified variant — but the user asked specifically for an "override" semantic, so a destructive replace is the right shape.
