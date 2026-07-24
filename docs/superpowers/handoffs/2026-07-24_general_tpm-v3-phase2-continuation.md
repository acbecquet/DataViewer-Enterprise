---
date: 2026-07-24
researcher: Claude (Fable 5, subagent-driven session)
git_commit: 9484030 (+ same-day build commits for the v2.10.1 smoke installer, see Artifacts)
branch: worktree-tpm-template-v3-research
repository: DataViewer-Enterprise
topic: "TPM Template v3 Phase 2 (write-back + manifest) Implementation Strategy"
tags: [tpm-template-v3, schema-driven, strangler-fig, excel-pipeline, manifest, metric-model]
status: phase-0-1-code-complete, awaiting owner E2E smoke + merge decision
last_updated: 2026-07-24
last_updated_by: Claude (Fable 5)
type: implementation_strategy
---

# Handoff: TPM Template v3 - continue from Phase 0+1 into Phase 2

## Task(s)

1. **Phase 0 + Phase 1 of the v3 free-standing metric model - DONE.**
   All 12 tasks of the implementation plan executed subagent-driven (implementer + spec review + quality review per task, final whole-implementation review passed with zero blocking findings).
   Production Excel parsing now runs through `DVE::model` (`DataProcessor::processFile/processSheet` are schema-driven); the legacy positional parser survives only as `processFileLegacy/processSheetLegacy`, the shadow-harness referee.
   Source plan: `docs/superpowers/plans/2026-07-23-tpm-v3-phase0-1-core-model-plan.md` (all checkboxes done).
2. **v2.10.1 internal smoke installer - built this session** (see Artifacts) so the owner can run a full E2E smoke of the flipped read path.
   Internal patch build per the versioning scheme; NEVER goes to Synology.
3. **Next up (not started): Phase 2 - write-back + manifest.**
   Scope per spec sections 7 and 14, refined by the owner feedback in spec section 18 (read it first, it is binding).

## Critical References

- `docs/superpowers/specs/2026-07-16-tpm-template-v3-metric-model-design.md` - THE design spec; sections 6 (standard-v1), 7 (parsing/write-back), 17 (deferred decisions incl. the open unit-reconciliation item), **18 (owner feedback 2026-07-24, governs Phase 2+)**, and the Phase log at the end.
- `docs/superpowers/plans/2026-07-23-tpm-v3-phase0-1-core-model-plan.md` - the executed Phase 0+1 plan; its "Machine + repo rules" and "Build & test commands" sections apply verbatim to Phase 2 work.
- Memory topic file `tpm-template-v3-initiative.md` (in the Claude project memory dir, outside the repo) - decision history + Phase 2 hardening ledger.

## Recent Changes

All on branch `worktree-tpm-template-v3-research` (LOCAL-ONLY, not pushed), base `9b2aa64` (= v2.10.0 main).

- `src/model/MetricDef.h`, `src/model/MetricSample.h` - value types (MetricDef/HeaderFieldDef/AggregateDef; MetricSeries/Sample).
- `src/model/TemplateSchema.h/.cpp` - schema container + lookups.
- `src/model/StandardSchema.h/.cpp` - `standardV1(bool perRowRegime, HeaderLayout)` with real-template header aliases and Cart/Project header-layout variants.
- `src/model/SchemaDrivenReader.h/.cpp` - `parseSheet(...)` with `ColumnResolution::NameFirst|Positional`; model::Sheet/File live here.
- `src/model/LegacyAdapter.h/.cpp` - model -> `ExcelReader::SampleData` lowering, mirrors `extractMetadata` cell-for-cell (incl. the voltage>0 && denom>0 power guard).
- `src/ExcelReader.h/.cpp` - added `currentSheetCells()` typed-grid accessor (only additive change).
- `src/pipeline/DataProcessor.h/.cpp` - THE FLIP: `processFile/processSheet` are schema-driven (positional resolution + per-block Cart/Project landmark sniff at the top of `processSheet`); `processFileLegacy/processSheetLegacy` retained as referee; shared `processSopSheet`.
- `tests/common/CorpusUtils.*`, `tests/common/JsonDiff.*` - harness utilities.
- `tests/tst_v3harness`, `tests/tst_v3model`, `tests/tst_v3shadow` - new suites (all registered in `tests/tests.pro`).
- `tests/generate_fixtures.py` + `tests/data/old_standard.xlsx` - 9th fixture locking the old-standardized heating-tech drop (proven non-vacuous by a deliberate red run).
- `DataViewerEnterprise.pro` - model sources registered; VERSION bumped to 2.10.1 for the smoke build.

## Learnings

- **Gate evidence (all green 2026-07-23):** tst_v3shadow 20/20 0-skipped (production byte-identical to legacy over 9 fixtures), tst_v3model 17/0, full suite `tests/run-tests.ps1` 58/0/0, -Werror app build, binary smoke 11/12 (only the postgres driver missing from the ad-hoc smoke tree - deploy artifact, not a regression).
- **Positional resolution in production is Phase-1-only.** Name-first matching would have reshuffled data columns on historical layouts (format_a is a misaligned 13-wide block); byte-identity required positional. NameFirst + aliases are built and unit-tested, dormant until manifests land. The owner has confirmed name-first is a core goal (spec section 18).
- **Cart/Project header-layout sniffing is PERMANENT legacy support** (landmark cells row1col0 "Cart" / row0col5 "Project:", mirroring `ExcelReader::extractMetadata`). It migrates into Phase 2's SchemaResolver; never delete it with Phase 4 scaffolding.
- **Plan-embedded code needs the same review as hand-written code.** The review loop caught a vacuous alias test, an asymmetric python-skip guard, and a dead taken[] guard - all authored in the plan itself.
- **Machine quirks (S1134987):** create source files with the Write/Edit tools, NOT python-writes (trusted Python MIP-labels files here - opposite of the project CLAUDE.md, see memory `mip-labels-python-writes`); run `python tools/decrypt_via_copy.py --apply` before builds; Qt Test binaries print NOTHING to stdout in this Bash - always `-o results.txt,txt`; qmake DEFINES with spaced paths need `$$shell_quote`; read ciphertext files via `git show HEAD:<path>`.
- **Smoke-testing the GUI on this machine:** the owner runs a live installed DataViewer; `SingleInstance` uses a fixed key, so launching a dev build forwards the file into the owner's session unless the dev build gets a temporary unique key. Never `taskkill /IM` DataViewer (kills the owner's instance); kill by PID. NEVER kill Excel processes (memory `never-kill-excel-processes`).
- **Worktree installer builds need a seeded release tree:** `installer.iss` packages Qt DLLs / libpq / plugins / `python_bundle.zip` from `release\`; a fresh worktree lacks them. Copy the main repo's complete `release\` tree first; `mingw32-make clean` removes only objects, not deployed DLLs.
- **Residual-risk register** (final review): R1 shadow gate QSKIPs without python (ensure CI/corpus runs have python); R2 gate proves identity-with-legacy, not ground truth (by design); R3 SOP/raw-table branch structurally shared but not gate-covered (optional fixture); R4 name-first unproven against real reordered workbooks until Phase 2; R5 final review was inspection-based, gates independently reproduced earlier.

## Artifacts

- Spec: `docs/superpowers/specs/2026-07-16-tpm-template-v3-metric-model-design.md` (sections 1-18 + Phase log).
- Plan (executed): `docs/superpowers/plans/2026-07-23-tpm-v3-phase0-1-core-model-plan.md`.
- This handoff: `docs/superpowers/handoffs/2026-07-24_general_tpm-v3-phase2-continuation.md`.
- Sprint board: `docs/sprint-tracker.html` (v3 sprint state, honest statuses).
- Corpus contract: `tests/corpus/README.md` (`DVE_TEST_CORPUS_DIR`; real workbooks NEVER committed - repo is public).
- Smoke installer: `dist\DataViewer-setup.exe` (v2.10.1, copied to the main repo `dist\` per the rebuild skill) + `release_overview/release_overview_v_2_10_1.txt`.
- Key commits (oldest first): `b239b33` spec; `2c78b59` plan; `7a05e62`/`b1e7ab0` harness utils; `43b7f0e`-range Group B (shadow baseline + corpus docs); `2efcefd`/`10087ea`/`5b92fd6` model + standard-v1; `be9b253` grid accessor; `19011a8`/`b02d47c`/`237dfba` SchemaDrivenReader; `efc0476` LegacyAdapter; `3b19a86` processFileV3 + shadow gate; `7168e99` old_standard fixture lock; `528ccf9` THE FLIP; `318988e` Phase log; `9484030` owner-feedback section 18.
- Plane: no v3 work items created yet (initiative tracked in-repo + memory; a chip exists for the Plane MCP list-endpoint 404s).

## Action Items & Next Steps

1. **Owner: E2E smoke with the v2.10.1 installer** - open several REAL workbooks (include at least one Cart-era, one Project-era, one old-standardized, one current per-row-regime file), check tables/plots/reports render identically, edit qualitative cells and confirm Excel write-back, exercise DB save/load + a second client if convenient.
   Write-back and DB paths are UNCHANGED by Phase 1 (only reading moved), so regressions there would indicate an adapter fidelity gap - report with the file that shows it.
2. **Owner: merge decision.** Per the branch-to-main workflow, merge/push only after the smoke passes.
   Branch is local-only; nothing is pushed anywhere.
3. **Phase 2 planning session (fresh session, after merge):** invoke superpowers writing-plans against spec sections 7 + 14 + 18.
   Scope: `_dve_schema` manifest reader + `SchemaResolver` (absorb the Cart/Project sniff, the Project sampleID join from `DataProcessor::processSheet`, and an extracted `readBlockHeaders`), `CellAddressMap` write-back replacing the `sampleIndex*12+col+1` math in `MainWindow`, NameFirst activation when a manifest is present (with real collision handling replacing the documented pass-2 gap), `MetricDef.tags` (freeform key->value from extra manifest columns), sidecar manifest emission, poka-yoke validation (warn, never gate).
   Round-trip harness (parse -> reconstruct -> cell diff) becomes the Phase 2 gate; skeleton pieces (corpus + JsonDiff) already exist.
4. **Owner decisions pending:** reconcile `MetricDef.unit` vs template header text for `tpm_power_density` and `oil_consumed` (spec section 17 last bullet); optional SOP fixture to close residual risk R3.
5. **Phase 3 note (not now):** the open-ended per-index payloads from section 18 (images, documents, blobs at row index i) require the `measurements` table to extend beyond value_num/value_text (value_blob BYTEA and/or value_json JSONB, or an attachment side-table); decide during Phase 3 planning, keeping row-index alignment (`sort_order`) as the invariant.

## Other Notes

- Build commands, suite-run recipe, and environment rules are in the plan's preamble - reuse them verbatim.
- The v2.10 sprint's DV-28 needs-save gates and DV-26 SampleColorMap are v3 constraints (Phase 3 save paths need a needs-save gate; schema-driven plots must route colors through `SampleColorMap::colorsForPlot`).
- Token-efficiency directive stands: no big agent fan-outs; subagent-driven per-task with double review worked well and caught real defects - reuse the cadence.
- The sprint-tracker stop hook fires on every turn that lands commits; keep `docs/sprint-tracker.html` honest as you go, statuses never overclaim ("ready" = awaiting owner sign-off).
- Docker/test-Postgres: suites skip cleanly without `DVE_TEST_PG_CONN`; if the DB suites are needed, ASK the owner to start Docker (memory `docker-ask-user-to-start`).
