# AGENT_PLAN: Live Word Backend (COM) for MsWord-ai-integration (v3)

> **Audience:** the coding agent working in this workspace (Antigravity).
> **Read first:** `WORKSPACE.md`, `.agents/rules/document-styling.md`, `src/oxml_engine.py`, `src/server.py`.
> **Human-facing rationale:** the separate PDF. This file is the executable spec. Do not ask the user to re-explain context that is already here.
> **Location:** store this file at `.agent-info/docs/plans/live-word-backend.md` (per `info-dir`).
> **v2 changes:** resolved contradictions (response prefix, C6 carve-outs, dry-run ordering), decided the tracked-deletion index rule, added experiments, hardened detection, added guided creation of one rule and one skill.
> **v3 changes:** (1) table paragraphs excluded from `[p:N]` by decision, to be confirmed in E1; (2) final-view reading computed from `doc.Revisions`, never by toggling view settings; (3) bulk-read performance design and budget; (4) worker-thread timeout and "poisoned" state moved into Phase 0; (5) Track Changes policy split: content edits tracked, pure formatting untracked (still one undo step); (6) agent revision author name (E7); (7) `find_and_replace` implemented by locate-then-set, removing the replacement-length limit and excluding hidden deleted text; path normalization hardened; insert style parity; backup retention; experiment time-boxes replaced by call caps; pre-decided fallbacks for the S1 prefix/outline conflicts; optional `InsertXML` spike (E8) before the two hardest Phase 2 ports.

---

## 0. How to execute this plan

1. Work **phase by phase**. Do not start a phase until the previous **gate** is passed.
2. Tick checkboxes in this file as tasks complete (edit in place).
3. Gates marked **[HUMAN]** need the user to act or verify in Word. Stop and ask. You cannot see Word's window.
4. If a **stop condition** (section 9) triggers, stop, report, wait. Per `problem-resolve`, this applies to **any** error or failed test: stop on the first failure.
5. Commit per task, following `git-usage`. Never bundle unrelated changes.
6. Experiments (E1-E8) produce **decisions, not product code**. Record results in `.agent-info/docs/reviews/live-word-experiments.md`. **Cap: at most 15 tool calls per experiment.** If it is unresolved at the cap, log what you learned and stop (S7). If an experiment contradicts an assumption in this plan, stop (S7) and report before building on it.

---

## 0.1 Workspace rules that apply (do not duplicate or override)

| Rule | How it applies here |
|------|---------------------|
| `problem-resolve` (always on) | On any error, failed test, or unexpected result: stop. No retries, workarounds, or speculative fixes. Report and wait. Overrides any "try again" wording elsewhere. |
| `git-usage` | Branch `feat/live-word-backend`. Stage by explicit path. Inline `-m` Conventional Commits, e.g. `feat(live): add open-document detection`. `git diff --stat` before any diff; never dump a diff over 300 lines (expect `live_engine.py` to exceed this; use targeted snippets). Never commit or push to `main`. |
| `gh-usage` | At the end: `git push -u origin feat/live-word-backend`, then a **draft** PR with inline `--title` and `--body`, no `--base`. Never `gh pr merge`; present PR number, URL, and strategy options to the user. |
| `info-dir` | Notes, experiment logs, gate reports, scratch files go in `.agent-info/` (gate reports: `.agent-info/docs/reviews/live-word-gate-N.md`). Source, tests, README, WORKSPACE.md, rules, and the skill stay in their normal locations. |
| `document-styling` | Pipeline unchanged. Live mode is transparent to it: same four tool calls plus the metadata check. Edited only as specified in Phase 4a. |
| `live-word-mode` (new) | Created in Phase 4a from Appendix A. |
| `live-word-editing` skill (new) | Created in Phase 4b from Appendix B, after Gate 3. |

---

## 1. Goal and non-goals

**Goal:** the 8 MCP tools can edit a document **currently open in Microsoft Word** (via COM), so the user can watch changes appear, undo them with Ctrl+Z, and (for content edits) review them as tracked changes, without closing the file.

**Non-goals:**
- Do not rewrite or refactor `DocxEngine` / `oxml_engine.py` beyond the additive items in 3.3.
- Do not change tool names, parameter names, or the 0-based `[p:N]` / `[table:N]` convention.
- Do not build an Office.js add-in, web server, or hosting.
- Do not support macOS/Linux live mode. Headless stays the cross-platform path.
- Do not add MCP tools. The surface stays at 8.

---

## 2. Hard constraints

| # | Rule |
|---|------|
| C1 | Tool-facing indices stay **0-based**. COM is 1-based and, because of tracked deletions and tables (3.5), not contiguous with tool indices. All conversion goes through one helper, `_com_idx(doc, n)`. Never inline `+ 1`. |
| C2 | `pywin32` is **optional**. `import src.server` must work without it (guard the import, set `HAS_COM`). |
| C3 | The existing **29 tests keep passing** at every commit (`.venv\Scripts\python.exe -m pytest tests/ -m "not live" -v`). |
| C4 | Live mode **never launches Word**. Attach only with `win32com.client.GetActiveObject("Word.Application")`. Never use `Dispatch`, `DispatchEx`, or `gencache.EnsureDispatch` in engine code. |
| C5 | Every mutating live call is **one undo step** (a custom undo record). **Content edits** (`mutate_element`, `find_and_replace`, `compose_document`, and the deleting flags of `clean_structure`) run with **Track Changes ON**. **Pure formatting/layout edits** (`set_page_layout`, `update_styles`, `format_table_grid`, and `clean_structure`'s color normalization) run with Track Changes **as the user has it** (not forced on), to avoid floods of formatting revisions; they are still one Ctrl+Z. Whatever the policy, the user's previous `TrackRevisions` value is restored afterwards, including on error. |
| C6 | If the file is not open in Word, behavior matches today, **except** the additive carve-outs in 3.3 (all must keep the 29 tests green unmodified). |
| C7 | No hardcoded absolute paths. Docs use relative links. |
| C8 | Live tests touch only `documents/live_test.docx`, created by the test fixture and listed in `.gitignore`. |
| C9 | **One dedicated COM worker thread** (a queue plus futures). Call `CoInitialize()`/`CoUninitialize()` there. COM objects never cross threads and are never cached between tool calls (re-attach each call). The worker has a **per-call timeout and a "poisoned" state** (3.8). |
| C10 | Obey `problem-resolve`: no retry behavior by the agent. Engine **code** may include one bounded retry (max 3 attempts, short backoff) for Word's "busy / call rejected" condition only; comment it in code. |
| C11 | Obey `git-usage`, `gh-usage`, `info-dir` (see 0.1). |
| C12 | The live engine **never calls Save or Close** on a user document. (Test fixture may close only its own `live_test.docx` with `SaveChanges=False`.) |
| C13 | The live engine **never changes the user's window, view, or application options**, including `View.ShowRevisionsAndComments`, markup mode, `ScreenUpdating`, and zoom. The only user-visible state it may touch is `TrackRevisions` and (if E7 approves) `Application.UserName`/`UserInitials`, each restored in `finally`. |

---

## 3. Target design

### 3.1 Files

| File | Action |
|------|--------|
| `src/live_engine.py` | **New.** `LiveWordEngine`, detection, `pick_backend()`, worker thread, undo/track wrapper, visible-paragraph map. |
| `src/server.py` | **Modify.** Each tool calls `pick_backend()` and dispatches; adds the `[backend: ...]` prefix in the server layer (see 3.3 for the pre-decided fallbacks). Signatures unchanged except additive optional params. |
| `requirements.txt` | Add `pywin32` with marker `sys_platform == "win32"`. |
| `tests/test_live_engine.py` | **New.** `@pytest.mark.live`, auto-skipped when Word/COM unavailable. |
| `pytest.ini` (or existing config) | Register the `live` marker. |
| `.gitignore` | Add `documents/.backups/` and `documents/live_test.docx`. |
| `WORKSPACE.md`, `README.md` | Document the backend (Phase 4c). |
| `.agents/rules/live-word-mode.md` | **New** (Phase 4a). |
| `.agents/rules/document-styling.md` | **One-step amendment** (Phase 4a). |
| `.agents/skills/live-word-editing/` | **New** (Phase 4b), subject to the location check there. |

### 3.2 Routing and detection

```
tool call (file_path, ...)
   -> pick_backend(file_path)
```

Evaluate in this order and stop at the first match:

1. `WORD_BACKEND=headless` -> headless.
2. `HAS_COM` false, or Word not running (`GetActiveObject` fails) -> go to step 6 (lock-file check), not straight to headless.
3. Word running. Compare `_norm(file_path)` with `_norm(doc.FullName)` for each item in `Application.Documents` -> **live** if found. `_norm` = `os.path.realpath`, then `os.path.normcase`, then (on Windows, when the file exists) `win32api.GetLongPathName` to expand 8.3 short names. If the strings differ but both paths exist on disk, fall back to `os.path.samefile` (covers mapped drives, UNC aliases, subst drives). Never match by basename alone.
4. Not in `Documents`: check `Application.ProtectedViewWindows` (Protected View documents are **not** in `Documents`). If found -> `LiveWordError` (Protected View; user must click Enable Editing).
5. Any open document whose `FullName` is an `http(s)://` URL and whose `Name` equals the target file's basename -> `LiveWordError` (cloud-hosted; copy local, or close it and set `WORD_BACKEND=headless`).
6. **Lock-file safety net:** look in the file's folder for a Word owner file (`~$` + the file name with its first 1-3 characters dropped; confirm the exact naming in experiment E3). If one exists but steps 3-5 found nothing, Word probably has the file open in a way we cannot attach to (elevated Word, second instance, stale lock):
   - **Mutating tools** -> `LiveWordError`: refuse to write; ask the user to close the file, run Word non-elevated, or confirm the lock is stale.
   - **`read_document`** -> run headless, and prefix `[backend: headless]` plus a note that the file appears open in Word and the content is the last saved state.
7. Otherwise -> headless.

`WORD_BACKEND=live` with the file not open -> `LiveWordError` ("not open in Word"). Env var values: `auto` (default) | `live` | `headless`.

### 3.3 Response contract, additive carve-outs, and pre-decided fallbacks

- **Prefix line.** Every tool response begins with `[backend: live]` or `[backend: headless]`. Add it **in `server.py`**, not in the engines, so engine-level tests are unaffected. **Phase 0 task:** inventory which of the 29 tests assert on tool-level response text.
- **Local outline.** `mutate_element` returns the local outline on **both** backends: up to 2 visible paragraphs before and after the edit, as `[p:N] text...` with fresh indices.
- **Allowed C6 carve-outs (all additive, defaults preserve current output otherwise):** the prefix line, the headless local outline, `dry_run` parameters, automatic backups, real replacement counts.
- **Pre-decided fallbacks (so Phase 0 does not dead-end).** If the inventory shows an original test would break:
  - *Prefix conflict:* do **not** edit the test. Fall back to **prefix on live responses only**; the absence of a `[backend: ...]` line means headless. Log the decision in the Gate 0 report; Appendix A rule 1 and the skill use the "live-only prefix" wording variant.
  - *Headless outline conflict:* do **not** edit the test. Fall back to **outline returned on live only**; on headless the agent re-reads the outline after each mutation. Appendix A rule 3 uses the "live-only outline" variant.
  - Any **other** conflict -> S1 as before.
  - Record which variant was chosen at Gate 0 so the rule, skill, docs, and docstrings stay consistent.

### 3.4 Error contract

All COM failures are caught and re-raised as one `LiveWordError` with an agent-readable message:

| Situation | Message must say |
|-----------|------------------|
| Index out of range | valid range, e.g. `[p:99] out of range (0..41)` |
| Read-only / protected / Protected View | the state, and that the user must enable editing in Word |
| OneDrive/SharePoint | cloud-hosted; copy local, or close it and set `WORD_BACKEND=headless` |
| Modal dialog / call timeout | Word may be waiting on a dialog; ask the user to dismiss it |
| Worker poisoned (earlier call still blocked) | an earlier call has not returned; Word is probably still blocked; ask the user to check Word for a dialog |
| `WORD_BACKEND=live` but not open | the file is not open in Word |
| Lock file but cannot attach | file appears open in Word; list likely causes (3.2 step 6) |
| Tool not yet live-capable (Phases 1-2) | name the tool and say the file is open in Word, so nothing was written |

Keep the exact strings in one module-level dictionary so docs and the skill's error catalog can be generated from it (Phase 4b).

### 3.5 Reading semantics (decided; verify in E1)

Live reads present the **final view of body paragraphs**:

- **Tracked deletions.** Text of tracked-deleted runs is excluded, and a paragraph that is **entirely** a tracked deletion (all text and its paragraph mark) is skipped in numbering. Tracked insertions are included. Reason: tracked deletions stay in `doc.Paragraphs`, which would otherwise make live indices diverge from headless indices after a delete. This also applies to tracked deletions the user made before the agent started.
- **Tables.** Paragraphs inside tables are **not** numbered as `[p:N]`; tables are addressed by `[table:N]`, matching the existing headless convention. How headless addresses text inside table cells is read from `oxml_engine.py` and recorded in E1; the live map follows that convention exactly. Use `Range.Information(wdWithInTable)` (or `Range.Tables.Count`) to classify.
- **How the final view is computed:** from `doc.Revisions` (type `wdRevisionDelete`) character ranges, **never** by toggling `ShowRevisionsAndComments`, markup mode, or any view setting (C13). Visible text of a paragraph = its `Range.Text` with deleted-revision ranges removed. Visible offset to range offset mapping is needed for `replace` and for `find_and_replace` (3.9).
- `_visible_map(doc)` returns the list of COM 1-based paragraph positions that are visible body paragraphs, plus per-paragraph deleted spans.
- `_com_idx(doc, n)` looks up `n` in that map and raises the out-of-range error with the visible range.
- **E1 must confirm** (a) headless numbering excludes `w:del` content and table paragraphs the same way, (b) how `Range.Text` behaves when deleted text is present, and (c) character-offset correctness around fields and hidden text. If headless and live numbering cannot be reconciled without changing headless behavior, stop (S2).

### 3.6 Undo and Track Changes wrapper

One context manager used by **every** mutating live call, signature roughly `live_call(tool, target, *, track: bool | None)`:
- start a custom undo record labeled `Agent: <tool> <target>`,
- save `doc.TrackRevisions`; if `track is True` set it `True`; if `None` leave it as the user has it,
- optionally set the revision author (3.10),
- sub-steps may use `with tracking(doc, on)` to flip tracking inside one undo record (needed by `clean_structure`, which mixes deleting and formatting flags),
- in `finally`: restore `TrackRevisions` (and author), end the undo record.

No mutating live call may bypass it.

### 3.7 Backups (Phase 3, decided approach)

Word's `Document` has **no `SaveCopyAs`**. Approach: before the first mutation on a file per server session, copy the **on-disk** file to `documents/.backups/<n>_<timestamp>.docx` with `shutil.copy2`. For live mode, record `doc.Saved`; if it is `False`, the response notes that the backup reflects the last saved state (open edits are protected by undo, and by Track Changes for content edits). **Retention:** keep the newest 20 backups per source file (module constant `BACKUP_KEEP = 20`); prune older ones after each new backup. E2 confirms the copy works while Word has the file open. If E2 fails, stop (S7).

### 3.8 COM worker, timeouts, and the poisoned state (Phase 0, not Phase 3)

- All COM work is submitted to the single worker thread through a queue; the caller waits on a future with a **per-call timeout** (default 20 s, env `LIVE_CALL_TIMEOUT`).
- A COM call blocked on a modal dialog cannot be cancelled. On timeout: mark the worker **poisoned**, return the "modal dialog / call timeout" error to the caller. While poisoned, later calls **fail fast** with the "worker poisoned" error and are **not** queued behind the stuck call. When the blocked call finally returns (user dismissed the dialog), clear the poisoned flag; the next call proceeds normally.
- Never spawn a second COM thread to "work around" a stuck one.
- Test this with a fake slow callable (no real dialog needed); the real-dialog case is a [HUMAN] check at Gate 3.

### 3.9 Find and replace design (live)

- **Locate-then-set**, not `Execute(Replace=wdReplaceAll)`: use `Range.Find.Execute` (no replace) in a loop to locate each match, skip any match whose range lies inside a tracked-deleted span, then set `match_range.Text = replacement`. Setting `Range.Text` is literal, so replacement strings need no escaping and have **no 255-character limit**.
- The **search** string keeps the 255-character limit on both backends, and special characters (`^`, etc.) are escaped for Word's Find. Wildcards off.
- Count = number of matches actually replaced (the real count, additive on headless too). `dry_run` returns the count and does not modify.
- Iterate so earlier replacements do not shift later match ranges incorrectly (collect ranges first and apply in reverse order, or re-find from the end of the last replacement).
- E1 checks whether `Find` matches hidden tracked-deleted text; if it does, the skip rule above is mandatory and gets a dedicated test.

### 3.10 Revision author (decided pending E7)

Agent edits would otherwise appear under the user's own name. If E7 shows it is safe, `live_call` sets `Application.UserName` (and `UserInitials`) to `LIVE_AUTHOR` (default `Agent`; empty string disables) for the duration of the call and restores both in `finally`. If E7 shows a crash can leave the name changed, or the setting is not honored for revisions, drop the feature and document that agent and user edits share the author name.

### 3.11 Performance design and budget

COM calls cost milliseconds each, so per-paragraph property reads do not scale.
- Build `_visible_map` from **bulk reads**: one read of the document text and paragraph boundaries, plus **one pass** over `doc.Revisions` filtered to deletions. Avoid per-paragraph revision queries.
- The map is cached **only within one tool call** (C9), never across calls. A mutation builds the map once to resolve the index, and the post-edit local outline may reuse the pre-edit map adjusted for the edit, or rebuild once.
- **Budget (measured in E6):** `read_document(outline)` on a 1,000-paragraph document and `mutate_element` end-to-end each complete in about 2 s or less. If E6 misses the budget, stop (S7) and report options (for example, resolving indices with the bulk map but building the outline window from a handful of per-paragraph calls).

---

## 4. Phases

### Phase 0: Preparation
- [x] Create branch `feat/live-word-backend`.
- [x] Run the full suite; record baseline (expect 29 passing).
- [x] Inventory tests that assert on tool-level response text (3.3). Report the result and which pre-decided variant applies (prefix: always vs live-only; outline: both vs live-only). Stop (S1) only for conflicts the fallbacks do not cover.
- [x] Add `pywin32` to `requirements.txt` (win32 marker); install into `.venv`.
- [x] Register the `live` marker; confirm `pytest -m "not live"` runs the 29 tests unchanged.
- [x] Update `.gitignore` (3.1).
- [x] Create `src/live_engine.py` with: guarded imports, `HAS_COM`, `LiveWordError` plus the error-message dictionary (3.4), `_norm` (3.2), `_com_idx`, `_visible_map` (stub with the bulk-read signature), `find_open_document`, `pick_backend` (3.2), the **COM worker thread with per-call timeout and poisoned state** (3.8, C9), and the undo/track wrapper (3.6). No tool logic yet.
- [x] Unit tests (no Word needed) for the worker: timeout returns the dialog error, poisoned state fails fast, recovery clears the flag, no second COM thread is created.

**Gate 0:** 29 tests pass; `python -c "import src.server"` works with and without `pywin32` importable (simulate with `HAS_COM = False`); worker unit tests pass; chosen prefix/outline variants recorded in `.agent-info/docs/reviews/live-word-gate-0.md`. [PASSED]

### Phase 1: Spike (3 tools) and experiments

**[HUMAN] pre-step:** ask the user to start Word (non-elevated) and leave it running, ideally with no other documents open. The agent never launches Word (C4). Wait for confirmation before live tests.

**Experiments** (cap: 15 tool calls each; log to the experiments file; use only `documents/live_test.docx`):
- **E1 Revisions, tables, and numbering:** (a) do fully tracked-deleted paragraphs stay in `Paragraphs`; how does headless count them; (b) how does headless address table content and does excluding table paragraphs from `[p:N]` match it; (c) how does `Range.Text` behave with deleted text present, and do character offsets stay consistent around fields/hidden text; (d) does `Find` match hidden tracked-deleted text. Confirms or corrects 3.5 and 3.9.
- **E2 Backup copy:** can the on-disk file be copied while Word has it open? What does `doc.Saved` report after an agent edit?
- **E3 Lock-file naming:** observe the `~$` owner file for names of different lengths; fix the matching rule in 3.2 step 6.
- **E4 (optional) Multiple Word instances:** is enumerating the Running Object Table feasible for detecting documents in a second instance? If not cheap, document the limitation; the lock-file net still protects writes.
- **E5 Undo record:** confirm one custom undo record equals one Ctrl+Z, including when the call raises an error mid-way, and that mixed tracked/untracked sub-steps (3.6) still undo as one step.
- **E6 Performance:** generate a 1,000-paragraph test document (with a table and some tracked deletions); measure `_visible_map` build time and end-to-end `outline` and `mutate_element` against the 3.11 budget.
- **E7 Revision author:** does setting `Application.UserName` change the author recorded on new revisions? Is it restored reliably, including after a forced error? Decide keep/drop for 3.10 (do not simulate a hard crash; reason about it and note the risk).

**Implement live versions of:**
1. **`read_document`**: modes `outline`, `full`, `metadata` (`table` may raise `NotImplementedError("phase 2")`), using the final view (3.5) and the bulk-read map (3.11).
2. **`mutate_element`**: `insert_before`, `insert_after`, `replace`, `delete`; returns the local outline (3.3, subject to the Gate 0 variant). Runs with Track Changes forced on (C5). `replace` keeps the paragraph mark. **Style parity on insert:** `InsertParagraphBefore/After` inherits the neighboring paragraph's style in COM (a new paragraph after a Heading may become a Heading). Read how headless styles a new paragraph when no style is given, and have the live engine set the style explicitly to match. Support optional style in `content` if headless does. Add a test for replace inside a **table cell** (end-of-cell marker).
3. **`find_and_replace`**: body only, `match_case`, implemented per 3.9, real replacement count (additive on headless too), search-string limit of 255 enforced on both backends.

Wire these through `pick_backend()` in `server.py`, including the full detection order (3.2: Protected View, cloud, lock-file checks belong here, not Phase 3). The other five tools stay headless-only; if called while the file is open in Word, return the "not yet live-capable" error (3.4). Never write to a file Word has open.

**Mapping reference (COM):**

| Concept | COM |
|---------|-----|
| Paragraph N (0-based) | `doc.Paragraphs(_com_idx(doc, N))` |
| Insert before / after | `Range.InsertParagraphBefore()` / `InsertParagraphAfter()` then set style explicitly |
| Replace text keeping the mark | shrink `Range` end by 1 character, then set `.Text` |
| Delete | `Range.Delete()` |
| Outline level | `Paragraph.OutlineLevel` (10 = body text) |
| Find (locate) | `Range.Find.Execute(...)` without replace; set `Range.Text` per match |
| Undo grouping | `Application.UndoRecord` custom record |
| Track Changes | `Document.TrackRevisions` |
| Deleted spans | `doc.Revisions` filtered by `Type == wdRevisionDelete` (2) |

**Automated tests** (`@pytest.mark.live`; skip if Word is not running). The fixture creates `documents/live_test.docx`, opens it with `Documents.Open` in the already-running Word, and closes it with `SaveChanges=False` at teardown. It never touches any other open document.
- [ ] Detection returns the doc when open, `None` when closed; path variants (case, short name, mapped/UNC alias where available) resolve to the same document.
- [ ] Live outline equals headless outline on the same file, **including a document with a table and a document with a pre-existing tracked deletion**.
- [ ] Insert, replace, delete each yield the expected text at the expected index; after a delete the next indices match headless.
- [ ] Insert after a heading yields the same style as headless.
- [ ] Ten consecutive mutations using only the returned local outline hit the right paragraphs.
- [ ] `TrackRevisions` (and author, if enabled) restored after success **and** after a forced error.
- [ ] One mutation equals one undo step (`doc.Undo()` once restores the text).
- [ ] `find_and_replace`: counts match; a match inside tracked-deleted text is skipped; a replacement longer than 255 characters works on live; special characters in the search are escaped.
- [ ] Closing the doc makes the same call fall back to headless with no error.
- [ ] Protected View / cloud / lock-file cases return the messages in 3.4 (simulate where possible; otherwise mark as manual and list in the gate report).

**Gate 1 [HUMAN]:** ask the user to run this checklist with Word open on `documents/live_test.docx`:
1. An agent edit appears on screen within about 1 second.
2. Ctrl+Z undoes exactly one agent action.
3. The edit appears as a tracked change that can be accepted and rejected (and shows the agent author name if E7 was adopted).
4. After closing Word, the same tool call works headless.
5. A deliberately wrong index returns the clear out-of-range message, not a crash.
6. Typing nothing, run one call while a Word dialog (for example Save As) is open: the agent gets the dialog message within the timeout, and the next call after dismissing the dialog works.

If any item fails: S3, stop and report; no fixes until the user decides. Save the result to `.agent-info/docs/reviews/live-word-gate-1.md`.

### Phase 2: Port the remaining 5 tools (only after Gate 1)

**Before items 3 and 4, run E8:**
- **E8 (optional spike, cap 15 calls): `Range.InsertXML` parity.** For `compose_document` and `format_table_grid`, can the existing `oxml_engine` generate the OOXML fragment and be inserted with `Range.InsertXML` as one undo step, with output matching headless? Record formatting fidelity, undo behavior, and how it behaves with Track Changes on. If it works, port those two tools that way (smaller, closer to headless parity); if not, use the hand-port approach below.

One tool per commit, each with live tests. Order:

| Order | Tool | Live approach | Tracking (C5) |
|-------|------|---------------|---------------|
| 0 | `clean_structure` **dry_run** (both backends) | Additive `dry_run: bool = False`; when true, report what would change and modify nothing. Must exist **before** the live port of this tool. | n/a |
| 1 | `set_page_layout` | `Sections(i).PageSetup`; points (1 cm = 28.35 pt); all sections; swap width/height for landscape. | untracked |
| 2 | `update_styles` | `doc.Styles(name)` `.Font` / `.ParagraphFormat`. Missing style -> error listing available names. | untracked |
| 3 | `compose_document` | `Range.InsertAfter`, paragraph styles, table add (or E8 approach); one undo record for the whole batch. | tracked |
| 4 | `format_table_grid` | `Table.Borders`, header shading, `Rows(1).HeadingFormat`, `Rows.AllowBreakAcrossPages = False`, cell padding (or E8 approach). | untracked |
| 5 | `clean_structure` | Highest risk. One flag at a time (empty paragraphs, empty tables, break capping, color normalization). Never delete content-bearing elements. Run dry-run before each real run in tests. | deleting flags tracked; color normalization untracked (use `tracking()` sub-steps inside one undo record) |

Also implement `read_document` `table` mode.

**Gate 2 (comparison method):** the live engine does not save, so do **not** compare files. Run the styling pipeline from `.agents/rules/document-styling.md` (the four tool calls, then the metadata check) twice on identical input: once live on the open test doc, once headless on a closed copy. Use the preset file that **actually exists** in `.agents/presets/` (list the directory; do not assume `academic-id` vs `academic_id`). Compare: (a) `read_document(mode="metadata")` output from both backends, and (b) values read back from COM (margins, style fonts, table borders) against values read from the headless result via python-docx. Margins, fonts, and table borders must match. Note: because formatting tools run untracked live and untracked headless, the revision lists should be empty of formatting revisions on both sides; confirm that as part of the gate report.

### Phase 3: Hardening
- [ ] Real-dialog behavior of the worker timeout (3.8) verified with the user ([HUMAN], already part of Gate 1 item 6; repeat for a long-running call).
- [ ] Read-only and document-protection detection (`doc.ReadOnly`, `doc.ProtectionType`) up front (3.4). Protected View, cloud, and lock-file detection already exist from Phase 1.
- [ ] `find_and_replace`: also search headers, footers, footnotes (other story ranges), using the same locate-then-set and deleted-span skip; real count across all searched ranges.
- [ ] `find_and_replace` `dry_run` on both backends (additive).
- [ ] Automatic backups per 3.7 on both backends, with `BACKUP_KEEP` pruning; response notes when a backup reflects last-saved state.

**Gate 3:** all live and headless tests pass; dry-run verified on both tools. Report actual passed / skipped / failed counts.

### Phase 4: Rules, skill, docs, PR

#### 4a. Rules (create after Gate 3)
Goal: a **small, always-correct** rule. Rules hold invariants; how-to material belongs in the skill.

1. `ls .agents/rules/` and read two sibling rules to match frontmatter style and tone.
2. Create `.agents/rules/live-word-mode.md` from **Appendix A**, choosing the wording variants recorded at Gate 0. Adapt only formatting to match siblings. Keep it to about 20 lines. Do not move playbook content into it.
3. Amend `.agents/rules/document-styling.md`: append exactly one new step 4 (text in section 10). Change nothing else in that file.
4. **Truth check** before committing: for each sentence in the new rule, verify it against the implemented behavior (prefix on all 8 tools or live-only per variant, outline returned by `mutate_element` per variant, the tracking policy per tool, the listed error situations). Fix the rule if reality differs; do not "fix" the code to match the rule without asking.
5. No change to `git-usage`, `gh-usage`, `info-dir`, `problem-resolve`.

#### 4b. Skill (create after Gate 3, so it documents verified behavior)
Why a skill: the rule is loaded when relevant but must stay tiny. The skill is the on-demand playbook (workflow, reading semantics, error triage, pipeline recipe). Each fact lives in exactly one place: invariants in the rule, procedures in the skill, one-sentence hints in tool docstrings.

1. **Location check (stop if unclear):** `ls .agents/` and look for an existing skills directory and a sibling skill to copy conventions from. If `.agents/skills/` exists, use `.agents/skills/live-word-editing/`. If there is no precedent or you are unsure of the convention, stop (S9) and ask the user where skills live. Do not guess.
2. Create `SKILL.md` from **Appendix B** plus `references/error-catalog.md`.
   - Frontmatter: `name` (must equal the directory name, kebab-case) and `description`.
   - **Description is the trigger.** State what the skill does and when to use it, naming concrete phrases: "open in Word", "live", "track changes", "undo", "edit this .docx", "format the document". Be a little pushy so it does not under-trigger. Keep it under about 600 characters.
   - Body: imperative steps, with the reason for each non-obvious step. Target under 150 lines. Link to `references/` with a note on when to open each file.
   - `references/error-catalog.md`: **generate it from the error-message dictionary in `live_engine.py`** (script or copy), one row per message: situation, exact text, what to tell the user, what not to do. Do not retype messages by hand.
3. No new behavior in the skill. If the skill needs to say something the code does not do, the code or the skill is wrong; stop and ask (S7).
4. **Validate [HUMAN]:** ask the user to run the Appendix C prompts in fresh agent sessions and report which triggered the skill and whether the agent behaved correctly. At most two revision rounds, adjusting the description first, the body second.
5. Save the validation result to `.agent-info/docs/reviews/live-word-skill-validation.md`.

#### 4c. Docs
- [ ] `WORKSPACE.md`: new backend, routing diagram (3.2), env vars (`WORD_BACKEND`, `LIVE_CALL_TIMEOUT`, `LIVE_AUTHOR`), the `[backend: ...]` line (per chosen variant), tracking policy table (C5), Windows-only note, limitations (single-instance detection, last-saved-state backups, replace takes first-character formatting), where the rule and skill live.
- [ ] Replace any `file:///c:/Users/...` links with relative links (C7).
- [ ] Reword the testing section: report the actual passing count; no "100% coverage" claims.
- [ ] `README.md` "Live mode" section: setup, `WORD_BACKEND`, content edits tracked / formatting edits untracked, one undo per call, "start Word before the agent; do not type in the document while a call is running".
- [ ] One-sentence live-behavior note in the docstrings of the three live-capable tools (later five as ported). Keep short.

#### 4d. PR
- [ ] Push the branch and open a draft PR per `gh-usage`. Do not merge.

**Gate 4 [HUMAN]:** user reviews docs, rule, skill, and validation notes; draft PR open. Merging is the user's decision.

---

## 5. Known pitfalls (read before coding)

1. **Index drift.** Insert/delete shifts later indices. Always return and use the local outline; never cache `[p:N]` across mutations.
2. **COM threading and hangs.** FastMCP may use worker threads. Follow C9 and 3.8; COM objects are apartment-bound, and a blocked call cannot be cancelled.
3. **Elevated Word.** Attach can fail. The lock-file net (3.2 step 6) turns this into a clear refusal for writes.
4. **Paragraph mark.** `Paragraph.Range.Text` ends with `\r`; cell paragraphs end with `\r\x07`. Strip for display; never include in replace content.
5. **Tables.** Table cell paragraphs are in `doc.Paragraphs` but are **not** `[p:N]` paragraphs (3.5). If outline equality on a table document fails, stop (S2).
6. **Find.** 255-character limit on the search string; `^` and other specials have meaning in Word's Find, so escape them. Hidden tracked-deleted text may match (3.9).
7. **Tracked deletions** stay in the COM model. Use the visible map (3.5) for every read, count, and index. Never "fix" this by toggling view settings (C13).
8. **Do not save for the user** (C12).
9. **Replace loses run-level formatting** inside the replaced range (it inherits the first character's). Note it in the skill; do not try to fix it in this project.
10. **Performance.** Per-paragraph COM property reads are slow. Use bulk reads (3.11).
11. **Inserted paragraphs inherit style.** Set the style explicitly to match headless.
12. **Revision author and Track Changes are user state.** Restore them in `finally`, always.

---

## 6. Test and verification summary

| Layer | Command | Must pass at |
|-------|---------|--------------|
| Headless (existing + new) | `.venv\Scripts\python.exe -m pytest tests/ -m "not live" -v` | every commit |
| Live (needs Word running) | `.venv\Scripts\python.exe -m pytest tests/ -m live -v` | gates 1, 2, 3 |
| Round-trip sanity | Re-open **headless-produced** results with `python-docx`; optionally convert to PDF with LibreOffice if installed. Never open the live document this way. | Phase 2 and 3 |

Report **actual** counts (passed / skipped / failed). Never write "100% coverage".

---

## 7. Fallback

If the live approach is abandoned: keep `src/live_engine.py` on the branch, do not merge, and implement the cheap alternative: each agent run saves to `<n>_vN.docx` and the user opens the new file. No COM, no lock problem.

---

## 8. Definition of done

- [ ] Gates 0-4 passed (Gate 1 and Gate 4 confirmed by the user).
- [ ] 29 original tests pass unmodified; new headless and live tests added.
- [ ] No tool name or existing parameter changed (only additive optional params like `dry_run`).
- [ ] Works with pywin32 absent, Word closed, and Word open.
- [ ] Writes are refused (not attempted) when the file is Protected View, cloud-hosted, or lock-file-flagged.
- [ ] A stuck Word call yields a clear error within the timeout and never wedges later calls.
- [ ] Performance budget (3.11) met on the 1,000-paragraph test document.
- [ ] Rule created and `document-styling` amended; rule statements truth-checked.
- [ ] Skill created, error catalog generated from code, validation notes saved.
- [ ] Docs updated, relative links, honest test claims.
- [ ] Draft PR open, not merged.

---

## 9. Stop conditions (stop and report)

| ID | Trigger |
|----|---------|
| S1 | An original test fails or would need editing, **and** the pre-decided fallbacks in 3.3 do not cover it. |
| S2 | Live and headless paragraph numbering disagree on the same document and cannot be reconciled without changing headless behavior. |
| S3 | Any Gate 1 item fails (stop on first failure; no fixes). |
| S4 | A change would alter a tool name, an existing parameter name, or the index convention. |
| S5 | Testing would need a document other than `documents/live_test.docx`. |
| S6 | Word shows an unexpected dialog or hangs during live tests (outside the deliberate dialog check in Gate 1 item 6). |
| S7 | Any other error, failed test, unexpected output, missed performance budget, or experiment result (or experiment call cap reached unresolved) that contradicts this plan. |
| S8 | A command would hit a confirmation gate in `git-usage` or `gh-usage`. |
| S9 | The skills directory convention is unclear, or a rule/skill would duplicate or contradict each other. |

**Report format** (matches `problem-resolve`), under 15 lines, then wait:
1. **What failed:** exact command, action, or error message.
2. **Context:** relevant context or root-cause hypothesis.
3. **Paths forward:** options or questions, with your recommendation.

---

## 10. Rule and skill directives (summary; details in Phase 4a/4b)

| Artifact | Purpose | Size target | Source |
|----------|---------|-------------|--------|
| Rule `live-word-mode` | Invariants the agent must always respect in live mode | about 20 lines | Appendix A |
| Skill `live-word-editing` | On-demand playbook and error triage | under 150 lines + `references/` | Appendix B |
| Tool docstrings | One sentence each on live behavior | 1 sentence | Phase 4c |

**Exact amendment to `.agents/rules/document-styling.md`:** append as a new step 4 (after "Verify output metadata..."):

> 4. If the target `.docx` is open in Word, the tools run live (content edits appear as tracked changes, formatting edits are untracked, each call is one undo step). Do not close, save, or reopen the document, and note the backend (the `[backend: ...]` line, or its absence if the live-only prefix variant is in use) from the first tool response in your summary. For live-mode procedure and error handling, see the `live-word-editing` skill.

(If the skill is not created, drop the last sentence.)

---

## Appendix A: Rule draft (`.agents/rules/live-word-mode.md`)

Choose the wording variants recorded at Gate 0 (bracketed alternatives below).

````markdown
---
trigger: model_decision
description: When editing, formatting, or inspecting a .docx with the word-docx MCP tools while the file may be open in Microsoft Word
---

# Live Word Mode Rule

1. **Backend is automatic.** [Variant A: The first line of every `word-docx` tool response shows `[backend: live]` or `[backend: headless]`.] [Variant B: Live responses begin with `[backend: live]`; no such line means headless.] Never switch backends or set `WORD_BACKEND` yourself.
2. **When live:** content edits appear in the open Word window as tracked changes; formatting and layout edits are applied untracked. Each tool call is one undo step. Never close, save, or reopen the document. Never touch the same file through any other method (python-docx, scripts, shell).
3. **Indices:** after any `mutate_element`, [Variant A: use the returned local outline for the next `[p:N]`.] [Variant B: when live, use the returned local outline for the next `[p:N]`; when headless, re-read the outline.] Do not reuse earlier indices. Re-read only if your target is not in the returned outline.
4. **Errors:** if a tool reports the file is read-only, protected, in Protected View, cloud-hosted, appears open but cannot be attached, or that Word looks blocked by a dialog, follow `problem-resolve`: stop, relay the tool's message, wait.
5. **How-to:** for workflow, reading semantics, and error triage, use the `live-word-editing` skill.
````

(Drop rule 5 if the skill is not created. Verify every sentence against implemented behavior before committing.)

## Appendix B: Skill draft (`.agents/skills/live-word-editing/SKILL.md`)

Write this after Gate 3. Replace bracketed parts with verified facts from the code. Keep the structure.

````markdown
---
name: live-word-editing
description: Procedure for editing, formatting, or inspecting a .docx with the word-docx MCP tools when the file may be open in Microsoft Word (live mode via COM). Use whenever the user mentions a Word document being open, watching changes live, track changes, undo of agent edits, or asks to edit, restyle, or format a .docx with these tools, even if they do not say "live". Covers backend detection, safe mutation workflow, reading semantics for tracked deletions and tables, error triage, and the styling pipeline in live mode.
---

# Live Word editing

Invariants (never close/save/reopen, use returned outline, stop on listed errors) are in the `live-word-mode` rule. This skill is the procedure.

## Workflow
1. Make the first tool call a read (`read_document`, `outline`). Note the backend [line / absence of line]. Why: it tells you whether edits appear in Word and whether undo is available.
2. Pick the target paragraph from the outline. Indices are 0-based and count visible body paragraphs only (tracked-deleted paragraphs and table-cell paragraphs are not `[p:N]`).
3. Mutate with `mutate_element` or `find_and_replace`. Read the returned local outline and take the next index from it. Why: indices shift after inserts and deletes.
4. After a batch of edits, tell the user the backend used, how many edits were made, and (if live) that each can be undone with one Ctrl+Z; content edits can also be reviewed as tracked changes [by author `Agent`, if E7 adopted].
5. Do not ask the user to close the file. Closing simply switches the next call to headless.

## Reading semantics
- Reads show the final view: tracked-deleted text is hidden and fully deleted paragraphs are skipped in numbering. [confirm against E1 result]
- A tracked change you made is still reversible by the user; do not "clean up" after yourself.

## Tracking policy
- Tracked: `mutate_element`, `find_and_replace`, `compose_document`, deleting flags of `clean_structure`.
- Untracked (one Ctrl+Z still undoes them): `set_page_layout`, `update_styles`, `format_table_grid`, color normalization. [verify against C5 implementation]

## Styling pipeline in live mode
Follow `document-styling`. The same four calls and the metadata check apply. [list which tools are live-capable at this commit]
- `clean_structure`: run with `dry_run=True` first, show the user the summary, then run for real.

## When something goes wrong
Stop per `problem-resolve`. Look up the message in `references/error-catalog.md` for what to tell the user and what not to do. Open that file only when a tool returns an error. If a call times out because Word may be blocked by a dialog, tell the user to dismiss it; do not retry.

## Known limits
- Windows desktop Word only; one Word instance is detected. [update with E4 outcome]
- Replacing text takes on the formatting of the first replaced character.
- Backups reflect the last saved state of the file; the newest 20 per file are kept.
````

`references/error-catalog.md` must be **generated** from the error dictionary in `src/live_engine.py` (columns: situation, exact message, what to tell the user, what not to do). Add a short table of contents if it exceeds 100 lines.

## Appendix C: Skill validation prompts (run by the user in fresh sessions)

Should **trigger** the skill:
1. "My report.docx is open in Word. Change the heading on page 2 to 'Results' and I'll watch it happen."
2. "Apply the academic preset to thesis.docx. I have it open, so don't make me close it."
3. "Undo-friendly edit: replace 'Q3' with 'Q4' everywhere in the open document."
4. "The agent said the file is in Protected View. What now?"

Should **not** trigger the skill:
5. "Commit my current changes with a conventional message."
6. "Write a Python script that reads a .docx with python-docx and prints headings." (no MCP tools, no live Word)

Record for each: triggered (yes/no), first tool call made, whether the agent respected the rule (no save/close, used the returned outline), and any confusion. If 5 or 6 trigger, narrow the description. If 1-4 do not trigger, add the missing phrases to the description.
