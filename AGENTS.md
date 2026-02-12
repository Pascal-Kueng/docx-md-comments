# AGENTS.md

AI maintainer guide for `docx-comments-roundtrip`.

## Mission

Preserve Word comment integrity through `.docx <-> .md` roundtrip.

Core requirements:

- Keep comment anchors and text intact.
- Reconstruct comments from markdown spans.
- Reconstruct threaded replies as native Word comment threads in output.
- Preserve comment state (`active` vs `resolved`) for roots and replies.
- Preserve reply comment state and parent linkage.
- Match Microsoft Word behavior (OnlyOffice parity is useful but not sufficient).

## Repository map

- `docx-comments`: main converter (single Python script).
- `docx2md`, `md2docx`: thin wrapper scripts.
- `tests/test_roundtrip_example.py`: fixture-backed roundtrip parity test.
- `tests/test_roundtrip_edges.py`: synthetic edge-case roundtrip tests.
- `tests/test_markdown_attr_transforms.py`: regression tests for AST-based markdown attr transforms.
- `tests/helpers/docx_inspector.py`: reads DOCX XML into comparable snapshots.
- `tests/helpers/markdown_inspector.py`: reads markdown comment spans into snapshots.
- `tests/helpers/diagnostics.py`: writes failure bundles and diffs.
- `Makefile`: test entrypoints and manual artifact workflow.

## Comment model and invariants

### Word package files in scope

- `word/comments.xml`: primary comment nodes and text.
- `word/commentsExtended.xml`: thread links and resolved state (`w15:done`).
- `word/commentsIds.xml`: mapping between thread para IDs and durable IDs.
- `word/commentsExtensible.xml`: durable ID extension entries (with `dateUtc`).
- `word/people.xml`: author records and `presenceInfo`.
- Story XMLs (`document.xml`, headers, footers, footnotes, endnotes): anchors/references.

### Markdown span contract

Expected span metadata:

- `.comment-start`: `id`, optional `author`, `date`, `parent`, `state`.
- `.comment-end`: `id`.

Additionally accepted on md->docx input:

- Shorthand milestone tokens in prose:
  - `///c1.START///` (start), `///c1.END///` (end) as canonical inner form
  - optional highlighted wrapper: `==///c1.START///==` / `==///c1.END///==`
  - numeric aliases: `///C1.START///` / `///C1.END///`
  - whitespace-tolerant variants: `/// c1 . start ///`
- In docx->md output, each root comment gets a blockquote callout inserted directly after the block that contains its root end marker (not batched at document end):
  - `> [!COMMENT <id>: <author> (<state>)]`
  - `> <!--CARD_META{#<id> "author":"...","date":"...","state":"..."}-->`
  - `> <root comment body>`
- Replies are nested inside the root callout as nested blockquotes:
  - `> > [!REPLY <id>: <author> (<state>)]`
  - `> > <!--CARD_META{#<id> "author":"...","date":"...","parent":"<root-id>","state":"..."}-->`
  - `> > <reply body>`
- Only root comments keep inline milestone markers in prose; reply anchors are carried by nested reply callouts.
- These are AST-normalized to canonical span markers before extraction.

Internal transport metadata (docx->md->docx only; must be stripped before pandoc md->docx):

- `.comment-start`: `paraId`, `durableId`, `presenceProvider`, `presenceUserId`.

Canonical span-ID rule:

- Keep comment IDs in explicit attribute form (`id="..."`) for both `.comment-start` and `.comment-end`.
- Do not rely on Pandoc identifier shorthand (`{#id ...}`) for comment markers.

`state` must normalize to:

- `resolved`
- `active` (default for missing/invalid values)

## Roundtrip invariants that must hold

1. Original root comment IDs survive roundtrip with same root order.
2. Child thread comments survive as threaded reply comments with valid `parentId`/`paraIdParent`.
3. Every roundtrip root has anchor/start/end/reference IDs in story XML.
4. Anchor span text for roots remains equivalent (normalized comparison).
5. `commentsExtended.xml` exists in roundtrip output and states are preserved.
6. `commentsIds.xml` and `commentsExtensible.xml` cover all roundtrip comments.
8. `commentsExtensible.xml` and `people.xml` are present when needed for state fidelity in Word.
9. For multi-paragraph comments, roundtrip thread `paraId` equals the last comment paragraph `paraId`.
10. Placeholder shape image markdown artifacts are removed.

## High-risk areas and regressions to avoid

Do not reintroduce these failure patterns:

1. Positional or overlap-only thread inference.
- Use ID and metadata mapping (`parentId`, `paraIdParent`, markdown `parent`).
- Overlapping/interleaved anchors are valid and must not be merged heuristically.

2. Synthetic title/header injection.
- Never inject heading/title content from `docProps/core.xml` to "help" anchors.
- This previously caused duplicate headings and first-line instability.

3. Breaking thread metadata parity.
- When writing threaded output, ensure children are present in:
  - `comments.xml` with `parentId`
  - `commentsExtended.xml` with `paraIdParent`
  - `commentsIds.xml` / `commentsExtensible.xml` with durable mappings

4. Over-broad markdown parent assignment.
- Only trust parent-child links for IDs confirmed as real `.comment-start` spans.

5. Losing first-character anchors.
- Comments that start at first document character/heading must survive.
- This is a known fragile area and must always be covered by tests.

6. Wrong thread paraId selection for multi-paragraph comments.
- Word binds `commentsExtended/commentsIds` to the thread paraId.
- In multi-paragraph comments this is the last paragraph `w14:paraId`, not the first.
- Using the first paragraph paraId leads to partial "resolved" restoration in Word.

7. Dropping `people.xml` presence metadata.
- Preserve `w15:presenceInfo` (`providerId`, `userId`) per author when present in source.
- Missing presence metadata can cause Word UI/status behavior drift despite seemingly correct XML counts.

8. Trusting non-Word viewers as source of truth.
- OnlyOffice may show resolved states even when Word does not.
- Use Word as the acceptance target for resolved/active parity.

9. Regex-only mutation of comment span attributes.
- Do not mutate `.comment-start` metadata with raw markdown regex replacement.
- Use Pandoc JSON AST traversal so only real comment spans are changed and code/prose literals are untouched.
- Keep regex-based repair for unbalanced/nested end markers, because malformed markers may not parse as spans before repair.

10. Dropping marker normalization after AST re-serialization.
- After AST markdown re-emit, run end-marker repair/normalization again before md->docx conversion.
- This prevents missing `commentRangeEnd` / `commentReference` regressions from nested wrappers.

11. Over-tolerant milestone marker parsing.
- Milestone parsing must stay strict: only `///<id>.START|END///` (plus optional `==...==` wrapper and compatibility for `///C<digits>.START|END///`) with controlled ID charset.
- Do not match arbitrary prose fragments, or false positives will create phantom comments.
- Do not silently auto-repair missing START/END marker pairs in md->docx; fail fast with actionable diagnostics instead.
- Reject one-sided wrapper forms (e.g., `==///C1.START///` or `///C1.END///==`) with line-specific errors.

## Operational constraints

- This tool is often stow-managed and symlinked into `~/.local/bin`.
- Never write runtime artifacts beside the script location.
- Keep temp files near the input/output document path.
- Use `TemporaryDirectory(...)` and clean up automatically.
- Keep `sys.dont_write_bytecode = True` behavior to avoid `__pycache__` sprawl.
- Validate `pandoc` availability/version at runtime and fail early with clear errors.

## Required verification workflow

Run before merging behavior changes:

1. `make test`
- Runs `make roundtrip-example` first and keeps:
  - `artifacts/out_test.md`
  - `artifacts/out_test.docx`
- Then runs the unittest suite.

2. `make test-roundtrip`
- Runs roundtrip-focused tests only.

2.1 Targeted AST transform checks:
- `python3 -m unittest -q tests.test_markdown_attr_transforms`
- Verifies code/literal fake marker text is not modified while real comment spans are.

3. Manual inspection in Word for fixture output:
- Comment count parity for all comments.
- No dropped first comment anchor.
- Replies appear as native nested threads.
- Resolved/active statuses preserved (roots and replies).

4. If tests fail, inspect the generated `failure_bundle` path from test output:
- `original_snapshot.json`
- `markdown_snapshot.json`
- `roundtrip_snapshot.json`
- `command_logs.json`
- `mismatch_report.txt`

## Change checklist for future agents

When changing comment logic, update both converter and tests in the same PR:

1. Update parser/writer behavior in `docx-comments`.
2. Extend inspector snapshots if new invariants are introduced.
3. Add/adjust edge cases in `tests/test_roundtrip_edges.py`.
4. Ensure fixture parity assertions in `tests/test_roundtrip_example.py` still pass.
5. Keep diagnostics actionable (include IDs and text diffs).
6. Avoid broad refactors without test expansion for first-char and overlapping cases.

## Current design status (Feb 2026)

- Root comment state is parsed from Word package metadata and roundtripped through markdown.
- Root/reply threading remains ID-based and is reconstructed package-wide.
- State reconstruction writes:
  - `commentsExtended.xml` (`w15:done`)
  - `commentsIds.xml` (`paraId` <-> `durableId`)
  - `commentsExtensible.xml` (`durableId`, `dateUtc`)
  - `people.xml` (`w15:person` + optional `w15:presenceInfo`)
- Thread paraId mapping uses the comment thread paraId (last paragraph `w14:paraId` for multi-paragraph comments).
- Tests now enforce:
  - resolved-count parity
  - state-supporting part/relationship/content-type presence
  - no invalid comment-level state attrs in `comments.xml`
  - presenceInfo preservation per author when present in source
  - thread paraId alignment to last paragraph paraId for multi-paragraph comments
  - parent-map parity (`parentId`) and `paraIdParent` linkage for replies
- Comment metadata transport now uses Pandoc JSON AST mutation for `.comment-start` attrs (not regex text replacement).
- AST re-serialization preserves user-requested writer when available (`-t/--to` passthrough).
- Comment marker IDs are normalized to `id="..."` attributes after AST operations to keep downstream marker repair stable.
- md->docx now normalizes shorthand milestone tokens (`///<id>.START|END///`, optional `==...==` highlight wrapper, spacing tolerant) to canonical comment spans via Pandoc AST before comment extraction.
- md->docx validates marker integrity before conversion and aborts with clear line-specific errors if root marker pairs are missing/duplicated/unbalanced.
- Current markdown card transport format is `COMMENT/REPLY` blockquote callouts with inline `CARD_META` HTML comments (no `CARD_START` markers, no fenced Div wrappers, no backward-compat parsing path).
