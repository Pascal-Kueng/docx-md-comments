# AGENTS.md

AI maintainer guide for `docx-comments-roundtrip`.

## Mission

Preserve Word comment integrity through `.docx <-> .md` roundtrip.

Core requirements:

- Keep comment anchors and text intact.
- Reconstruct comments from markdown spans.
- Flatten threaded replies into root comment bodies in Word output.
- Preserve root comment state (`active` vs `resolved`).
- Avoid duplicate standalone reply comments after flattening.

## Repository map

- `docx-comments`: main converter (single Python script).
- `docx2md`, `md2docx`: thin wrapper scripts.
- `tests/test_roundtrip_example.py`: fixture-backed roundtrip parity test.
- `tests/test_roundtrip_edges.py`: synthetic edge-case roundtrip tests.
- `tests/helpers/docx_inspector.py`: reads DOCX XML into comparable snapshots.
- `tests/helpers/markdown_inspector.py`: reads markdown comment spans into snapshots.
- `tests/helpers/diagnostics.py`: writes failure bundles and diffs.
- `Makefile`: test entrypoints and manual artifact workflow.

## Comment model and invariants

### Word package files in scope

- `word/comments.xml`: primary comment nodes and text.
- `word/commentsExtended.xml`: thread links and resolved state (`w15:done`).
- `word/commentsIds.xml`: optional helper mapping (order/para IDs in some files).
- Story XMLs (`document.xml`, headers, footers, footnotes, endnotes): anchors/references.

### Markdown span contract

Expected span metadata:

- `.comment-start`: `id`, optional `author`, `date`, `parent`, `state`.
- `.comment-end`: `id`.

`state` must normalize to:

- `resolved`
- `active` (default for missing/invalid values)

## Roundtrip invariants that must hold

1. Original root comment IDs survive roundtrip with same root order.
2. Child thread comments do not survive as standalone comments.
3. Flattened root comment text contains reply blocks with author/date separators.
4. Every roundtrip root has anchor/start/end/reference IDs in story XML.
5. Anchor span text for roots remains equivalent (normalized comparison).
6. `commentsExtended.xml` exists in roundtrip output and root states are preserved.
7. `commentsIds.xml` does not reintroduce orphaned child thread artifacts.
8. Placeholder shape image markdown artifacts are removed.

## High-risk areas and regressions to avoid

Do not reintroduce these failure patterns:

1. Positional or overlap-only thread inference.
- Use ID and metadata mapping (`parentId`, `paraIdParent`, markdown `parent`).
- Overlapping/interleaved anchors are valid and must not be merged heuristically.

2. Synthetic title/header injection.
- Never inject heading/title content from `docProps/core.xml` to "help" anchors.
- This previously caused duplicate headings and first-line instability.

3. Partial pruning of child artifacts.
- If flattening removes child comments, prune child nodes across:
  - story anchors/references
  - `comments.xml`
  - `commentsExtended.xml`
  - `commentsIds.xml` (when present)

4. Over-broad markdown parent assignment.
- Only trust parent-child links for IDs confirmed as real `.comment-start` spans.

5. Losing first-character anchors.
- Comments that start at first document character/heading must survive.
- This is a known fragile area and must always be covered by tests.

## Operational constraints

- This tool is often stow-managed and symlinked into `~/.local/bin`.
- Never write runtime artifacts beside the script location.
- Keep temp files near the input/output document path.
- Use `TemporaryDirectory(...)` and clean up automatically.
- Keep `sys.dont_write_bytecode = True` behavior to avoid `__pycache__` sprawl.

## Required verification workflow

Run before merging behavior changes:

1. `make test`
- Runs `make roundtrip-example` first and keeps:
  - `artifacts/out_test.md`
  - `artifacts/out_test.docx`
- Then runs the unittest suite.

2. `make test-roundtrip`
- Runs roundtrip-focused tests only.

3. Manual inspection in Word for fixture output:
- Comment count parity for roots.
- No dropped first comment anchor.
- Replies flattened into roots.
- Resolved/active statuses preserved.

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
