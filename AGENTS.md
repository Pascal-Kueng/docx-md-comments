# AGENTS.md

This file documents implementation history, failed approaches, and guardrails so future agents can continue without reintroducing regressions.

## Project scope

This repository maintains a practical converter for:

- `.docx -> .md` with comments preserved as markdown spans.
- `.md -> .docx` with comments reconstructed.
- Threaded Word comments flattened into the parent comment body for stable roundtrip behavior.

## Environment and deployment assumptions

- The original script lived in a GNU Stow-managed dotfiles tree and is symlinked into `~/.local/bin`.
- Writing runtime artifacts beside the stowed script is unacceptable.
- Temporary artifacts must be created near the document being converted and cleaned up automatically.

## What was implemented

1. Unified CLI and wrappers
- Main script: `docx-comments` with auto mode detection by extension.
- Wrappers: `docx2md`, `md2docx`.

2. Thread flattening model
- Parent/child relationships are derived from comment metadata (`comments.xml`, `commentsExtended.xml`, `commentsIds.xml`) and markdown attributes.
- Flattening output format embeds replies in parent text using a visual separator:
  - `---`
  - `Reply from: <Author> (<Timestamp>)`
  - `---`

3. Reverse conversion hardening
- Only root comments are rewritten in `comments.xml`.
- Child comment anchors/references and child comment nodes are pruned from package XML to avoid duplicate standalone comments.
- Extension/id companion files are also pruned for consistency.

4. Placeholder image artifact filtering
- Pandoc may emit tiny shape placeholders like `![](media/image1.png "Shape"){...}` and even `None.` lines.
- These placeholders are filtered out before writing markdown and before md->docx conversion.
- Unreferenced extracted media files are pruned.

5. Comment text fidelity improvements
- Inline traversal was corrected for rich inline node types (`Strong`, `Emph`, `Span`, `Quoted`, `Cite`, `Link`, etc.) so comment text is not silently dropped.
- Line breaks are preserved as plain text paragraphs when reconstructing Word comments.

6. Cache/temp hygiene
- `sys.dont_write_bytecode = True` prevents `__pycache__` in stow-managed script folders.
- `TemporaryDirectory(...)` is used so temporary conversion directories are cleaned up even on failure.

## What did NOT work (and why)

1. Positional/overlap-based thread inference
- Failed on overlapping or interleaved comment spans.
- Incorrectly merged independent threads.
- Replaced with ID-based parent mapping.

2. Early flattening variants produced duplicate comments
- Replies appeared both inside the flattened parent and as standalone comments.
- Root cause: child comments were not fully removed from all Word package internals.
- Fixed by pruning child artifacts from story XML + comments XML + extension/id files.

3. Title heading injection from `docProps/core.xml`
- Attempted fix for missing title comments added synthetic markdown heading.
- Side effects:
  - duplicated top heading
  - unstable first-line behavior
- Fully reverted. Do not reintroduce title injection unless it is proven safe with comprehensive regression tests.

4. Over-broad parent assignment from markdown metadata
- Some IDs were treated as threaded without being confirmed as real `comment-start` spans.
- Could cause root misclassification and apparent “missing parent comment”.
- Fixed by validating that both child and parent IDs were actually seen as parsed comment-start spans.

## Known important edge cases

Always test these before merging changes:

1. Comment starts at first character of document / heading line.
2. Nested thread: root + multiple replies.
3. Overlapping/interleaved comment spans.
4. Single non-threaded comments still roundtrip unchanged.
5. Long comments with bold/emphasis and manual line breaks.
6. Documents containing real images vs shape-placeholder artifacts.
7. Comments in title/heading regions.
8. Roundtrip parity: no missing original root comments, no new unintended root comments.

## Verification protocol (minimum)

For a representative fixture document:

1. Convert `docx -> md`.
2. Convert `md -> docx`.
3. Compare comment IDs and parent/root structure between original and roundtrip outputs.
4. Confirm flattened mode expectations:
   - All original roots still present.
   - Child comments not present as standalone comments.
   - Reply content embedded in root body with author/date markers.
5. Inspect first heading/title area to ensure no duplicated heading regression.
6. Ensure no runtime artifacts are left in the repository/stow directory.

## Operational guidance for future agents

- Prefer ID-based reconstruction logic over text-span heuristics.
- Avoid any “helpful” synthetic content insertion (especially title/header injection) unless backed by explicit tests.
- Keep conversion temp data close to input/output files and ephemeral.
- If behavior changes, include before/after comment parity outputs in commit notes.
- When in doubt, prioritize preserving original root comments and preventing duplicate standalone replies.
