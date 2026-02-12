# docx-comments-roundtrip

Bidirectional `.docx <-> .md` conversion with Word comment roundtrip support.

## Scope

This tool is intentionally focused on comment-safe conversion:

- `docx -> md`: keeps comment anchors as markdown spans (`.comment-start` / `.comment-end`) with metadata.
- `md -> docx`: reconstructs comments from markdown spans.
- `md -> docx`: also accepts shorthand milestone markers in prose and normalizes them to comment spans:
  - `///c1.START///` / `///c1.END///` (canonical)
  - numeric aliases like `///C1.START///` / `///C1.END///`
  - whitespace-tolerant forms like `/// c1 . start ///`
  - backward-compatible legacy forms: `DC_COMMENT(c1.s)` and `{[c1.s]}`
- `docx -> md`: emits milestone markers in prose for root comments and inserts matching `.comment-card` blocks right after the block that closes the root marker span.
- Each `.comment-card` is human-readable:
  - blockquote callout header `>[!COMMENT <id>: <author> (<state>)]`
  - hidden JSON transport comment `<!--{...}-->` for robust roundtrip-only fields
  - comment body text
- Thread replies are represented as nested `.comment-card .comment-reply-card` blocks under their root card (not as extra inline prose markers).
- Threaded Word replies are flattened into the parent comment body for stable roundtrip output.
- Comment state is preserved for roots (`active` vs `resolved`).
- Known pandoc shape-placeholder image artifacts are filtered out of markdown.

This is not a full-fidelity Word layout converter. The primary goal is preserving comment structure and text through roundtrip.

## Installation

### Prerequisites

- Python 3.10+
- `pandoc` on `PATH`

### Local use from this repository

Run directly:

```bash
./docx-comments --help
```

Optional wrappers:

- `./docx2md`
- `./md2docx`

### Optional global install

If you manage dotfiles with GNU Stow, symlink these scripts into your `PATH` (for example `~/.local/bin`), keeping the repo as the source of truth.

## Usage

### Main CLI (`auto` mode by extension)

```bash
./docx-comments input.docx -o output.md
./docx-comments input.md -o output.docx
```

### Explicit mode wrappers

```bash
./docx2md input.docx -o output.md
./md2docx input.md -o output.docx
```

### Extra pandoc arguments

Unknown CLI flags are forwarded to pandoc:

```bash
./docx-comments input.docx -o output.md --reference-doc=template.docx
```

## Testing and Inspection

Run the full suite:

```bash
make test
```

`make test` always runs an example roundtrip first and keeps outputs for manual inspection:

- `artifacts/out_test.md`
- `artifacts/out_test.docx`

Run only roundtrip-focused tests:

```bash
make test-roundtrip
```

Run only the example conversion (no unittest assertions):

```bash
make roundtrip-example
```

Remove manual artifacts:

```bash
make clean-roundtrip-example
```

On failures, tests write a `failure_bundle/` directory in a temp case folder with snapshots and diffs (`original`, `markdown`, `roundtrip`, command logs, and mismatch report).
