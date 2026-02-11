# docx-comments-roundtrip

Bidirectional `.docx <-> .md` conversion with Word comment preservation and robust thread flattening.

## What this project does

- Converts Word documents to Markdown while keeping comment markers and metadata.
- Converts Markdown back to Word and restores comments.
- Flattens threaded comments into the parent comment body in Word output.
- Preserves line breaks/paragraph-like structure in flattened comments.
- Filters known pandoc shape-placeholder image artifacts that break roundtrip.

## Scripts

- `docx-comments`: main converter with auto mode detection by extension.
- `docx2md`: wrapper for docx -> markdown.
- `md2docx`: wrapper for markdown -> docx.

## Requirements

- `pandoc` available on `PATH`
- Python 3.10+

## Usage

Auto-detect mode from input extension:

```bash
docx-comments input.docx -o output.md
docx-comments input.md -o output.docx
```

Explicit wrappers:

```bash
docx2md input.docx -o output.md
md2docx input.md -o output.docx
```

Pass additional pandoc args through unchanged:

```bash
docx-comments input.docx -o output.md -- --reference-doc=template.docx
```

## Current behavior choices

- Thread replies are flattened into the root comment text in Word output.
- Child comments are removed from Word package internals after flattening to avoid duplicate standalone comments.
- Title comments at document start are supported via comment-ID-based parent mapping (not positional heuristics).

## Tests

Install optional test dependencies:

```bash
python3 -m pip install -r requirements-dev.txt
```

Run all tests:

```bash
python3 -m unittest -q
# or
make test
```

`make test` also runs the example roundtrip first and refreshes:

- `artifacts/out_test.md`
- `artifacts/out_test.docx`

Run roundtrip integration tests only:

```bash
python3 -m unittest -q tests.test_roundtrip_example
python3 -m unittest -q tests.test_roundtrip_edges
# or
make test-roundtrip
```

Run a pure example roundtrip and keep outputs for manual inspection:

```bash
make roundtrip-example
```

This writes:

- `artifacts/out_test.md`
- `artifacts/out_test.docx`

Clean those artifacts:

```bash
make clean-roundtrip-example
```

When a roundtrip assertion fails, tests write a `failure_bundle/` folder in the test temp directory with snapshots and diffs:

- `original_snapshot.json`
- `markdown_snapshot.json`
- `roundtrip_snapshot.json`
- `expected_flatten.json`
- `command_logs.json`
- `mismatch_report.txt`
