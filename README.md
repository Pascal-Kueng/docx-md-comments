# docx-md-comments

Convert Word files with comments to Markdown, edit them, and convert back to Word while keeping comment threads intact.

This tool preserves:

- comment anchors
- threaded replies
- active/resolved state

If your workflow is:

1. start with a `.docx` that has comments
2. edit in Markdown (manually or with an LLM)
3. convert back to `.docx`

then this package is built for that.

## Before You Start

You need:

- Python 3.10+
- Pandoc installed and available on `PATH`

Install Pandoc:

- macOS: `brew install pandoc`
- Ubuntu/Debian: `sudo apt-get install pandoc`
- Windows (PowerShell): `choco install pandoc -y`

## Install (Recommended)

Use `pipx` (cleanest option for command-line apps):

```bash
pipx install docx-md-comments
```

That gives you these commands:

- `dmc`
- `docx-comments`
- `docx2md` / `d2m`
- `md2docx` / `m2d`

## Update To Latest Version

If installed with `pipx`:

```bash
pipx upgrade docx-md-comments
```

If installed with `pip`:

```bash
python -m pip install --upgrade docx-md-comments
```

## Install (Alternative)

If you prefer `pip`:

```bash
python -m pip install docx-md-comments
```

## Quick Start (Most Users)

Convert Word to Markdown:

```bash
dmc draft.docx
```

This creates `draft.md`.

Convert Markdown back to Word:

```bash
dmc draft.md
```

This creates `draft.docx`.

## More Commands (Optional)

These commands are equivalent; use whichever you prefer.

DOCX -> Markdown:

```bash
docx-comments draft.docx
docx2md draft.docx
d2m draft.docx
```

Markdown -> DOCX:

```bash
md2docx draft.md
m2d draft.md
```

### Explicit input/output paths

DOCX -> Markdown:

```bash
docx2md draft.docx -o draft.md
d2m draft.docx -o draft.md
dmc docx2md draft.docx -o draft.md
```

Markdown -> DOCX:

```bash
md2docx draft.md -o draft.docx
m2d draft.md -o draft.docx
dmc md2docx draft.md -o draft.docx
```

Use a reference Word document for styling:

```bash
md2docx draft.md --ref original.docx -o final.docx
m2d draft.md -r original.docx -o final.docx
```

`--ref` maps to Pandoc `--reference-doc`.

### Pass-through Pandoc arguments (advanced)

Unknown flags are passed through to Pandoc:

```bash
docx2md draft.docx -o draft.md --extract-media=media
dmc md2docx draft.md --reference-doc=template.docx
```

## Limitations

- **Tracked Changes:** Word revisions are not preserved through roundtrip. Resolve them in Word first.
- **Formatting:** Main focus is preserving comment threads. Very complex Word layouts may not roundtrip perfectly.

## Help

```bash
dmc --help
docx2md --help
md2docx --help
```

## For Contributors: Testing

Run full suite:

```bash
make test
```

Roundtrip-focused tests only:

```bash
make test-roundtrip
```

`make test` also writes:

- `artifacts/out_test.md`
- `artifacts/out_test.docx`

## Report Issues

Please open bugs/feature requests at:

https://github.com/Pascal-Kueng/docx-md-comments/issues

When reporting a conversion bug, include:

- input sample (or minimal repro)
- command used
- expected vs actual behavior (Word view)
- failing `failure_bundle` path if tests failed

## Technical Notes (Brief)

- Markdown marker style uses `///C<ID>.START///` / `///C<ID>.END///` (with optional `==...==` highlight wrapper).
- Reply relationships are reconstructed as native Word threads (`commentsExtended.xml` `paraIdParent` + story markers).
- The validator fails fast on malformed marker edits with line-specific diagnostics.

For deeper maintainer details, see `AGENTS.md`.
