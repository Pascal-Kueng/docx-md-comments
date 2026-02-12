# docx-comments-roundtrip

Lossless `.docx <-> .md` conversion focused on **Word comment fidelity**:

- comment anchors
- threaded replies
- active/resolved state

If your workflow is "edit in markdown/LLM, then return to Word without breaking comments", this tool is for that.

## Install

### Prerequisites

- Python 3.10+
- Pandoc available on `PATH`

Install Pandoc:

- macOS: `brew install pandoc`
- Ubuntu/Debian: `sudo apt-get install pandoc`
- Windows (PowerShell): `choco install pandoc -y`

### Recommended (isolated): `pipx`

```bash
pipx install docx-comments-roundtrip
```

Upgrade later:

```bash
pipx upgrade docx-comments-roundtrip
```

### Alternative: `pip`

```bash
python -m pip install docx-comments-roundtrip
```

## Quick usage

### Auto mode

All of these are equivalent:

```bash
dmt draft.docx
docx-comments draft.docx
docx2md draft.docx
d2m draft.docx
```

All of these are equivalent:

```bash
dmt draft.md
docx-comments draft.md
md2docx draft.md
m2d draft.md
```

### Explicit mode

DOCX -> Markdown:

```bash
docx2md draft.docx -o draft.md
d2m draft.docx -o draft.md
dmt docx2md draft.docx -o draft.md
```

Markdown -> DOCX:

```bash
md2docx draft.md -o draft.docx
m2d draft.md -o draft.docx
dmt md2docx draft.md -o draft.docx
```

Use a reference Word document for styling:

```bash
md2docx draft.md --ref original.docx -o final.docx
m2d draft.md -r original.docx -o final.docx
```

`--ref` maps to Pandoc `--reference-doc`.

### Pass-through Pandoc arguments

Unknown flags are passed through to Pandoc:

```bash
docx2md draft.docx -o draft.md --extract-media=media
dmt md2docx draft.md --reference-doc=template.docx
```

## Help

```bash
dmt --help
docx2md --help
md2docx --help
```

## Testing

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

## Report issues

Please open bugs/feature requests at:

https://github.com/pascalkueng/docx-comments-roundtrip/issues

When reporting a conversion bug, include:

- input sample (or minimal repro)
- command used
- expected vs actual behavior (Word view)
- failing `failure_bundle` path if tests failed

## Technical notes (brief)

- Markdown marker style uses `///C<ID>.START///` / `///C<ID>.END///` (with optional `==...==` highlight wrapper).
- Reply relationships are reconstructed as native Word threads (`commentsExtended.xml` `paraIdParent` + story markers).
- The validator fails fast on malformed marker edits with line-specific diagnostics.

For deeper maintainer details, see `AGENTS.md`.
