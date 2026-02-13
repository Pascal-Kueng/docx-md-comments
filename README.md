# docx-md-comments

Takes word documents with comment threads and converts them to markdown that is easily readable by LLMs and humans. The markdown documents can be backconverted into .docx files with nested comment threads fully restored. 

Lossless `.docx <-> .md` comments conversion:

- comment anchors
- threaded replies
- active/resolved state

If your workflow is "convert .docx with comments --> edit in markdown using LLMs --> return to Word, restoring remaining (or new) comments or LLM comments", this tool is for that.

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
pipx install docx-md-comments
```

Upgrade later:

```bash
pipx upgrade docx-md-comments
```

### Alternative: `pip`

```bash
python -m pip install docx-md-comments
```

## Quick usage

DOCX -> MD (all of these are equivalent):

```bash
dmc draft.docx
docx-comments draft.docx
docx2md draft.docx
d2m draft.docx
```

MD -> DOCX (all of these are equivalent):

```bash
dmc draft.md
docx-comments draft.md
md2docx draft.md
m2d draft.md
```

### Explicit mode

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

### Pass-through Pandoc arguments

Unknown flags are passed through to Pandoc:

```bash
docx2md draft.docx -o draft.md --extract-media=media
dmc md2docx draft.md --reference-doc=template.docx
```

## Help

```bash
dmc --help
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
