# Publishing Checklist

This is a practical checklist to run tests and publish `docx-comments-roundtrip`.

## 1) Pre-release checks

Run from repository root:

```bash
make test
python -m unittest -q tests.test_cli_entrypoints
```

Optional manual smoke artifact:

```bash
make roundtrip-example
```

Inspect:

- `artifacts/out_test.md`
- `artifacts/out_test.docx`

## 2) Clean build environment

Create and activate a clean virtualenv:

```bash
python -m venv .venv-release
source .venv-release/bin/activate
```

Install build tooling:

```bash
python -m pip install --upgrade pip build twine
```

## 3) Build package artifacts

```bash
rm -rf dist build *.egg-info
python -m build
python -m twine check dist/*
```

Expected outputs:

- `dist/docx_comments_roundtrip-<VERSION>-py3-none-any.whl`
- `dist/docx_comments_roundtrip-<VERSION>.tar.gz`

## 4) Local install smoke test

Install from wheel:

```bash
python -m pip install --force-reinstall dist/*.whl
```

Verify CLI entry points:

```bash
dmt --help
docx-comments --help
docx2md --help
md2docx --help
d2m --help
m2d --help
```

## 5) Publish to TestPyPI first (recommended)

Create a TestPyPI API token and configure `~/.pypirc`, or use env vars.

Upload:

```bash
python -m twine upload --repository testpypi dist/*
```

Install from TestPyPI in a fresh env:

```bash
python -m pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple docx-comments-roundtrip
dmt --help
```

## 6) Publish to PyPI

```bash
python -m twine upload dist/*
```

## 7) Tag and release

```bash
git tag v<VERSION>
git push origin v<VERSION>
```

Create a GitHub release and include:

- install command: `pipx install docx-comments-roundtrip`
- upgrade command: `pipx upgrade docx-comments-roundtrip`
- key changes and any migration notes.

## 8) Post-release verification

In a clean env:

```bash
python -m pip install --upgrade docx-comments-roundtrip
dmt --help
```

Run one real conversion sanity check:

```bash
docx2md your.docx -o /tmp/your.md
md2docx /tmp/your.md -o /tmp/your-roundtrip.docx
```
