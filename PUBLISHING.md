# Publishing Checklist

This is a practical checklist to run tests and publish `docx-md-comments`.

## 0) Finalize repository state

Before building artifacts:

- Commit packaging and CLI changes.
- Confirm the package name in `pyproject.toml` is what you want on PyPI.
- Ensure licensing is consistent:
  - `pyproject.toml` declares `Apache-2.0`
  - repository contains a matching `LICENSE` file

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

Create and activate a clean virtualenv (this keeps your system clean):

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

- `dist/docx_md_comments-<VERSION>-py3-none-any.whl`
- `dist/docx_md_comments-<VERSION>.tar.gz`

## 4) Local install smoke test

Install from wheel within the build env to verify structure:

```bash
python -m pip install --force-reinstall dist/*.whl
```

Verify CLI entry points:

```bash
dmc --help
docx-comments --help
docx2md --help
md2docx --help
d2m --help
m2d --help
```

## 5) Publish to TestPyPI first (recommended)

Upload using the pre-configured `~/.pypirc` tokens:

```bash
python -m twine upload --repository testpypi dist/*
```

**Verification Step:**
Open a NEW terminal window (don't use the build env) to simulate a fresh user:

```bash
cd /tmp
python -m venv test-pypi-env
source test-pypi-env/bin/activate
python -m pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple docx-md-comments
dmc --help
```

**Cleanup**

```bash
deactivate
rm -rf test-pypi-env
```

## 6) Publish to PyPI

Back in your build terminal (`.venv-release`), upload to the live registry:

```bash
python -m twine upload dist/*
```

Alternative:

- Use Trusted Publishing via GitHub Actions (recommended) instead of manual `twine` upload.

## 7) Tag and release

```bash
git tag v<VERSION>
git push origin v<VERSION>
```

Create a GitHub release and include:

- install command: `pipx install docx-md-comments`
- upgrade command: `pipx upgrade docx-md-comments`
- key changes and any migration notes.

## 8) Post-release cleanup

Remove the build environment and artifacts:

```bash
deactivate
rm -rf .venv-release dist build *.egg-info
```

## 9) Final user verification

In your normal terminal (or using pipx):

```bash
pipx install docx-md-comments
dmc --help
```

Recommended additional verification:

- Run install/CLI checks on macOS, Linux, and Windows.
