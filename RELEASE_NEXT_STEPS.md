# Release Next Steps

The package is close to distribution-ready. Use this checklist to publish safely.

## 1) Finalize repository state

1. Commit all packaging and CLI changes.
2. Confirm the package name in `pyproject.toml` is the one you want on PyPI.
3. Ensure licensing is consistent:
   - `pyproject.toml` declares Apache-2.0
   - repository contains a matching `LICENSE` file

## 2) Build and validate artifacts

Install build tools:

```bash
python -m pip install build twine
```

Build sdist + wheel:

```bash
python -m build
```

Validate metadata:

```bash
python -m twine check dist/*
```

## 3) Smoke-test installation in a clean environment

Install built wheel:

```bash
python -m pip install dist/*.whl
```

Verify CLI entry points:

```bash
dmt --help
docx2md --help
md2docx --help
d2m --help
m2d --help
```

## 4) Publish strategy

1. Publish to TestPyPI first (recommended).
2. Run one install test from TestPyPI.
3. Publish to PyPI.

Use either:

- Trusted Publishing via GitHub Actions (recommended)
- `twine upload dist/*`

## 5) First release hygiene

1. Tag release (for example `v0.1.0`).
2. Add release notes with install/upgrade commands:
   - `pipx install <package>`
   - `pipx upgrade <package>`
3. Verify install and CLI behavior on macOS, Linux, and Windows.
