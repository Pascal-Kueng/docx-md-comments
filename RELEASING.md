# Releasing Guide

Use this for every new release after `1.0.0`.

Detailed publish commands live in `PUBLISHING.md`.

## 1) Choose version (semver)

- Patch (`X.Y.Z+1`): bug fixes only.
- Minor (`X.Y+1.0`): backward-compatible features.
- Major (`X+1.0.0`): breaking changes.

## 2) Update version in code

Edit both files:

- `pyproject.toml` -> `[project].version`
- `src/dmc/version.py` -> `__version__`

## 3) Run checks

```bash
make test
python -m unittest -q tests.test_cli_entrypoints
```

If needed:

```bash
make roundtrip-example
```

## 4) Commit + tag

```bash
git add -A
git commit -m "release: vX.Y.Z"
git tag vX.Y.Z
git push
git push origin vX.Y.Z
```

## 5) Build + publish

Follow `PUBLISHING.md` sections:

1. Build/check artifacts
2. TestPyPI publish + install smoke test
3. PyPI publish

## 6) GitHub release notes

Include:

- highlights / fixes
- install command: `pipx install docx-md-comments`
- upgrade command: `pipx upgrade docx-md-comments`
- any migration notes (if breaking changes)

## 7) Post-release verification

In a clean environment:

```bash
python -m pip install --upgrade docx-md-comments
dmc --help
```

Run one real conversion smoke test with a sample `.docx`.
