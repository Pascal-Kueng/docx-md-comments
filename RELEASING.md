# Releasing Guide

Single source of truth for future package updates and releases.

## TL;DR (normal future release)

After version bump + tests, this is the production path:

```bash
git tag vX.Y.Z
git push origin vX.Y.Z
gh release create vX.Y.Z --title "vX.Y.Z" --generate-notes
```

Publishing the GitHub release triggers `.github/workflows/publish.yml`, which publishes to PyPI automatically (Trusted Publishing / OIDC).

This repository uses Trusted Publishing via GitHub Actions:

- Workflow: `.github/workflows/publish.yml`
- Trigger 1: publish a GitHub Release -> auto publish to PyPI
- Trigger 2: manual run with `target=testpypi|pypi`
- Duplicate uploads are ignored (`skip-existing: true`)

## 0) One-time setup (already done once, keep for reference)

GitHub repository:

- Create environments: `pypi`, `testpypi`

PyPI trusted publisher:

- Owner: `Pascal-Kueng`
- Repository: `docx-md-comments`
- Workflow: `publish.yml`
- Environment: `pypi`

TestPyPI trusted publisher:

- Owner: `Pascal-Kueng`
- Repository: `docx-md-comments`
- Workflow: `publish.yml`
- Environment: `testpypi`

No PyPI API token secrets are required for this flow.

## 1) Choose version

Follow semver:

- Patch: `X.Y.Z+1` (bug fixes)
- Minor: `X.Y+1.0` (new backward-compatible features)
- Major: `X+1.0.0` (breaking changes)

## 2) Bump version in code

Update both files to the exact same value:

- `pyproject.toml` -> `[project].version`
- `src/dmc/version.py` -> `__version__`

Example for `1.0.1`:

```bash
rg -n '^version =|__version__' pyproject.toml src/dmc/version.py
```

## 3) Run release checks locally

Required:

```bash
make test
python -m unittest -q tests.test_cli_entrypoints
```

Optional extra artifact check:

```bash
make roundtrip-example
```

## 4) Commit and push release candidate

```bash
git add pyproject.toml src/dmc/version.py
git add .github/workflows/publish.yml RELEASING.md
git commit -m "release: vX.Y.Z"
git push
```

If workflow/docs were not changed in this release, omit those files from `git add`.

## 5) Optional TestPyPI dry run (recommended)

Use manual workflow dispatch:

```bash
gh workflow run Publish -f target=testpypi
gh run list --workflow Publish --limit 5
```

Wait for run success before production release.

## 6) Tag and create GitHub Release (production publish)

Create and push tag:

```bash
git tag vX.Y.Z
git push origin vX.Y.Z
```

Create/publish GitHub Release (this triggers PyPI publish automatically):

```bash
gh release create vX.Y.Z --title "vX.Y.Z" --generate-notes
```

Equivalent UI path:

1. GitHub -> Releases -> Draft a new release
2. Select tag `vX.Y.Z`
3. Click `Publish release`

## 7) Verify published package

Check workflow result:

```bash
gh run list --workflow Publish --limit 5
```

Install from PyPI in a clean environment:

```bash
pipx install docx-md-comments
dmc --help
```

Or if already installed:

```bash
pipx upgrade docx-md-comments
dmc --help
```

## 8) Release notes checklist

Include in GitHub Release notes:

- main fixes/features
- install command: `pipx install docx-md-comments`
- upgrade command: `pipx upgrade docx-md-comments`
- migration notes for breaking changes

## 9) Common failure modes

`Publish` failed with auth/trusted-publisher errors:

- Re-check PyPI/TestPyPI trusted publisher owner/repo/workflow/environment values.

`Publish` failed due duplicate version:

- With current config, duplicates should be skipped.
- If nothing new uploaded, bump version and release new tag.

Local tag already exists but remote does not:

```bash
git push origin <tag>
```

Need to recreate local tag:

```bash
git tag -d <tag>
git tag <tag>
git push origin <tag>
```
