from __future__ import annotations

import difflib
import json
from dataclasses import asdict, is_dataclass
from pathlib import Path


def _to_jsonable(value):
    if hasattr(value, "to_dict") and callable(value.to_dict):
        return value.to_dict()
    if is_dataclass(value):
        return asdict(value)
    if isinstance(value, dict):
        return {k: _to_jsonable(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_to_jsonable(v) for v in value]
    return value


def _write_json(path: Path, value) -> None:
    path.write_text(json.dumps(_to_jsonable(value), indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def text_diff(expected: str, actual: str, label: str, max_lines: int = 60) -> str:
    expected_lines = (expected or "").splitlines()
    actual_lines = (actual or "").splitlines()
    lines = list(
        difflib.unified_diff(
            expected_lines,
            actual_lines,
            fromfile=f"{label}:expected",
            tofile=f"{label}:actual",
            lineterm="",
        )
    )
    if len(lines) > max_lines:
        lines = lines[:max_lines] + ["... diff truncated ..."]
    return "\n".join(lines)


def write_failure_bundle(
    bundle_dir: Path,
    *,
    original_snapshot,
    markdown_snapshot,
    roundtrip_snapshot,
    expected_flatten,
    command_logs: list[dict],
    errors: list[str],
    text_mismatch_diffs: dict[str, str] | None = None,
) -> Path:
    bundle_dir = Path(bundle_dir)
    bundle_dir.mkdir(parents=True, exist_ok=True)

    _write_json(bundle_dir / "original_snapshot.json", original_snapshot)
    _write_json(bundle_dir / "markdown_snapshot.json", markdown_snapshot)
    _write_json(bundle_dir / "roundtrip_snapshot.json", roundtrip_snapshot)
    _write_json(bundle_dir / "expected_flatten.json", expected_flatten)
    _write_json(bundle_dir / "command_logs.json", command_logs)

    if text_mismatch_diffs:
        text_diff_dir = bundle_dir / "text_diffs"
        text_diff_dir.mkdir(parents=True, exist_ok=True)
        for comment_id, diff in text_mismatch_diffs.items():
            (text_diff_dir / f"comment_{comment_id}.diff").write_text(diff + "\n", encoding="utf-8")

    report_lines = []
    report_lines.append("Roundtrip comment/thread verification failed.")
    report_lines.append("")
    report_lines.append("Errors:")
    for idx, error in enumerate(errors, start=1):
        report_lines.append(f"{idx}. {error}")

    if text_mismatch_diffs:
        report_lines.append("")
        report_lines.append("Text diffs:")
        for comment_id in sorted(text_mismatch_diffs):
            report_lines.append(f"- text_diffs/comment_{comment_id}.diff")

    report_lines.append("")
    report_lines.append("Artifacts:")
    report_lines.append("- original_snapshot.json")
    report_lines.append("- markdown_snapshot.json")
    report_lines.append("- roundtrip_snapshot.json")
    report_lines.append("- expected_flatten.json")
    report_lines.append("- command_logs.json")

    (bundle_dir / "mismatch_report.txt").write_text("\n".join(report_lines) + "\n", encoding="utf-8")
    return bundle_dir
