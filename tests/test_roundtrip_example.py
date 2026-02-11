from __future__ import annotations

import shutil
import subprocess
import tempfile
import unittest
from hashlib import sha256
from pathlib import Path

from tests.helpers.diagnostics import text_diff, write_failure_bundle
from tests.helpers.docx_inspector import (
    build_flatten_expectation,
    inspect_docx,
    normalize_comment_text,
)
from tests.helpers.markdown_inspector import inspect_markdown_comments

REPO_ROOT = Path(__file__).resolve().parents[1]
CONVERTER_PATH = REPO_ROOT / "docx-comments"
EXAMPLE_DOCX = REPO_ROOT / "Preregistration_Original.docx"


def run_converter(converter_path: Path, input_path: Path, output_path: Path, cwd: Path) -> dict:
    cmd = [str(converter_path), str(input_path), "-o", str(output_path)]
    proc = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True)
    return {
        "cmd": cmd,
        "returncode": proc.returncode,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
    }


def format_set(values) -> str:
    return "[" + ", ".join(sorted(values, key=lambda value: (len(value), value))) + "]"


class TestPreregistrationRoundtrip(unittest.TestCase):
    maxDiff = None

    @classmethod
    def setUpClass(cls) -> None:
        if shutil.which("pandoc") is None:
            raise unittest.SkipTest("pandoc not found on PATH")
        if not CONVERTER_PATH.exists():
            raise unittest.SkipTest(f"converter script not found: {CONVERTER_PATH}")
        if not EXAMPLE_DOCX.exists():
            raise unittest.SkipTest(f"example fixture not found: {EXAMPLE_DOCX}")

    def test_comment_integrity_and_thread_flattening(self) -> None:
        case_dir = Path(tempfile.mkdtemp(prefix="roundtrip-example-", dir="/tmp"))

        source_docx = case_dir / "input.docx"
        intermediate_md = case_dir / "roundtrip.md"
        roundtrip_docx = case_dir / "roundtrip.docx"
        original_digest_before = sha256(EXAMPLE_DOCX.read_bytes()).hexdigest()
        shutil.copyfile(EXAMPLE_DOCX, source_docx)

        command_logs = []
        command_logs.append(run_converter(CONVERTER_PATH, source_docx, intermediate_md, REPO_ROOT))
        command_logs.append(run_converter(CONVERTER_PATH, intermediate_md, roundtrip_docx, REPO_ROOT))
        original_digest_after = sha256(EXAMPLE_DOCX.read_bytes()).hexdigest()
        if original_digest_before != original_digest_after:
            self.fail("Original fixture was modified during roundtrip test, which must never happen.")

        for log in command_logs:
            if log["returncode"] != 0:
                self.fail(
                    "Converter command failed.\n"
                    f"cmd={' '.join(log['cmd'])}\n"
                    f"returncode={log['returncode']}\n"
                    f"stdout={log['stdout']}\n"
                    f"stderr={log['stderr']}"
                )

        original = inspect_docx(source_docx)
        markdown = inspect_markdown_comments(intermediate_md)
        roundtrip = inspect_docx(roundtrip_docx)
        expected_from_original = build_flatten_expectation(original)

        errors: list[str] = []
        text_mismatch_diffs: dict[str, str] = {}

        original_root_set = set(expected_from_original.root_ids_order)
        roundtrip_root_set = set(roundtrip.comment_ids_order)
        if roundtrip_root_set != original_root_set:
            missing = original_root_set - roundtrip_root_set
            unexpected = roundtrip_root_set - original_root_set
            if missing:
                errors.append(f"Missing original root comment IDs in roundtrip: {format_set(missing)}")
            if unexpected:
                errors.append(f"Unexpected roundtrip root comment IDs: {format_set(unexpected)}")

        expected_parent_map = {
            child_id: parent_id
            for child_id, parent_id in original.parent_map.items()
            if child_id in markdown.start_ids_order and parent_id in markdown.start_ids_order
        }
        if markdown.parent_by_id != expected_parent_map:
            missing = sorted(set(expected_parent_map.items()) - set(markdown.parent_by_id.items()))
            unexpected = sorted(set(markdown.parent_by_id.items()) - set(expected_parent_map.items()))
            if missing:
                errors.append(f"Missing parent attributes in markdown spans: {missing}")
            if unexpected:
                errors.append(f"Unexpected parent attributes in markdown spans: {unexpected}")

        original_anchor_set = set(original.anchor_ids_order)
        markdown_start_set = set(markdown.start_ids_order)
        missing_markdown_starts = sorted(
            original_anchor_set - markdown_start_set, key=lambda value: (len(value), value)
        )
        unexpected_markdown_starts = sorted(
            markdown_start_set - original_anchor_set, key=lambda value: (len(value), value)
        )
        if missing_markdown_starts:
            errors.append(f"Markdown missing comment-start IDs from original anchors: {missing_markdown_starts}")
        if unexpected_markdown_starts:
            errors.append(f"Markdown has unexpected comment-start IDs not in original anchors: {unexpected_markdown_starts}")

        if len(markdown.start_ids_order) != len(markdown_start_set):
            errors.append("Markdown comment-start IDs are not unique.")

        original_child_set = set(expected_from_original.child_ids)
        for label, values in [
            ("roundtrip comment nodes", roundtrip.comment_ids_order),
            ("roundtrip anchor IDs", roundtrip.anchor_ids_order),
            ("roundtrip range-start IDs", roundtrip.range_start_ids),
            ("roundtrip reference IDs", roundtrip.reference_ids),
        ]:
            leaked = original_child_set.intersection(values)
            if leaked:
                errors.append(f"Thread child IDs leaked into {label}: {format_set(leaked)}")

        if roundtrip.parent_map:
            errors.append(f"Roundtrip parent map must be empty after flattening but got: {roundtrip.parent_map}")

        expected_roundtrip_order = [cid for cid in expected_from_original.root_ids_order if cid in roundtrip.comments_by_id]
        if roundtrip.comment_ids_order != expected_roundtrip_order:
            errors.append(
                "Roundtrip root comment order mismatch. "
                f"expected={expected_roundtrip_order} actual={roundtrip.comment_ids_order}"
            )

        if markdown.root_ids_order != expected_from_original.root_ids_order:
            errors.append(
                "Markdown root comment order mismatch vs original-derived roots. "
                f"expected={expected_from_original.root_ids_order} actual={markdown.root_ids_order}"
            )

        for root_id in expected_from_original.root_ids_order:
            expected_from_orig_text = expected_from_original.flattened_by_root.get(root_id, "")
            markdown_text = markdown.flattened_by_root.get(root_id, "")
            if normalize_comment_text(expected_from_orig_text) != normalize_comment_text(markdown_text):
                errors.append(f"Docx->markdown flattened text mismatch for root comment {root_id}")
                text_mismatch_diffs[f"md_{root_id}"] = text_diff(
                    expected_from_orig_text, markdown_text, f"markdown comment {root_id}"
                )

        for root_id in expected_roundtrip_order:
            expected_text = expected_from_original.flattened_by_root.get(root_id, "")
            actual_node = roundtrip.comments_by_id.get(root_id)
            if actual_node is None:
                errors.append(f"Roundtrip missing expected root comment: {root_id}")
                continue
            expected_norm = normalize_comment_text(expected_text)
            actual_norm = normalize_comment_text(actual_node.text)
            if expected_norm != actual_norm:
                errors.append(f"Flattened text mismatch for root comment {root_id}")
                text_mismatch_diffs[root_id] = text_diff(expected_text, actual_node.text, f"comment {root_id}")

        for root_id in expected_from_original.root_ids_order:
            expected_anchor_text = normalize_comment_text(original.anchor_text_by_id.get(root_id, ""))
            actual_anchor_text = normalize_comment_text(roundtrip.anchor_text_by_id.get(root_id, ""))
            if expected_anchor_text != actual_anchor_text:
                errors.append(f"Anchor span text mismatch for root comment {root_id}")
                text_mismatch_diffs[f"anchor_{root_id}"] = text_diff(
                    original.anchor_text_by_id.get(root_id, ""),
                    roundtrip.anchor_text_by_id.get(root_id, ""),
                    f"anchor {root_id}",
                )

        roundtrip_id_set = set(roundtrip.comment_ids_order)
        for label, observed_ids in [
            ("anchor", set(roundtrip.anchor_ids_order)),
            ("range-start", set(roundtrip.range_start_ids)),
            ("range-end", set(roundtrip.range_end_ids)),
            ("commentReference", set(roundtrip.reference_ids)),
        ]:
            missing = sorted(roundtrip_id_set - observed_ids, key=lambda value: (len(value), value))
            unexpected = sorted(observed_ids - roundtrip_id_set, key=lambda value: (len(value), value))
            if missing:
                errors.append(f"Roundtrip missing {label} IDs for comments: {missing}")
            if unexpected:
                errors.append(f"Roundtrip has unexpected {label} IDs not in comments.xml: {unexpected}")

        if roundtrip.has_comments_extended:
            errors.append("Roundtrip still contains word/commentsExtended.xml")
        if roundtrip.has_comments_ids:
            errors.append("Roundtrip still contains word/commentsIds.xml")

        if markdown.placeholder_shape_match_count > 0:
            errors.append(
                "Intermediate markdown still contains shape-placeholder image markers. "
                f"count={markdown.placeholder_shape_match_count}"
            )

        if errors:
            failure_bundle = write_failure_bundle(
                case_dir / "failure_bundle",
                original_snapshot=original,
                markdown_snapshot=markdown,
                roundtrip_snapshot=roundtrip,
                expected_flatten={
                    "from_original": expected_from_original.to_dict(),
                    "from_markdown": markdown.flattened_by_root,
                    "none_line_count_in_markdown": markdown.none_line_count,
                },
                command_logs=command_logs,
                errors=errors,
                text_mismatch_diffs=text_mismatch_diffs,
            )
            self.fail(
                "Roundtrip comment/thread assertions failed. "
                f"Diagnostics: {failure_bundle}\n- " + "\n- ".join(errors)
            )
