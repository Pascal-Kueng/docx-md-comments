from __future__ import annotations

import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

from tests.helpers.diagnostics import text_diff, write_failure_bundle
from tests.helpers.docx_inspector import inspect_docx, normalize_anchor_text, normalize_comment_text
from tests.helpers.markdown_inspector import inspect_markdown_comments

REPO_ROOT = Path(__file__).resolve().parents[1]
CONVERTER_PATH = REPO_ROOT / "docx-comments"


def run_converter(converter_path: Path, input_path: Path, output_path: Path, cwd: Path) -> dict:
    cmd = [str(converter_path), str(input_path), "-o", str(output_path)]
    proc = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True)
    return {
        "cmd": cmd,
        "returncode": proc.returncode,
        "stdout": proc.stdout,
        "stderr": proc.stderr,
    }


EDGE_CASES = [
    {
        "name": "heading_start_single_comment",
        "expected_root_ids": ["1", "2"],
        "expected_state_by_root": {"1": "active", "2": "active"},
        "markdown": (
            '# [Title]{.comment-start id="1" author="A" date="2026-01-01T00:00:00Z"}'
            ' Heading[]{.comment-end id="1"}\n\n'
            'Plain paragraph with [another]{.comment-start id="2" author="B" date="2026-01-01T00:01:00Z"}'
            ' span[]{.comment-end id="2"}.\n'
        ),
    },
    {
        "name": "nested_replies",
        "expected_root_ids": ["10"],
        "expected_state_by_root": {"10": "active"},
        "markdown": (
            "Paragraph with [root note]{.comment-start id=\"10\" author=\"Root\" date=\"2026-01-01T00:00:00Z\"}"
            "[first child]{.comment-start id=\"11\" author=\"R1\" date=\"2026-01-01T00:10:00Z\" parent=\"10\"}"
            "[grand child]{.comment-start id=\"13\" author=\"R2\" date=\"2026-01-01T00:20:00Z\" parent=\"11\"}"
            " text[]{.comment-end id=\"13\"}[]{.comment-end id=\"11\"}"
            " and [second child]{.comment-start id=\"12\" author=\"R3\" date=\"2026-01-01T00:30:00Z\" parent=\"10\"}"
            " text[]{.comment-end id=\"12\"}"
            " end[]{.comment-end id=\"10\"}.\n"
        ),
    },
    {
        "name": "interleaved_root_comments",
        "expected_root_ids": ["20", "21"],
        "expected_state_by_root": {"20": "active", "21": "active"},
        "markdown": (
            "Start [A]{.comment-start id=\"20\" author=\"A\" date=\"2026-01-02T00:00:00Z\"}"
            " mid [B]{.comment-start id=\"21\" author=\"B\" date=\"2026-01-02T00:01:00Z\"}"
            " inner[]{.comment-end id=\"20\"} tail[]{.comment-end id=\"21\"}.\n"
        ),
    },
    {
        "name": "nested_end_wrapper_with_multiple_inner_markers",
        "expected_root_ids": ["0", "1", "2"],
        "expected_state_by_root": {"0": "active", "1": "active", "2": "active"},
        "markdown": (
            '# [A]{.comment-start id="0" author="A" date="2026-01-03T00:00:00Z"}'
            '[reply]{.comment-start id="1" author="B" date="2026-01-03T00:01:00Z"}'
            '[B]{.comment-start id="2" author="C" date="2026-01-03T00:02:00Z"}'
            'x[[]{.comment-end id="1"}[]{.comment-end id="0"}]{.comment-end id="2"}\n'
        ),
    },
    {
        "name": "resolved_root_comment",
        "expected_root_ids": ["70"],
        "expected_state_by_root": {"70": "resolved"},
        "markdown": (
            "Root [resolved]{.comment-start id=\"70\" author=\"A\" date=\"2026-01-04T00:00:00Z\" state=\"resolved\"}"
            " mark[]{.comment-end id=\"70\"}.\n"
        ),
    },
    {
        "name": "mixed_root_states",
        "expected_root_ids": ["71", "72"],
        "expected_state_by_root": {"71": "resolved", "72": "active"},
        "markdown": (
            "A [resolved]{.comment-start id=\"71\" author=\"A\" date=\"2026-01-05T00:00:00Z\" state=\"resolved\"}"
            " span[]{.comment-end id=\"71\"} and "
            "[active]{.comment-start id=\"72\" author=\"B\" date=\"2026-01-05T00:01:00Z\" state=\"active\"}"
            " span[]{.comment-end id=\"72\"}.\n"
        ),
    },
    {
        "name": "invalid_state_defaults_active",
        "expected_root_ids": ["73"],
        "expected_state_by_root": {"73": "active"},
        "markdown": (
            "Root [invalid]{.comment-start id=\"73\" author=\"A\" date=\"2026-01-06T00:00:00Z\" state=\"banana\"}"
            " span[]{.comment-end id=\"73\"}.\n"
        ),
    },
]


class TestEdgeRoundtrips(unittest.TestCase):
    maxDiff = None

    @classmethod
    def setUpClass(cls) -> None:
        if shutil.which("pandoc") is None:
            raise unittest.SkipTest("pandoc not found on PATH")
        if not CONVERTER_PATH.exists():
            raise unittest.SkipTest(f"converter script not found: {CONVERTER_PATH}")

    def test_markdown_seed_roundtrip_stability(self) -> None:
        for case in EDGE_CASES:
            with self.subTest(case=case["name"]):
                self._run_case(case)

    def _run_case(self, case: dict) -> None:
        case_dir = Path(tempfile.mkdtemp(prefix=f"edge-{case['name']}-", dir="/tmp"))
        seed_md = case_dir / "seed.md"
        seed_docx = case_dir / "seed.docx"
        middle_md = case_dir / "middle.md"
        roundtrip_docx = case_dir / "roundtrip.docx"
        seed_md.write_text(case["markdown"], encoding="utf-8")

        command_logs = []
        command_logs.append(run_converter(CONVERTER_PATH, seed_md, seed_docx, REPO_ROOT))
        command_logs.append(run_converter(CONVERTER_PATH, seed_docx, middle_md, REPO_ROOT))
        command_logs.append(run_converter(CONVERTER_PATH, middle_md, roundtrip_docx, REPO_ROOT))

        for log in command_logs:
            if log["returncode"] != 0:
                self.fail(
                    f"Command failed for case {case['name']}\n"
                    f"cmd={' '.join(log['cmd'])}\n"
                    f"returncode={log['returncode']}\n"
                    f"stdout={log['stdout']}\n"
                    f"stderr={log['stderr']}"
                )

        seed_snapshot = inspect_docx(seed_docx)
        middle_snapshot = inspect_markdown_comments(middle_md)
        roundtrip_snapshot = inspect_docx(roundtrip_docx)

        errors: list[str] = []
        text_mismatch_diffs: dict[str, str] = {}

        expected_root_ids = case["expected_root_ids"]
        expected_state_by_root = case.get("expected_state_by_root") or {
            cid: "active" for cid in expected_root_ids
        }
        if seed_snapshot.comment_ids_order != expected_root_ids:
            errors.append(
                f"Unexpected seed root IDs for case {case['name']}. "
                f"expected={expected_root_ids} actual={seed_snapshot.comment_ids_order}"
            )

        for comment_id in expected_root_ids:
            expected_state = expected_state_by_root.get(comment_id, "active") == "resolved"
            actual_state = bool(seed_snapshot.resolved_by_id.get(comment_id, False))
            if expected_state != actual_state:
                errors.append(
                    f"Unexpected seed resolved-state for id={comment_id}. "
                    f"expected={expected_state} actual={actual_state}"
                )

        if roundtrip_snapshot.comment_ids_order != seed_snapshot.comment_ids_order:
            errors.append(
                "Roundtrip root IDs differ from seed docx roots. "
                f"seed={seed_snapshot.comment_ids_order} roundtrip={roundtrip_snapshot.comment_ids_order}"
            )

        if roundtrip_snapshot.parent_map:
            errors.append(
                f"Roundtrip parent map must be empty after flattening but got: {roundtrip_snapshot.parent_map}"
            )

        seed_anchor_set = set(seed_snapshot.anchor_ids_order)
        middle_start_set = set(middle_snapshot.start_ids_order)
        missing_markdown_starts = sorted(seed_anchor_set - middle_start_set, key=lambda value: (len(value), value))
        unexpected_markdown_starts = sorted(middle_start_set - seed_anchor_set, key=lambda value: (len(value), value))
        if missing_markdown_starts:
            errors.append(f"Markdown missing comment-start IDs from seed anchors: {missing_markdown_starts}")
        if unexpected_markdown_starts:
            errors.append(f"Markdown has unexpected comment-start IDs not in seed anchors: {unexpected_markdown_starts}")
        if len(middle_snapshot.start_ids_order) != len(middle_start_set):
            errors.append("Markdown comment-start IDs are not unique.")
        invalid_state_ids = sorted(
            [cid for cid in middle_snapshot.start_ids_order if middle_snapshot.state_by_id.get(cid) not in {"active", "resolved"}],
            key=lambda value: (len(value), value),
        )
        if invalid_state_ids:
            errors.append(f"Markdown has invalid state values for comment-start IDs: {invalid_state_ids}")
        for comment_id in middle_snapshot.start_ids_order:
            expected_state = "resolved" if seed_snapshot.resolved_by_id.get(comment_id, False) else "active"
            actual_state = middle_snapshot.state_by_id.get(comment_id, "active")
            if expected_state != actual_state:
                errors.append(
                    f"Markdown state mismatch for id={comment_id}. "
                    f"expected={expected_state} actual={actual_state}"
                )

        roundtrip_id_set = set(roundtrip_snapshot.comment_ids_order)
        for label, observed_ids in [
            ("anchor", set(roundtrip_snapshot.anchor_ids_order)),
            ("range-start", set(roundtrip_snapshot.range_start_ids)),
            ("range-end", set(roundtrip_snapshot.range_end_ids)),
            ("commentReference", set(roundtrip_snapshot.reference_ids)),
        ]:
            missing = sorted(roundtrip_id_set - observed_ids, key=lambda value: (len(value), value))
            unexpected = sorted(observed_ids - roundtrip_id_set, key=lambda value: (len(value), value))
            if missing:
                errors.append(f"Roundtrip missing {label} IDs for comments: {missing}")
            if unexpected:
                errors.append(f"Roundtrip has unexpected {label} IDs not in comments.xml: {unexpected}")

        if middle_snapshot.parent_by_id:
            errors.append(
                "Seed docx -> markdown should not reintroduce threaded parent attributes, "
                f"but found: {middle_snapshot.parent_by_id}"
            )

        if middle_snapshot.root_ids_order != seed_snapshot.comment_ids_order:
            errors.append(
                "Intermediate markdown root order mismatch. "
                f"expected={seed_snapshot.comment_ids_order} actual={middle_snapshot.root_ids_order}"
            )

        for comment_id in seed_snapshot.comment_ids_order:
            seed_node = seed_snapshot.comments_by_id.get(comment_id)
            roundtrip_node = roundtrip_snapshot.comments_by_id.get(comment_id)
            if seed_node is None or roundtrip_node is None:
                errors.append(f"Missing comment node for id={comment_id} in seed/roundtrip comparison")
                continue
            expected_norm = normalize_comment_text(seed_node.text)
            actual_norm = normalize_comment_text(roundtrip_node.text)
            if expected_norm != actual_norm:
                errors.append(f"Comment text drift for id={comment_id}")
                text_mismatch_diffs[comment_id] = text_diff(
                    seed_node.text, roundtrip_node.text, f"{case['name']}:{comment_id}"
                )

        for comment_id in seed_snapshot.comment_ids_order:
            expected_anchor_text = normalize_anchor_text(seed_snapshot.anchor_text_by_id.get(comment_id, ""))
            actual_anchor_text = normalize_anchor_text(roundtrip_snapshot.anchor_text_by_id.get(comment_id, ""))
            if expected_anchor_text != actual_anchor_text:
                errors.append(f"Anchor span text drift for id={comment_id}")
                text_mismatch_diffs[f"anchor_{comment_id}"] = text_diff(
                    seed_snapshot.anchor_text_by_id.get(comment_id, ""),
                    roundtrip_snapshot.anchor_text_by_id.get(comment_id, ""),
                    f"{case['name']}:anchor:{comment_id}",
                )

        if not roundtrip_snapshot.has_comments_extended:
            errors.append("Roundtrip missing word/commentsExtended.xml for resolved-state preservation")
        if roundtrip_snapshot.has_comments_ids:
            errors.append("Roundtrip contains word/commentsIds.xml")

        for comment_id in seed_snapshot.comment_ids_order:
            expected_state = bool(seed_snapshot.resolved_by_id.get(comment_id, False))
            actual_state = bool(roundtrip_snapshot.resolved_by_id.get(comment_id, False))
            if expected_state != actual_state:
                errors.append(
                    f"Roundtrip resolved-state drift for id={comment_id}. "
                    f"expected={expected_state} actual={actual_state}"
                )

        if errors:
            failure_bundle = write_failure_bundle(
                case_dir / "failure_bundle",
                original_snapshot=seed_snapshot,
                markdown_snapshot=middle_snapshot,
                roundtrip_snapshot=roundtrip_snapshot,
                expected_flatten={"case": case["name"], "expected_root_ids": expected_root_ids},
                command_logs=command_logs,
                errors=errors,
                text_mismatch_diffs=text_mismatch_diffs,
            )
            self.fail(
                f"Edge case failed: {case['name']}. Diagnostics: {failure_bundle}\n- " + "\n- ".join(errors)
            )
