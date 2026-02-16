from __future__ import annotations

import os
import shutil
import subprocess
import sys
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
    if os.name == "nt":
        cmd = [sys.executable] + cmd
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
        "expected_comment_ids": ["1", "2"],
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
        "expected_comment_ids": ["10", "11", "13", "12"],
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
        "expected_comment_ids": ["20", "21"],
        "expected_state_by_root": {"20": "active", "21": "active"},
        "markdown": (
            "Start [A]{.comment-start id=\"20\" author=\"A\" date=\"2026-01-02T00:00:00Z\"}"
            " mid [B]{.comment-start id=\"21\" author=\"B\" date=\"2026-01-02T00:01:00Z\"}"
            " inner[]{.comment-end id=\"20\"} tail[]{.comment-end id=\"21\"}.\n"
        ),
    },
    {
        "name": "nested_end_wrapper_with_multiple_inner_markers",
        "expected_comment_ids": ["0", "1", "2"],
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
        "expected_comment_ids": ["70"],
        "expected_state_by_root": {"70": "resolved"},
        "markdown": (
            "Root [resolved]{.comment-start id=\"70\" author=\"A\" date=\"2026-01-04T00:00:00Z\" state=\"resolved\"}"
            " mark[]{.comment-end id=\"70\"}.\n"
        ),
    },
    {
        "name": "mixed_root_states",
        "expected_comment_ids": ["71", "72"],
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
        "expected_comment_ids": ["73"],
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
        case_dir = Path(tempfile.mkdtemp(prefix=f"edge-{case['name']}-"))
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
        for label, snapshot in [("seed", seed_snapshot), ("roundtrip", roundtrip_snapshot)]:
            compatibility_mode = (snapshot.settings_compatibility_mode or "").strip()
            compatibility_mode_int = None
            try:
                compatibility_mode_int = int(compatibility_mode)
            except ValueError:
                compatibility_mode_int = None
            if compatibility_mode_int is None or compatibility_mode_int < 15:
                errors.append(
                    f"{label} settings.xml missing modern compatibility mode for case {case['name']}. "
                    f"expected>=15 actual={compatibility_mode or '(missing)'}"
                )

        expected_comment_ids = case["expected_comment_ids"]
        expected_state_by_root = case.get("expected_state_by_root") or {
            cid: "active" for cid in expected_comment_ids
        }
        if seed_snapshot.comment_ids_order != expected_comment_ids:
            errors.append(
                f"Unexpected seed comment IDs for case {case['name']}. "
                f"expected={expected_comment_ids} actual={seed_snapshot.comment_ids_order}"
            )

        for comment_id in expected_comment_ids:
            expected_state = expected_state_by_root.get(comment_id, "active") == "resolved"
            actual_state = bool(seed_snapshot.resolved_by_id.get(comment_id, False))
            if expected_state != actual_state:
                errors.append(
                    f"Unexpected seed resolved-state for id={comment_id}. "
                    f"expected={expected_state} actual={actual_state}"
                )

        if roundtrip_snapshot.comment_ids_order != seed_snapshot.comment_ids_order:
            errors.append(
                "Roundtrip comment IDs differ from seed docx comments. "
                f"seed={seed_snapshot.comment_ids_order} roundtrip={roundtrip_snapshot.comment_ids_order}"
            )

        if roundtrip_snapshot.parent_map != seed_snapshot.parent_map:
            errors.append(
                "Roundtrip parent map drift. "
                f"expected={seed_snapshot.parent_map} actual={roundtrip_snapshot.parent_map}"
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

        expected_anchor_ids = set(seed_snapshot.comment_ids_order)
        for label, observed_ids in [
            ("anchor", set(roundtrip_snapshot.anchor_ids_order)),
            ("range-start", set(roundtrip_snapshot.range_start_ids)),
            ("range-end", set(roundtrip_snapshot.range_end_ids)),
            ("commentReference", set(roundtrip_snapshot.reference_ids)),
        ]:
            missing = sorted(expected_anchor_ids - observed_ids, key=lambda value: (len(value), value))
            unexpected = sorted(observed_ids - expected_anchor_ids, key=lambda value: (len(value), value))
            if missing:
                errors.append(f"Roundtrip missing {label} IDs for comments: {missing}")
            if unexpected:
                errors.append(f"Roundtrip has unexpected {label} IDs not in comments.xml: {unexpected}")

        expected_middle_parent_map = {
            child_id: parent_id
            for child_id, parent_id in seed_snapshot.parent_map.items()
            if child_id in middle_snapshot.start_ids_order and parent_id in middle_snapshot.start_ids_order
        }
        if middle_snapshot.parent_by_id != expected_middle_parent_map:
            errors.append(
                "Seed docx -> markdown parent attributes drift. "
                f"expected={expected_middle_parent_map} actual={middle_snapshot.parent_by_id}"
            )

        seed_root_ids = [cid for cid in seed_snapshot.comment_ids_order if cid not in seed_snapshot.parent_map]
        if middle_snapshot.root_ids_order != seed_root_ids:
            errors.append(
                "Intermediate markdown root order mismatch. "
                f"expected={seed_root_ids} actual={middle_snapshot.root_ids_order}"
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

        for comment_id in seed_root_ids:
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

        for comment_id in seed_snapshot.comment_ids_order:
            expected_state = bool(seed_snapshot.resolved_by_id.get(comment_id, False))
            actual_state = bool(roundtrip_snapshot.resolved_by_id.get(comment_id, False))
            if expected_state != actual_state:
                errors.append(
                    f"Roundtrip resolved-state drift for id={comment_id}. "
                    f"expected={expected_state} actual={actual_state}"
                )

            seed_node = seed_snapshot.comments_by_id.get(comment_id)
            roundtrip_node = roundtrip_snapshot.comments_by_id.get(comment_id)
            if seed_node is not None and roundtrip_node is not None:
                expected_para_id = (seed_node.para_id or "").strip()
                actual_para_id = (roundtrip_node.para_id or "").strip()
                if expected_para_id and actual_para_id and expected_para_id != actual_para_id:
                    errors.append(
                        f"Roundtrip paraId drift for id={comment_id}. "
                        f"expected={expected_para_id} actual={actual_para_id}"
                    )

                expected_durable_id = (
                    seed_snapshot.comments_ids_durable_by_para.get(expected_para_id, "").strip()
                    if expected_para_id
                    else ""
                )
                actual_durable_id = (
                    roundtrip_snapshot.comments_ids_durable_by_para.get(actual_para_id, "").strip()
                    if actual_para_id
                    else ""
                )
                if expected_durable_id and actual_durable_id and expected_durable_id != actual_durable_id:
                    errors.append(
                        f"Roundtrip durableId drift for id={comment_id}. "
                        f"expected={expected_durable_id} actual={actual_durable_id}"
                    )

        expected_resolved_comment_ids = sorted(
            [
                cid
                for cid in seed_snapshot.comment_ids_order
                if bool(seed_snapshot.resolved_by_id.get(cid, False))
            ],
            key=lambda value: (len(value), value),
        )
        if expected_resolved_comment_ids:
            expected_resolved_count = len(expected_resolved_comment_ids)
            actual_resolved_count = len(
                [cid for cid in seed_snapshot.comment_ids_order if bool(roundtrip_snapshot.resolved_by_id.get(cid, False))]
            )
            if actual_resolved_count != expected_resolved_count:
                errors.append(
                    "Resolved root-count mismatch after roundtrip: "
                    f"expected={expected_resolved_count} actual={actual_resolved_count}"
                )

            required_flags = [
                ("word/commentsExtended.xml", roundtrip_snapshot.has_comments_extended),
                ("word/commentsIds.xml", roundtrip_snapshot.has_comments_ids),
                ("word/commentsExtensible.xml", roundtrip_snapshot.has_comments_extensible),
                ("document.xml.rels commentsExtended relationship", roundtrip_snapshot.has_comments_extended_rel),
                ("document.xml.rels commentsIds relationship", roundtrip_snapshot.has_comments_ids_rel),
                ("document.xml.rels commentsExtensible relationship", roundtrip_snapshot.has_comments_extensible_rel),
                (
                    "[Content_Types].xml commentsExtended override",
                    roundtrip_snapshot.has_comments_extended_content_type,
                ),
                ("[Content_Types].xml commentsIds override", roundtrip_snapshot.has_comments_ids_content_type),
                (
                    "[Content_Types].xml commentsExtensible override",
                    roundtrip_snapshot.has_comments_extensible_content_type,
                ),
            ]
            if seed_snapshot.has_people or seed_snapshot.has_people_rel or seed_snapshot.has_people_content_type:
                required_flags.extend(
                    [
                        ("word/people.xml", roundtrip_snapshot.has_people),
                        ("document.xml.rels people relationship", roundtrip_snapshot.has_people_rel),
                        ("[Content_Types].xml people override", roundtrip_snapshot.has_people_content_type),
                    ]
                )

            for label, present in required_flags:
                if not present:
                    errors.append(f"Roundtrip missing state-supporting package component: {label}")

            if seed_snapshot.has_people:
                expected_authors = sorted(
                    {
                        (seed_snapshot.comments_by_id.get(cid).author or "").strip()
                        for cid in seed_snapshot.comment_ids_order
                        if seed_snapshot.comments_by_id.get(cid)
                    }
                )
                expected_authors = [author for author in expected_authors if author]
                for author in expected_authors:
                    expected_provider = (seed_snapshot.people_presence_provider_by_author.get(author) or "").strip()
                    expected_user = (seed_snapshot.people_presence_user_by_author.get(author) or "").strip()
                    if not (expected_provider or expected_user):
                        continue
                    actual_provider = (roundtrip_snapshot.people_presence_provider_by_author.get(author) or "").strip()
                    actual_user = (roundtrip_snapshot.people_presence_user_by_author.get(author) or "").strip()
                    if not actual_provider or not actual_user:
                        errors.append(
                            "Roundtrip missing people.xml presenceInfo for author "
                            f"'{author}' required for Word-compatible resolved state."
                        )
                        continue
                    if actual_provider != expected_provider or actual_user != expected_user:
                        errors.append(
                            "Roundtrip people.xml presenceInfo mismatch for author "
                            f"'{author}': expected provider/user=({expected_provider}, {expected_user}) "
                            f"actual=({actual_provider}, {actual_user})"
                        )

            introduced_state_attr_ids = sorted(
                [
                    cid
                    for cid in seed_snapshot.comment_ids_order
                    if not (seed_snapshot.comment_state_attr_by_id.get(cid) or "").strip()
                    and (roundtrip_snapshot.comment_state_attr_by_id.get(cid) or "").strip()
                ],
                key=lambda value: (len(value), value),
            )
            if introduced_state_attr_ids:
                errors.append(
                    "Roundtrip introduced unsupported comments.xml state attributes on roots: "
                    f"{introduced_state_attr_ids}"
                )

            introduced_para_attr_ids = sorted(
                [
                    cid
                    for cid in seed_snapshot.comment_ids_order
                    if not (seed_snapshot.comment_para_attr_by_id.get(cid) or "").strip()
                    and (roundtrip_snapshot.comment_para_attr_by_id.get(cid) or "").strip()
                ],
                key=lambda value: (len(value), value),
            )
            if introduced_para_attr_ids:
                errors.append(
                    "Roundtrip introduced unsupported comments.xml paraId attributes on roots: "
                    f"{introduced_para_attr_ids}"
                )

            introduced_durable_attr_ids = sorted(
                [
                    cid
                    for cid in seed_snapshot.comment_ids_order
                    if not (seed_snapshot.comment_durable_attr_by_id.get(cid) or "").strip()
                    and (roundtrip_snapshot.comment_durable_attr_by_id.get(cid) or "").strip()
                ],
                key=lambda value: (len(value), value),
            )
            if introduced_durable_attr_ids:
                errors.append(
                    "Roundtrip introduced unsupported comments.xml durableId attributes on roots: "
                    f"{introduced_durable_attr_ids}"
                )

            root_para_ids = set()
            missing_root_para_ids = []
            thread_para_mismatch_ids = []
            for cid in seed_snapshot.comment_ids_order:
                node = roundtrip_snapshot.comments_by_id.get(cid)
                para_id = node.para_id if node else ""
                last_para_id = (roundtrip_snapshot.last_paragraph_para_by_id.get(cid, "") or "").strip()
                if roundtrip_snapshot.paragraph_count_by_id.get(cid, 0) > 1 and last_para_id and para_id != last_para_id:
                    thread_para_mismatch_ids.append(cid)
                if not para_id:
                    missing_root_para_ids.append(cid)
                    continue
                root_para_ids.add(para_id)
            if thread_para_mismatch_ids:
                errors.append(
                    "Roundtrip thread paraId must match last paragraph paraId for multi-paragraph comments: "
                    f"{sorted(thread_para_mismatch_ids, key=lambda value: (len(value), value))}"
                )
            if missing_root_para_ids:
                errors.append(
                    "Roundtrip comments missing paraId mapping required for Word state resolution: "
                    f"{sorted(missing_root_para_ids, key=lambda value: (len(value), value))}"
                )

            comments_extended_para_ids = set(roundtrip_snapshot.comments_extended_para_ids)
            comments_ids_para_ids = set(roundtrip_snapshot.comments_ids_para_ids)
            missing_in_extended = sorted(
                root_para_ids - comments_extended_para_ids, key=lambda value: (len(value), value)
            )
            missing_in_ids = sorted(
                root_para_ids - comments_ids_para_ids, key=lambda value: (len(value), value)
            )
            if missing_in_extended:
                errors.append(
                    "Roundtrip comment paraIds missing from commentsExtended.xml: "
                    f"{missing_in_extended}"
                )
            if missing_in_ids:
                errors.append(
                    "Roundtrip comment paraIds missing from commentsIds.xml: "
                    f"{missing_in_ids}"
                )

            expected_parent_para_ids = sorted(
                [
                    roundtrip_snapshot.comments_by_id[parent_id].para_id
                    for parent_id in roundtrip_snapshot.parent_map.values()
                    if parent_id in roundtrip_snapshot.comments_by_id and roundtrip_snapshot.comments_by_id[parent_id].para_id
                ],
                key=lambda value: (len(value), value),
            )
            missing_parent_para_ids = sorted(
                [
                    para_id
                    for para_id in expected_parent_para_ids
                    if para_id not in set(roundtrip_snapshot.comments_extended_parent_para_ids)
                ],
                key=lambda value: (len(value), value),
            )
            if missing_parent_para_ids:
                errors.append(
                    "Roundtrip parent paraIds missing from commentsExtended.xml paraIdParent entries: "
                    f"{missing_parent_para_ids}"
                )

            if comments_ids_para_ids:
                missing_durable_para_ids = sorted(
                    [
                        para_id
                        for para_id in comments_ids_para_ids
                        if not roundtrip_snapshot.comments_ids_durable_by_para.get(para_id)
                    ],
                    key=lambda value: (len(value), value),
                )
                if missing_durable_para_ids:
                    errors.append(
                        "commentsIds.xml has entries without durableId, cannot bind to commentsExtensible.xml: "
                        f"{missing_durable_para_ids}"
                    )

            if roundtrip_snapshot.comments_ids_durable_ids:
                missing_extensible_ids = sorted(
                    set(roundtrip_snapshot.comments_ids_durable_ids)
                    - set(roundtrip_snapshot.comments_extensible_durable_ids),
                    key=lambda value: (len(value), value),
                )
                if missing_extensible_ids:
                    errors.append(
                        "Durable IDs from commentsIds.xml missing in commentsExtensible.xml: "
                        f"{missing_extensible_ids}"
                    )

        if errors:
            failure_bundle = write_failure_bundle(
                case_dir / "failure_bundle",
                original_snapshot=seed_snapshot,
                markdown_snapshot=middle_snapshot,
                roundtrip_snapshot=roundtrip_snapshot,
                expected_flatten={"case": case["name"], "expected_comment_ids": expected_comment_ids},
                command_logs=command_logs,
                errors=errors,
                text_mismatch_diffs=text_mismatch_diffs,
            )
            self.fail(
                f"Edge case failed: {case['name']}. Diagnostics: {failure_bundle}\n- " + "\n- ".join(errors)
            )
