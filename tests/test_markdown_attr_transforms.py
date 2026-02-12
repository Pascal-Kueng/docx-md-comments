from __future__ import annotations

import re
import runpy
import shutil
import tempfile
import unittest
from pathlib import Path

from tests.helpers.markdown_inspector import extract_comment_start_attrs

REPO_ROOT = Path(__file__).resolve().parents[1]
CONVERTER_PATH = REPO_ROOT / "docx-comments"


class TestMarkdownAttrTransforms(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if shutil.which("pandoc") is None:
            raise unittest.SkipTest("pandoc not found on PATH")
        if not CONVERTER_PATH.exists():
            raise unittest.SkipTest(f"converter script not found: {CONVERTER_PATH}")
        cls.converter_mod = runpy.run_path(str(CONVERTER_PATH))

    def test_annotate_comment_attrs_does_not_touch_code_literals(self) -> None:
        annotate = self.converter_mod["annotate_markdown_comment_attrs"]
        work_dir = Path(tempfile.mkdtemp(prefix="ast-annotate-", dir="/tmp"))
        md_path = work_dir / "input.md"
        md_path.write_text(
            (
                "```text\n"
                "FAKE {.comment-start id=\"999\"}\n"
                "```\n\n"
                "Real [anchor]{.comment-start id=\"10\" author=\"A\" date=\"2026-01-01T00:00:00Z\"}"
                " text[]{.comment-end id=\"10\"}.\n"
            ),
            encoding="utf-8",
        )

        changed = annotate(
            md_path,
            parent_map={},
            state_by_id={"10": "resolved"},
            para_by_id={"10": "AAAABBBB"},
            durable_by_id={"10": "CCCCDDDD"},
            presence_provider_by_id={},
            presence_user_by_id={},
            pandoc_extra_args=None,
            writer_format="markdown",
            cwd=work_dir,
        )
        self.assertGreater(changed, 0)

        output = md_path.read_text(encoding="utf-8")
        self.assertIn('FAKE {.comment-start id="999"}', output)

        attrs = extract_comment_start_attrs(md_path)
        self.assertEqual(attrs["10"].get("state"), "resolved")
        self.assertEqual(attrs["10"].get("paraId"), "AAAABBBB")
        self.assertEqual(attrs["10"].get("durableId"), "CCCCDDDD")

    def test_strip_transport_attrs_does_not_touch_code_literals(self) -> None:
        strip_ast = self.converter_mod["strip_comment_transport_attrs_ast"]
        work_dir = Path(tempfile.mkdtemp(prefix="ast-strip-", dir="/tmp"))
        in_md = work_dir / "input.md"
        out_md = work_dir / "output.md"
        in_md.write_text(
            (
                "```text\n"
                "FAKE {.comment-start id=\"999\" paraId=\"BAD1\" durableId=\"BAD2\" "
                "presenceProvider=\"BAD3\" presenceUserId=\"BAD4\"}\n"
                "```\n\n"
                "Real [anchor]{.comment-start id=\"20\" author=\"A\" date=\"2026-01-01T00:00:00Z\" "
                "paraId=\"AAAA0001\" durableId=\"BBBB0002\" presenceProvider=\"AD\" "
                "presenceUserId=\"USR\"} text[]{.comment-end id=\"20\"}.\n"
            ),
            encoding="utf-8",
        )

        strip_ast(
            in_md,
            out_md,
            pandoc_extra_args=None,
            writer_format="markdown",
            cwd=work_dir,
        )

        output = out_md.read_text(encoding="utf-8")
        self.assertIn(
            'FAKE {.comment-start id="999" paraId="BAD1" durableId="BAD2" '
            'presenceProvider="BAD3" presenceUserId="BAD4"}',
            output,
        )

        attrs = extract_comment_start_attrs(out_md)
        self.assertNotIn("paraId", attrs["20"])
        self.assertNotIn("durableId", attrs["20"])
        self.assertNotIn("presenceProvider", attrs["20"])
        self.assertNotIn("presenceUserId", attrs["20"])

    def test_writer_passthrough_helpers(self) -> None:
        resolve_writer = self.converter_mod["resolve_pandoc_writer_format"]
        render_args = self.converter_mod["pandoc_args_for_json_markdown_render"]

        self.assertEqual(resolve_writer(["-t", "commonmark"], default_format="markdown"), "commonmark")
        self.assertEqual(resolve_writer(["--to=gfm"], default_format="markdown"), "gfm")
        self.assertEqual(resolve_writer([], default_format="markdown"), "markdown")

        filtered = render_args(
            [
                "-t",
                "commonmark",
                "--extract-media=.",
                "--output=out.md",
                "--wrap=none",
                "--columns=120",
            ]
        )
        self.assertIn("--wrap=none", filtered)
        self.assertIn("--columns=120", filtered)
        self.assertNotIn("-t", filtered)
        self.assertTrue(all(not arg.startswith("--to") for arg in filtered))
        self.assertTrue(all(not arg.startswith("--output") for arg in filtered))
        self.assertTrue(all(not arg.startswith("--extract-media") for arg in filtered))

    def test_milestone_tokens_expand_with_flexible_spacing(self) -> None:
        normalize_tokens = self.converter_mod["normalize_milestone_tokens_ast"]
        work_dir = Path(tempfile.mkdtemp(prefix="ast-milestone-", dir="/tmp"))
        in_md = work_dir / "input.md"
        out_md = work_dir / "output.md"
        in_md.write_text(
            (
                "```text\n"
                "literal ///c99.START/// should stay untouched in code\n"
                "```\n\n"
                "Start ==/// c1 . START ///==alpha==///c1.eNd///== and ///c2.s///beta/// c2 . E ///.\n\n"
                "Regular ==highlight text== should remain highlight.\n"
            ),
            encoding="utf-8",
        )

        replaced, _ = normalize_tokens(
            in_md,
            out_md,
            pandoc_extra_args=None,
            writer_format="markdown",
            cwd=work_dir,
        )
        self.assertEqual(replaced, 4)

        output = out_md.read_text(encoding="utf-8")
        self.assertIn("literal ///c99.START/// should stay untouched in code", output)
        self.assertNotIn("/// c1 . START ///", output)
        self.assertNotIn("///c2.s///", output)
        self.assertNotIn("==[]{.comment-start", output)
        self.assertNotIn("==[]{.comment-end", output)
        self.assertIn("Regular ==highlight text== should remain highlight.", output)

        attrs = extract_comment_start_attrs(out_md)
        self.assertIn("c1", attrs)
        self.assertIn("c2", attrs)
        self.assertIn('[]{.comment-end id="c1"}', output)
        self.assertIn('[]{.comment-end id="c2"}', output)

    def test_marker_validation_reports_missing_root_end_with_line(self) -> None:
        validate_markers = self.converter_mod["validate_comment_marker_integrity"]
        source = (
            "==///C11.START///== alpha\n\n"
            "> [!COMMENT 11: Alice (active)]\n"
            '> <!--CARD_META{#11 "author":"Alice","date":"2026-01-01T00:00:00Z","state":"active"}-->\n'
            "> body\n"
        )
        normalized = '[]{.comment-start id="11"} alpha\n'
        with self.assertRaises(ValueError) as ctx:
            validate_markers(
                source,
                normalized,
                card_by_id={"11": {"author": "Alice", "state": "active"}},
                source_label="broken.md",
            )
        message = str(ctx.exception)
        self.assertIn("broken.md", message)
        self.assertIn("Root comment 11", message)
        self.assertIn("START=1, END=0", message)
        self.assertIn("line 1", message)

    def test_marker_validation_does_not_flag_regular_highlight(self) -> None:
        validate_markers = self.converter_mod["validate_comment_marker_integrity"]
        source = "Regular ==highlight== stays as-is.\n"
        normalized = source
        validate_markers(source, normalized, card_by_id={}, source_label="ok.md")

    def test_marker_validation_rejects_one_sided_wrapper(self) -> None:
        validate_markers = self.converter_mod["validate_comment_marker_integrity"]
        source = (
            "==///C11.START/// broken anchor ///C11.END///\n\n"
            "> [!COMMENT 11: Alice (active)]\n"
            '> <!--CARD_META{#11 "author":"Alice","date":"2026-01-01T00:00:00Z","state":"active"}-->\n'
            "> body\n"
        )
        normalized = '[]{.comment-start id="11"} broken anchor []{.comment-end id="11"}\n'
        with self.assertRaises(ValueError) as ctx:
            validate_markers(
                source,
                normalized,
                card_by_id={"11": {"author": "Alice", "state": "active"}},
                source_label="one-sided.md",
            )
        message = str(ctx.exception)
        self.assertIn("one-sided.md", message)
        self.assertIn("one-sided highlight wrapper", message)
        self.assertIn("Comment 11", message)

    def test_marker_validation_allows_plain_or_balanced_wrappers(self) -> None:
        validate_markers = self.converter_mod["validate_comment_marker_integrity"]
        normalized = '[]{.comment-start id="11"} blabla []{.comment-end id="11"}\n'

        plain = (
            "///C11.START/// blabla ///C11.END///\n\n"
            "> [!COMMENT 11: Alice (active)]\n"
            '> <!--CARD_META{#11 "author":"Alice","date":"2026-01-01T00:00:00Z","state":"active"}-->\n'
            "> body\n"
        )
        validate_markers(
            plain,
            normalized,
            card_by_id={"11": {"author": "Alice", "state": "active"}},
            source_label="plain.md",
        )

        wrapped = (
            "====///C11.START///== blabla ==///C11.END///====\n\n"
            "> [!COMMENT 11: Alice (active)]\n"
            '> <!--CARD_META{#11 "author":"Alice","date":"2026-01-01T00:00:00Z","state":"active"}-->\n'
            "> body\n"
        )
        validate_markers(
            wrapped,
            normalized,
            card_by_id={"11": {"author": "Alice", "state": "active"}},
            source_label="wrapped.md",
        )

    def test_comment_cards_are_inserted_after_anchor_paragraph(self) -> None:
        emit_cards = self.converter_mod["emit_milestones_and_cards_ast"]
        run_pandoc_json = self.converter_mod["run_pandoc_json"]

        work_dir = Path(tempfile.mkdtemp(prefix="ast-card-placement-", dir="/tmp"))
        md_path = work_dir / "input.md"
        md_path.write_text(
            (
                "[Alpha]{.comment-start id=\"c1\" author=\"Alice\" date=\"2026-01-01T00:00:00Z\"} "
                "text[]{.comment-end id=\"c1\"}.\n\n"
                "Second paragraph.\n"
            ),
            encoding="utf-8",
        )

        changed, card_count = emit_cards(
            md_path,
            comment_cards_by_id={
                "c1": {
                    "author": "Alice",
                    "date": "2026-01-01T00:00:00Z",
                    "state": "active",
                    "text": "Comment body",
                }
            },
            child_ids=set(),
            pandoc_extra_args=None,
            writer_format="markdown",
            cwd=work_dir,
        )
        self.assertGreater(changed, 0)
        self.assertEqual(card_count, 1)

        doc = run_pandoc_json(md_path, fmt_from="markdown", extra_args=None)
        blocks = [b for b in doc.get("blocks", []) if isinstance(b, dict)]
        self.assertGreaterEqual(len(blocks), 3)
        self.assertEqual(blocks[0].get("t"), "Para")
        self.assertEqual(blocks[1].get("t"), "BlockQuote")
        self.assertEqual(blocks[2].get("t"), "Para")
        rendered = md_path.read_text(encoding="utf-8")
        self.assertIn("!COMMENT c1: Alice (active)", rendered)
        self.assertIn('<!--CARD_META{#c1 "author":"Alice","date":"2026-01-01T00:00:00Z","state":"active"}-->', rendered)
        self.assertRegex(rendered, r'> \[!COMMENT c1: Alice \(active\)\]\n> <!--CARD_META\{#c1[^}]*\}-->')
        self.assertNotRegex(rendered, r'> \[!COMMENT c1: Alice \(active\)\]\n>[ \t]*\n>[ \t]*<!--CARD_META')

    def test_reply_markers_move_to_cards_only(self) -> None:
        emit_cards = self.converter_mod["emit_milestones_and_cards_ast"]
        normalize_tokens = self.converter_mod["normalize_milestone_tokens_ast"]
        extract_comments = self.converter_mod["extract_comment_texts_from_markdown"]

        work_dir = Path(tempfile.mkdtemp(prefix="ast-reply-cards-", dir="/tmp"))
        md_path = work_dir / "input.md"
        normalized_path = work_dir / "normalized.md"
        md_path.write_text(
            (
                "[Root]{.comment-start id=\"c1\" author=\"Alice\" date=\"2026-01-01T00:00:00Z\"}"
                " body "
                "[Reply anchor]{.comment-start id=\"c2\" author=\"Bob\" date=\"2026-01-01T01:00:00Z\" parent=\"c1\"}"
                "text[]{.comment-end id=\"c2\"}"
                " end[]{.comment-end id=\"c1\"}.\n"
            ),
            encoding="utf-8",
        )

        changed, _ = emit_cards(
            md_path,
            comment_cards_by_id={
                "c1": {
                    "author": "Alice",
                    "date": "2026-01-01T00:00:00Z",
                    "state": "active",
                    "text": "Root body",
                },
                "c2": {
                    "author": "Bob",
                    "date": "2026-01-01T01:00:00Z",
                    "parent": "c1",
                    "state": "active",
                    "text": "Reply body",
                },
            },
            child_ids={"c2"},
            pandoc_extra_args=None,
            writer_format="markdown",
            cwd=work_dir,
        )
        self.assertGreater(changed, 0)

        emitted = md_path.read_text(encoding="utf-8")
        self.assertIn("==///c1.START///==", emitted)
        self.assertIn("==///c1.END///==", emitted)
        self.assertIn("///c1.START///", emitted)
        self.assertIn("///c1.END///", emitted)
        self.assertNotIn("///c2.START///", emitted)
        self.assertNotIn("///c2.END///", emitted)
        self.assertIn("> [!COMMENT c1: Alice (active)]", emitted)
        self.assertIn("[!REPLY c2: Bob (active)]", emitted)
        self.assertIn('<!--CARD_META{#c2 "author":"Bob","date":"2026-01-01T01:00:00Z","parent":"c1","state":"active"}-->', emitted)

        replaced, card_by_id = normalize_tokens(
            md_path,
            normalized_path,
            pandoc_extra_args=None,
            writer_format="markdown",
            cwd=work_dir,
        )
        self.assertGreaterEqual(replaced, 2)
        self.assertIn("c2", card_by_id)
        self.assertEqual(card_by_id["c2"].get("parent"), "c1")

        comment_data = extract_comments(normalized_path, pandoc_extra_args=None, card_by_id=card_by_id)
        self.assertIn("c2", comment_data["child_ids"])
        self.assertEqual(comment_data["parent_by_id"].get("c2"), "c1")
        self.assertIn("Reply from: Bob", comment_data["flattened_by_id"].get("c1", ""))

    def test_root_end_marker_stays_before_card_block(self) -> None:
        emit_cards = self.converter_mod["emit_milestones_and_cards_ast"]

        work_dir = Path(tempfile.mkdtemp(prefix="ast-end-before-card-", dir="/tmp"))
        md_path = work_dir / "input.md"
        md_path.write_text(
            (
                "[Alpha]{.comment-start id=\"c1\" author=\"Alice\" date=\"2026-01-01T00:00:00Z\"}"
                " first paragraph.\n\n"
                "Second paragraph[]{.comment-end id=\"c1\"}.\n"
            ),
            encoding="utf-8",
        )

        changed, card_count = emit_cards(
            md_path,
            comment_cards_by_id={
                "c1": {
                    "author": "Alice",
                    "date": "2026-01-01T00:00:00Z",
                    "state": "active",
                    "text": "Card body",
                }
            },
            child_ids=set(),
            pandoc_extra_args=None,
            writer_format="markdown",
            cwd=work_dir,
        )
        self.assertGreater(changed, 0)
        self.assertEqual(card_count, 1)

        output = md_path.read_text(encoding="utf-8")
        end_pos = output.find("///c1.END///")
        card_pos = output.find("CARD_META{#c1")
        self.assertNotEqual(end_pos, -1)
        self.assertNotEqual(card_pos, -1)
        self.assertLess(end_pos, card_pos)

    def test_extract_comments_rejects_unknown_parent(self) -> None:
        extract_comments = self.converter_mod["extract_comment_texts_from_markdown"]
        work_dir = Path(tempfile.mkdtemp(prefix="ast-parent-unknown-", dir="/tmp"))
        md_path = work_dir / "input.md"
        md_path.write_text(
            (
                "==///c1.START///==anchor==///c1.END///==\n\n"
                "> [!COMMENT c1: Alice (active)]\n"
                '> <!--CARD_META{#c1 "author":"Alice","date":"2026-01-01T00:00:00Z","state":"active"}-->\n'
                "> root body\n"
            ),
            encoding="utf-8",
        )

        with self.assertRaises(ValueError) as ctx:
            extract_comments(
                md_path,
                pandoc_extra_args=None,
                card_by_id={
                    "c1": {"author": "Alice", "state": "active", "text": "root body"},
                    "c2": {"author": "Bob", "state": "active", "parent": "c99", "text": "reply body"},
                },
            )
        message = str(ctx.exception)
        self.assertIn("unknown parent", message)
        self.assertIn("c99", message)

    def test_extract_comments_rejects_parent_cycle(self) -> None:
        extract_comments = self.converter_mod["extract_comment_texts_from_markdown"]
        work_dir = Path(tempfile.mkdtemp(prefix="ast-parent-cycle-", dir="/tmp"))
        md_path = work_dir / "input.md"
        md_path.write_text("Cycle test.\n", encoding="utf-8")

        with self.assertRaises(ValueError) as ctx:
            extract_comments(
                md_path,
                pandoc_extra_args=None,
                card_by_id={
                    "c1": {"author": "Alice", "state": "active", "parent": "c2", "text": "A"},
                    "c2": {"author": "Bob", "state": "active", "parent": "c1", "text": "B"},
                },
            )
        message = str(ctx.exception)
        self.assertIn("cycle", message)
        self.assertIn("c1", message)
        self.assertIn("c2", message)
