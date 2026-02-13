from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO_ROOT = Path(__file__).resolve().parents[1]
FIXTURE_DOCX = REPO_ROOT / "Preregistration_Original.docx"
SRC_ROOT = REPO_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))


class TestCliEntrypoints(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        if shutil.which("pandoc") is None:
            raise unittest.SkipTest("pandoc not found on PATH")

    def run_cmd(self, cmd, cwd=None, extra_env=None):
        env = os.environ.copy()
        if extra_env:
            env.update(extra_env)
        proc = subprocess.run(cmd, cwd=str(cwd or REPO_ROOT), capture_output=True, text=True, env=env)
        return proc.returncode, proc.stdout, proc.stderr

    def test_help_screens_exit_zero(self):
        commands = [
            ["./dmc", "--help"],
            ["./docx2md", "--help"],
            ["./md2docx", "--help"],
            ["./docx-comments", "--help"],
            ["./d2m", "--help"],
            ["./m2d", "--help"],
            ["python3", "-m", "dmc", "--help"],
        ]
        for cmd in commands:
            with self.subTest(cmd=" ".join(cmd)):
                env = None
                if cmd[:3] == ["python3", "-m", "dmc"]:
                    env = {"PYTHONPATH": str(REPO_ROOT / "src")}
                code, stdout, stderr = self.run_cmd(cmd, extra_env=env)
                self.assertEqual(
                    code,
                    0,
                    msg=f"Command failed: {' '.join(cmd)}\nstdout={stdout}\nstderr={stderr}",
                )

    def test_dmc_auto_detect_roundtrip(self):
        case_dir = Path(tempfile.mkdtemp(prefix="cli-auto-", dir="/tmp"))
        seed_md = case_dir / "seed.md"
        out_docx = case_dir / "out.docx"
        out_md = case_dir / "out.md"
        seed_md.write_text("Simple paragraph for CLI auto detection.\n", encoding="utf-8")

        code1, out1, err1 = self.run_cmd(["./dmc", str(seed_md), "-o", str(out_docx)])
        self.assertEqual(code1, 0, msg=f"dmc md->docx failed\nstdout={out1}\nstderr={err1}")
        self.assertTrue(out_docx.exists(), "Expected dmc to create DOCX output")

        code2, out2, err2 = self.run_cmd(["./dmc", str(out_docx), "-o", str(out_md)])
        self.assertEqual(code2, 0, msg=f"dmc docx->md failed\nstdout={out2}\nstderr={err2}")
        self.assertTrue(out_md.exists(), "Expected dmc to create markdown output")

    def test_dmc_subcommands_operate(self):
        case_dir = Path(tempfile.mkdtemp(prefix="cli-subcmd-", dir="/tmp"))
        source_docx = case_dir / "input.docx"
        out_md = case_dir / "from_subcommand.md"
        seed_md = case_dir / "seed.md"
        out_docx = case_dir / "from_subcommand.docx"
        shutil.copyfile(FIXTURE_DOCX, source_docx)
        seed_md.write_text("Subcommand conversion test.\n", encoding="utf-8")

        code1, out1, err1 = self.run_cmd(["./dmc", "docx2md", str(source_docx), "-o", str(out_md)])
        self.assertEqual(code1, 0, msg=f"dmc docx2md failed\nstdout={out1}\nstderr={err1}")
        self.assertTrue(out_md.exists(), "Expected dmc docx2md to create markdown output")

        code2, out2, err2 = self.run_cmd(
            ["./dmc", "md2docx", str(seed_md), "--ref", str(FIXTURE_DOCX), "-o", str(out_docx)]
        )
        self.assertEqual(code2, 0, msg=f"dmc md2docx failed\nstdout={out2}\nstderr={err2}")
        self.assertTrue(out_docx.exists(), "Expected dmc md2docx to create docx output")

    def test_ref_option_maps_to_reference_doc(self):
        from dmc import commands

        with mock.patch("dmc.commands.converter.run_conversion") as run_conversion:
            run_conversion.return_value = 0
            output = commands.run_md2docx(
                Path("draft.md"),
                output_path=Path("draft.docx"),
                reference_doc=Path("template.docx"),
                pandoc_extra_args=["--track-changes"],
            )

        self.assertEqual(output, 0)
        run_conversion.assert_called_once()
        _, kwargs = run_conversion.call_args
        self.assertEqual(kwargs, {})
        mode, input_path, output_path, extra = run_conversion.call_args.args
        self.assertEqual(mode, "md2docx")
        self.assertEqual(input_path, Path("draft.md"))
        self.assertEqual(output_path, Path("draft.docx"))
        self.assertIn("--track-changes", extra)
        self.assertIn("--reference-doc", extra)
        ref_idx = extra.index("--reference-doc")
        self.assertEqual(extra[ref_idx + 1], "template.docx")

    def test_legacy_converter_still_operates(self):
        case_dir = Path(tempfile.mkdtemp(prefix="cli-legacy-", dir="/tmp"))
        source_docx = case_dir / "input.docx"
        out_md = case_dir / "out.md"
        shutil.copyfile(FIXTURE_DOCX, source_docx)

        code, stdout, stderr = self.run_cmd(["./docx-comments", str(source_docx), "-o", str(out_md)])
        self.assertEqual(
            code,
            0,
            msg=f"docx-comments conversion failed\nstdout={stdout}\nstderr={stderr}",
        )
        self.assertTrue(out_md.exists(), "Legacy converter script did not create markdown output")


if __name__ == "__main__":
    unittest.main()
