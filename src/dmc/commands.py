from __future__ import annotations

from pathlib import Path

from . import converter


def _append_reference_doc_arg(extra_args, reference_doc: Path | None):
    args = list(extra_args or [])
    if reference_doc is None:
        return args
    args.extend(["--reference-doc", str(reference_doc)])
    return args


def run_auto(input_path: Path, output_path: Path | None = None, pandoc_extra_args=None):
    return converter.run_conversion("auto", input_path, output_path, list(pandoc_extra_args or []))


def run_docx2md(input_path: Path, output_path: Path | None = None, pandoc_extra_args=None):
    return converter.run_conversion("docx2md", input_path, output_path, list(pandoc_extra_args or []))


def run_md2docx(
    input_path: Path,
    output_path: Path | None = None,
    reference_doc: Path | None = None,
    pandoc_extra_args=None,
):
    args = _append_reference_doc_arg(pandoc_extra_args, reference_doc)
    return converter.run_conversion("md2docx", input_path, output_path, args)
