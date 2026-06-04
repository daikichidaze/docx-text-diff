from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import replace
from pathlib import Path

from .errors import DocxDiffError, ValidationError
from .word_compare import CompareOptions, compare_documents


LOGGER = logging.getLogger("docx_diff_track")


def parse_bool(value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
        return False
    raise argparse.ArgumentTypeError("expected true or false")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx-text-diff",
        description=(
            "Compare two DOCX files with Microsoft Word and save a tracked-changes DOCX."
        ),
    )
    parser.add_argument("original", type=Path, help="Original DOCX path")
    parser.add_argument("revised", type=Path, help="Revised DOCX path")
    parser.add_argument("output", type=Path, help="Output DOCX path")
    parser.add_argument(
        "--granularity",
        choices=("word", "char"),
        default="word",
        help="Comparison granularity used by Word",
    )
    parser.add_argument(
        "--author",
        default="DocxDiff",
        help="Revision author shown in Word tracked changes",
    )
    parser.add_argument(
        "--compare-formatting",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="Compare formatting changes",
    )
    parser.add_argument(
        "--compare-case",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare case changes",
    )
    parser.add_argument(
        "--compare-whitespace",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare whitespace changes",
    )
    parser.add_argument(
        "--compare-tables",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare table changes",
    )
    parser.add_argument(
        "--compare-headers",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare header and footer changes",
    )
    parser.add_argument(
        "--compare-footnotes",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare footnote and endnote changes",
    )
    parser.add_argument(
        "--compare-textboxes",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare textbox changes",
    )
    parser.add_argument(
        "--compare-fields",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare field changes",
    )
    parser.add_argument(
        "--compare-comments",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare comment changes",
    )
    parser.add_argument(
        "--compare-moves",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Compare moved content",
    )
    parser.add_argument(
        "--ignore-all-comparison-warnings",
        action="store_true",
        help="Ask Word to ignore comparison warnings",
    )
    parser.add_argument(
        "--visible",
        type=parse_bool,
        default=False,
        metavar="true|false",
        help="Show the Word UI while processing",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite output file if it already exists",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )
    return parser


def validate_paths(original: Path, revised: Path, output: Path, overwrite: bool) -> None:
    for label, path in (("original", original), ("revised", revised)):
        if not path.exists():
            raise ValidationError(f"{label} file does not exist: {path}")
        if not path.is_file():
            raise ValidationError(f"{label} path is not a file: {path}")
        if path.suffix.lower() != ".docx":
            raise ValidationError(f"{label} file must be a .docx file: {path}")

    if output.suffix.lower() != ".docx":
        raise ValidationError(f"output file must be a .docx file: {output}")
    if not output.parent.exists():
        raise ValidationError(f"output directory does not exist: {output.parent}")
    if output.exists() and not overwrite:
        raise ValidationError(
            f"output file already exists: {output}. Use --overwrite to replace it."
        )


def options_from_args(args: argparse.Namespace) -> CompareOptions:
    return replace(
        CompareOptions(),
        granularity=args.granularity,
        author=args.author,
        compare_formatting=args.compare_formatting,
        compare_case=args.compare_case,
        compare_whitespace=args.compare_whitespace,
        compare_tables=args.compare_tables,
        compare_headers=args.compare_headers,
        compare_footnotes=args.compare_footnotes,
        compare_textboxes=args.compare_textboxes,
        compare_fields=args.compare_fields,
        compare_comments=args.compare_comments,
        compare_moves=args.compare_moves,
        ignore_all_comparison_warnings=args.ignore_all_comparison_warnings,
        visible=args.visible,
    )


def run(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    try:
        validate_paths(args.original, args.revised, args.output, args.overwrite)
        output_path = compare_documents(
            args.original.resolve(),
            args.revised.resolve(),
            args.output.resolve(),
            options_from_args(args),
        )
    except DocxDiffError as exc:
        LOGGER.error("%s", exc)
        return 1

    LOGGER.info("Created tracked-changes DOCX: %s", output_path)
    return 0


def main(argv: list[str] | None = None) -> int:
    return run(sys.argv[1:] if argv is None else argv)
