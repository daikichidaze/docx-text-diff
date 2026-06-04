from pathlib import Path

import pytest

from docx_diff_track.cli import build_parser, options_from_args, parse_bool


def test_parse_bool_accepts_common_values():
    assert parse_bool("true") is True
    assert parse_bool("0") is False


def test_parse_bool_rejects_invalid_value():
    with pytest.raises(Exception):
        parse_bool("maybe")


def test_default_options_from_args():
    parser = build_parser()
    args = parser.parse_args(
        [str(Path("old.docx")), str(Path("new.docx")), str(Path("out.docx"))]
    )

    options = options_from_args(args)

    assert options.granularity == "word"
    assert options.author == "DocxDiff"
    assert options.compare_formatting is False
    assert options.compare_tables is True
    assert options.visible is False


def test_cli_options_override_defaults():
    parser = build_parser()
    args = parser.parse_args(
        [
            "old.docx",
            "new.docx",
            "out.docx",
            "--granularity",
            "char",
            "--author",
            "Reviewer",
            "--compare-formatting",
            "--no-compare-comments",
            "--visible",
            "true",
        ]
    )

    options = options_from_args(args)

    assert options.granularity == "char"
    assert options.author == "Reviewer"
    assert options.compare_formatting is True
    assert options.compare_comments is False
    assert options.visible is True
