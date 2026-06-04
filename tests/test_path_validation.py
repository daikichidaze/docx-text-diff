import pytest

from docx_diff_track.cli import validate_paths
from docx_diff_track.errors import ValidationError


def test_validate_paths_accepts_valid_docx_files(tmp_path):
    original = tmp_path / "original.docx"
    revised = tmp_path / "revised.docx"
    output = tmp_path / "output.docx"
    original.write_bytes(b"placeholder")
    revised.write_bytes(b"placeholder")

    validate_paths(original, revised, output, overwrite=False)


def test_validate_paths_rejects_missing_input(tmp_path):
    revised = tmp_path / "revised.docx"
    output = tmp_path / "output.docx"
    revised.write_bytes(b"placeholder")

    with pytest.raises(ValidationError):
        validate_paths(tmp_path / "missing.docx", revised, output, overwrite=False)


def test_validate_paths_rejects_non_docx_input(tmp_path):
    original = tmp_path / "original.txt"
    revised = tmp_path / "revised.docx"
    output = tmp_path / "output.docx"
    original.write_bytes(b"placeholder")
    revised.write_bytes(b"placeholder")

    with pytest.raises(ValidationError):
        validate_paths(original, revised, output, overwrite=False)


def test_validate_paths_rejects_existing_output_without_overwrite(tmp_path):
    original = tmp_path / "original.docx"
    revised = tmp_path / "revised.docx"
    output = tmp_path / "output.docx"
    for path in (original, revised, output):
        path.write_bytes(b"placeholder")

    with pytest.raises(ValidationError):
        validate_paths(original, revised, output, overwrite=False)


def test_validate_paths_allows_existing_output_with_overwrite(tmp_path):
    original = tmp_path / "original.docx"
    revised = tmp_path / "revised.docx"
    output = tmp_path / "output.docx"
    for path in (original, revised, output):
        path.write_bytes(b"placeholder")

    validate_paths(original, revised, output, overwrite=True)
