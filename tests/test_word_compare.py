import platform

import pytest

from docx_diff_track.errors import WordAutomationError
from docx_diff_track.word_compare import _word_granularity, compare_documents


def test_word_granularity_maps_supported_values():
    assert _word_granularity("word") == 1
    assert _word_granularity("char") == 0


def test_word_granularity_rejects_unknown_value():
    with pytest.raises(WordAutomationError):
        _word_granularity("line")


def test_compare_documents_requires_windows(tmp_path):
    if platform.system() == "Windows":
        pytest.skip("non-Windows guard is only meaningful outside Windows")

    with pytest.raises(WordAutomationError):
        compare_documents(
            tmp_path / "old.docx",
            tmp_path / "new.docx",
            tmp_path / "out.docx",
        )
