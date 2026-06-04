from __future__ import annotations

import platform
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .errors import WordAutomationError


WD_COMPARE_TARGET_NEW = 2
WD_GRANULARITY_CHAR_LEVEL = 0
WD_GRANULARITY_WORD_LEVEL = 1
WD_FORMAT_XML_DOCUMENT = 12
WD_DO_NOT_SAVE_CHANGES = 0
WD_ALERTS_NONE = 0


@dataclass(frozen=True)
class CompareOptions:
    granularity: str = "word"
    author: str = "DocxDiff"
    compare_formatting: bool = False
    compare_case: bool = True
    compare_whitespace: bool = True
    compare_tables: bool = True
    compare_headers: bool = True
    compare_footnotes: bool = True
    compare_textboxes: bool = True
    compare_fields: bool = True
    compare_comments: bool = True
    compare_moves: bool = True
    ignore_all_comparison_warnings: bool = False
    visible: bool = False


def _word_granularity(value: str) -> int:
    if value == "word":
        return WD_GRANULARITY_WORD_LEVEL
    if value == "char":
        return WD_GRANULARITY_CHAR_LEVEL
    raise WordAutomationError(f"unsupported granularity: {value}")


def _close_document(document: Any) -> None:
    if document is not None:
        document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)


def compare_documents(
    original_path: Path,
    revised_path: Path,
    output_path: Path,
    options: CompareOptions | None = None,
) -> Path:
    """Compare DOCX files with Microsoft Word and save a tracked-changes DOCX."""

    if platform.system() != "Windows":
        raise WordAutomationError(
            "Microsoft Word COM automation requires Windows with Word installed."
        )

    options = options or CompareOptions()

    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise WordAutomationError(
            "pywin32 is required on Windows. Install it with: python -m pip install pywin32"
        ) from exc

    word = None
    original_doc = None
    revised_doc = None
    comparison_doc = None

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as exc:  # pragma: no cover - COM only
            raise WordAutomationError(
                "failed to start Microsoft Word through COM automation"
            ) from exc

        word.Visible = bool(options.visible)
        word.DisplayAlerts = WD_ALERTS_NONE

        original_doc = word.Documents.Open(
            str(original_path),
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        revised_doc = word.Documents.Open(
            str(revised_path),
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        comparison_doc = word.CompareDocuments(
            OriginalDocument=original_doc,
            RevisedDocument=revised_doc,
            Destination=WD_COMPARE_TARGET_NEW,
            Granularity=_word_granularity(options.granularity),
            CompareFormatting=options.compare_formatting,
            CompareCaseChanges=options.compare_case,
            CompareWhitespace=options.compare_whitespace,
            CompareTables=options.compare_tables,
            CompareHeaders=options.compare_headers,
            CompareFootnotes=options.compare_footnotes,
            CompareTextboxes=options.compare_textboxes,
            CompareFields=options.compare_fields,
            CompareComments=options.compare_comments,
            CompareMoves=options.compare_moves,
            RevisedAuthor=options.author,
            IgnoreAllComparisonWarnings=options.ignore_all_comparison_warnings,
        )
        comparison_doc.SaveAs2(
            str(output_path),
            FileFormat=WD_FORMAT_XML_DOCUMENT,
            AddToRecentFiles=False,
        )
        return output_path
    except WordAutomationError:
        raise
    except Exception as exc:  # pragma: no cover - COM only
        raise WordAutomationError(f"Word comparison failed: {exc}") from exc
    finally:
        for document in (comparison_doc, revised_doc, original_doc):
            try:
                _close_document(document)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except Exception:
                pass
        pythoncom.CoUninitialize()
