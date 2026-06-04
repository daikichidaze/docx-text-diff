class DocxDiffError(Exception):
    """Base class for user-facing errors."""


class ValidationError(DocxDiffError):
    """Raised when input or output paths are invalid."""


class WordAutomationError(DocxDiffError):
    """Raised when Microsoft Word automation fails."""
