# docx-text-diff

Compare two `.docx` files with Microsoft Word and save one output `.docx` that displays the differences as Word tracked changes.

## Requirements

- Windows 11
- Microsoft Word installed
- Python 3.11 or newer

Install dependencies:

```powershell
python -m pip install -e .
```

For tests:

```powershell
python -m pip install -e ".[dev]"
```

## Usage

```powershell
docx-text-diff original.docx revised.docx output.docx
```

Recommended Japanese text comparison:

```powershell
docx-text-diff original.docx revised.docx output.docx --granularity char
```

Useful options:

```powershell
docx-text-diff original.docx revised.docx output.docx `
  --author "Comparison" `
  --granularity word `
  --no-compare-formatting `
  --overwrite
```

The script uses Microsoft Word's `Application.CompareDocuments` COM API. Word performs the actual comparison and creates the tracked-changes markup.

## Test fixtures

The repository includes a small fixture pair:

- `tests/fixtures/original.docx`
- `tests/fixtures/revised.docx`

Try them on Windows with Word installed:

```powershell
docx-text-diff tests\fixtures\original.docx tests\fixtures\revised.docx output.docx --granularity char --overwrite
```

Regenerate the fixtures:

```powershell
python scripts\build_test_fixtures.py
```

## Notes

This tool is intended for local desktop automation. Microsoft Office desktop applications are not a good fit for unattended server-side automation.
