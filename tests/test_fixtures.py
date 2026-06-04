from pathlib import Path
from zipfile import ZipFile


FIXTURES = Path(__file__).parent / "fixtures"


def read_document_xml(path: Path) -> str:
    with ZipFile(path) as docx:
        return docx.read("word/document.xml").decode("utf-8")


def test_fixture_docx_files_exist():
    assert (FIXTURES / "original.docx").is_file()
    assert (FIXTURES / "revised.docx").is_file()


def test_fixture_docx_files_have_required_parts():
    for path in (FIXTURES / "original.docx", FIXTURES / "revised.docx"):
        with ZipFile(path) as docx:
            names = set(docx.namelist())

        assert "[Content_Types].xml" in names
        assert "_rels/.rels" in names
        assert "word/document.xml" in names
        assert "word/_rels/document.xml.rels" in names


def test_fixture_pair_contains_expected_text_changes():
    original_xml = read_document_xml(FIXTURES / "original.docx")
    revised_xml = read_document_xml(FIXTURES / "revised.docx")

    assert "June 10, 2026" in original_xml
    assert "June 12, 2026" in revised_xml
    assert "30 days" in original_xml
    assert "45 days" in revised_xml
    assert "金額は1000円" in original_xml
    assert "金額は1200円" in revised_xml
    assert "Additional approval is required." in revised_xml
