from pathlib import Path
from tempfile import TemporaryDirectory

from docx import Document

from extract_articles import extract, iter_articles


def test_split_and_export_two_articles():
    with TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        source = tmp / "source.docx"
        doc = Document()
        doc.add_paragraph("One", style="Heading 1")
        doc.add_paragraph("English text.")
        doc.add_paragraph("中文正文。")
        doc.add_paragraph("Two", style="Heading 1")
        doc.add_paragraph("More English.")
        doc.save(source)

        assert [title for title, _ in iter_articles(Document(source))] == ["One", "Two"]
        assert extract(source, tmp / "out", make_pdf=False) == 2
        assert len(list((tmp / "out" / "word").glob("*.docx"))) == 2


