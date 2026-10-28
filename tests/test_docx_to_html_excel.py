import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx_to_html_excel import blocks_to_html, convert_docx, parse_docx_blocks


DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello DOCX</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Score</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Alice</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>95</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""


class TestDocxToHtmlExcel(unittest.TestCase):
    def _create_docx(self, directory: Path) -> Path:
        docx_path = directory / "sample.docx"
        with ZipFile(docx_path, "w") as docx:
            docx.writestr("word/document.xml", DOCUMENT_XML)
        return docx_path

    def test_parse_docx_blocks_extracts_paragraph_and_table(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_path = self._create_docx(Path(temp_dir))
            blocks = parse_docx_blocks(docx_path)

        self.assertEqual(blocks[0]["type"], "paragraph")
        self.assertEqual(blocks[0]["text"], "Hello DOCX")
        self.assertEqual(blocks[1]["type"], "table")
        self.assertEqual(blocks[1]["rows"][1][0], "Alice")

    def test_blocks_to_html_renders_table(self):
        blocks = [
            {"type": "paragraph", "text": "Hello"},
            {"type": "table", "rows": [["A", "B"], ["1", "2"]]},
        ]

        html = blocks_to_html(blocks)

        self.assertIn("<p>Hello</p>", html)
        self.assertIn("<table border=\"1\">", html)
        self.assertIn("<td>2</td>", html)

    def test_convert_docx_writes_html_and_xlsx(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            docx_path = self._create_docx(temp_path)
            html_path = temp_path / "out.html"
            excel_path = temp_path / "out.xlsx"

            convert_docx(docx_path, html_path, excel_path)

            self.assertTrue(html_path.exists())
            self.assertTrue(excel_path.exists())

            html_text = html_path.read_text(encoding="utf-8")
            self.assertIn("Hello DOCX", html_text)

            with ZipFile(excel_path, "r") as xlsx:
                sheet_xml = xlsx.read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertIn("Alice", sheet_xml)
                self.assertIn("Score", sheet_xml)


if __name__ == "__main__":
    unittest.main()
