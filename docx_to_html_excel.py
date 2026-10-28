import argparse
import html
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def parse_docx_blocks(docx_path: Path):
    with ZipFile(docx_path, "r") as docx_zip:
        document_xml = docx_zip.read("word/document.xml")

    root = ET.fromstring(document_xml)
    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        return []

    blocks = []
    for child in body:
        if child.tag == f"{{{W_NS}}}p":
            texts = [node.text for node in child.findall(f".//{{{W_NS}}}t") if node.text]
            paragraph = "".join(texts).strip()
            if paragraph:
                blocks.append({"type": "paragraph", "text": paragraph})
        elif child.tag == f"{{{W_NS}}}tbl":
            rows = []
            for row in child.findall(f"{{{W_NS}}}tr"):
                cells = []
                for cell in row.findall(f"{{{W_NS}}}tc"):
                    cell_texts = [node.text for node in cell.findall(f".//{{{W_NS}}}t") if node.text]
                    cells.append("".join(cell_texts).strip())
                if cells:
                    rows.append(cells)
            if rows:
                blocks.append({"type": "table", "rows": rows})

    return blocks


def blocks_to_html(blocks):
    parts = ["<html><body>"]
    for block in blocks:
        if block["type"] == "paragraph":
            parts.append(f"<p>{html.escape(block['text'])}</p>")
        elif block["type"] == "table":
            parts.append("<table border=\"1\">")
            for row in block["rows"]:
                parts.append("<tr>")
                for cell in row:
                    parts.append(f"<td>{html.escape(cell)}</td>")
                parts.append("</tr>")
            parts.append("</table>")
    parts.append("</body></html>")
    return "".join(parts)


def _col_letter(index: int) -> str:
    letters = []
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def _sheet_xml_from_tables(tables):
    rows_xml = []
    row_num = 1

    if not tables:
        tables = [["No tables found"]]

    for table_index, table in enumerate(tables):
        for row in table:
            cells_xml = []
            for col_index, value in enumerate(row, start=1):
                cell_ref = f"{_col_letter(col_index)}{row_num}"
                escaped = html.escape(value or "")
                cells_xml.append(
                    f"<c r=\"{cell_ref}\" t=\"inlineStr\"><is><t>{escaped}</t></is></c>"
                )
            rows_xml.append(f"<row r=\"{row_num}\">{''.join(cells_xml)}</row>")
            row_num += 1

        if table_index < len(tables) - 1:
            row_num += 1

    worksheet_open = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheetData>"
    )
    worksheet_close = "</sheetData></worksheet>"
    return f"{worksheet_open}{''.join(rows_xml)}{worksheet_close}"


def blocks_to_xlsx(blocks, excel_path: Path):
    tables = [block["rows"] for block in blocks if block["type"] == "table"]

    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
  <Default Extension=\"xml\" ContentType=\"application/xml\"/>
  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
</Types>
""".strip()

    rels = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>
""".strip()

    workbook = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
  <sheets>
    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>
  </sheets>
</workbook>
""".strip()

    workbook_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
</Relationships>
""".strip()

    sheet_xml = _sheet_xml_from_tables(tables)

    with ZipFile(excel_path, "w", compression=ZIP_DEFLATED) as xlsx:
        xlsx.writestr("[Content_Types].xml", content_types)
        xlsx.writestr("_rels/.rels", rels)
        xlsx.writestr("xl/workbook.xml", workbook)
        xlsx.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        xlsx.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def convert_docx(docx_path: Path, html_path: Path, excel_path: Path):
    blocks = parse_docx_blocks(docx_path)
    html_path.write_text(blocks_to_html(blocks), encoding="utf-8")
    blocks_to_xlsx(blocks, excel_path)


def main():
    parser = argparse.ArgumentParser(description="Convert DOCX (Word ZIP) to HTML and Excel.")
    parser.add_argument("input", type=Path, help="Path to input .docx file")
    parser.add_argument("--html-out", type=Path, help="Path to output HTML file")
    parser.add_argument("--excel-out", type=Path, help="Path to output XLSX file")
    args = parser.parse_args()

    if args.input.suffix.lower() != ".docx":
        raise SystemExit("Input file must be a .docx file")

    html_out = args.html_out or args.input.with_suffix(".html")
    excel_out = args.excel_out or args.input.with_suffix(".xlsx")

    convert_docx(args.input, html_out, excel_out)


if __name__ == "__main__":
    main()
