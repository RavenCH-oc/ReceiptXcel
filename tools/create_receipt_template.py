"""Create the internal receipt template by minimally patching the reference DOCX.

The reference file remains untouched. Existing paragraphs, runs, table geometry,
styles, and package parts are copied; only known sample values and the blank
receipt entry cells receive placeholders.
"""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def qn(local_name: str) -> str:
    return f"{{{W_NS}}}{local_name}"


def first_text(cell):
    texts = cell.xpath(".//w:t", namespaces=NS)
    if texts:
        return texts[0]

    paragraph = cell.find(".//w:p", namespaces=NS)
    if paragraph is None:
        paragraph = etree.SubElement(cell, qn("p"))

    run = etree.SubElement(paragraph, qn("r"))
    return etree.SubElement(run, qn("t"))


def replace_sample_text(root) -> None:
    replacements = {
        "中華民國112年12月30日": "中華民國{{ROC_YEAR}}年{{MONTH}}月{{DAY}}日",
        "1120051": "{{RECEIPT_NUMBER}}",
        "A實業股份有限公司": "{{PAYER}}",
        "運動會禮金": "{{REASON}}",
        "金額(大寫)：壹仟元整": "金額(大寫)：{{AMOUNT_UPPER}}",
    }

    for text in root.xpath(".//w:t", namespaces=NS):
        if not text.text:
            continue
        value = text.text
        for source, replacement in replacements.items():
            value = value.replace(source, replacement)
        text.text = value


def patch_amount_and_handler_cells(root) -> None:
    amount_names = [
        "AMOUNT_1000000",
        "AMOUNT_100000",
        "AMOUNT_10000",
        "AMOUNT_1000",
        "AMOUNT_100",
        "AMOUNT_10",
        "AMOUNT_1",
    ]

    tables = root.xpath("./w:body/w:tbl", namespaces=NS)
    if len(tables) != 3:
        raise ValueError(f"Expected three receipt tables, found {len(tables)}")

    for table in tables:
        rows = table.findall("./w:tr", namespaces=NS)
        if len(rows) != 5:
            raise ValueError("Expected each receipt table to contain five rows")

        data_cells = rows[2].findall("./w:tc", namespaces=NS)
        if len(data_cells) < 9:
            raise ValueError("Expected amount cells in the receipt data row")
        for cell_index, placeholder in enumerate(amount_names, start=2):
            first_text(data_cells[cell_index]).text = "{{" + placeholder + "}}"

        handler_cells = rows[4].findall("./w:tc", namespaces=NS)
        if len(handler_cells) < 2 or "經手人" not in "".join(
            text.text or "" for text in handler_cells[0].xpath(".//w:t", namespaces=NS)
        ):
            raise ValueError("Could not reliably locate the 經手人 row")
        first_text(handler_cells[1]).text = "{{HANDLER}}"


def build(reference_path: Path, output_path: Path) -> None:
    with zipfile.ZipFile(reference_path, "r") as source_zip:
        document_xml = source_zip.read("word/document.xml")
        parser = etree.XMLParser(remove_blank_text=False)
        root = etree.fromstring(document_xml, parser)
        replace_sample_text(root)
        patch_amount_and_handler_cells(root)
        patched_document = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w") as output_zip:
            for info in source_zip.infolist():
                data = patched_document if info.filename == "word/document.xml" else source_zip.read(info.filename)
                output_zip.writestr(info, data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("reference", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    build(args.reference, args.output)


if __name__ == "__main__":
    main()
