"""Create deterministic, repository-owned Excel COM integration fixtures.

The script uses only Python's standard library. It deliberately writes a small
subset of Open XML itself so fixture creation does not depend on a local Excel
profile or user workbook. The `.xlsm` fixture is macro-enabled by content type
only and contains no VBA project, macro relationship, external link, query,
printer setting, hidden worksheet, or document-personal metadata.
"""

from __future__ import annotations

import pathlib
import zipfile


ROOT = pathlib.Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "excel-com" / "tests" / "fixtures"
EPOCH = (1980, 1, 1, 0, 0, 0)


def entry(path: str, contents: str) -> tuple[str, bytes]:
    return path, contents.encode("utf-8")


def worksheet(rows: list[list[object]]) -> str:
    cells: list[str] = []
    for row_number, row in enumerate(rows, start=1):
        row_cells: list[str] = []
        for column_number, value in enumerate(row, start=1):
            letter = chr(ord("A") + column_number - 1)
            reference = f"{letter}{row_number}"
            if isinstance(value, (int, float)):
                row_cells.append(f'<c r="{reference}"><v>{value}</v></c>')
            else:
                escaped = str(value).replace("&", "&amp;").replace("<", "&lt;")
                row_cells.append(
                    f'<c r="{reference}" t="inlineStr"><is><t>{escaped}</t></is></c>'
                )
        cells.append(f'<row r="{row_number}">{"".join(row_cells)}</row>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(cells)}</sheetData></worksheet>'
    )


def workbook(entries: list[tuple[str, bytes]], macro_enabled: bool) -> list[tuple[str, bytes]]:
    workbook_type = (
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
        if macro_enabled
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    )
    return [
        entry(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{workbook_type}"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>',
        ),
        entry(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        ),
        entry(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        ),
        entry(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>',
        ),
        entry(
            "xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>'
            '<cellXfs count="1"><xf fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '</styleSheet>',
        ),
        *entries,
    ]


def write_fixture(name: str, rows: list[list[object]], macro_enabled: bool = False) -> None:
    path = FIXTURES / name
    contents = workbook([entry("xl/worksheets/sheet1.xml", worksheet(rows))], macro_enabled)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for member, data in sorted(contents):
            info = zipfile.ZipInfo(member, EPOCH)
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data)


def main() -> None:
    FIXTURES.mkdir(parents=True, exist_ok=True)
    write_fixture("blank.xlsx", [])
    write_fixture("blank.xlsm", [], macro_enabled=True)
    write_fixture("chart-source.xlsx", [["Category", "Value"], ["Alpha", 10], ["Beta", 20], ["Gamma", 15]])
    write_fixture("pivot-source.xlsx", [["Region", "Product", "Amount"], ["North", "A", 10], ["South", "B", 20], ["North", "B", 30]])
    write_fixture("local-query-source.xlsx", [["Id", "Name"], [1, "One"], [2, "Two"]])


if __name__ == "__main__":
    main()
