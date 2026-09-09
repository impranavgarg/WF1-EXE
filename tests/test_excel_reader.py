from io import BytesIO

from openpyxl import Workbook

from src.excel_reader import read_sheet_names, read_source_sheet


def _source_workbook_bytes() -> bytes:
    workbook = Workbook()
    daily_sheet = workbook.active
    daily_sheet.title = "Daily Sales"
    daily_sheet.append(["Distributor", "Sales"])
    daily_sheet.append(["Acme", 125])
    workbook.create_sheet("Notes")

    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


def test_read_sheet_names_returns_workbook_order() -> None:
    assert read_sheet_names(_source_workbook_bytes()) == ["Daily Sales", "Notes"]


def test_read_source_sheet_returns_dataframe() -> None:
    dataframe = read_source_sheet(_source_workbook_bytes(), "Daily Sales")

    assert dataframe.to_dict(orient="records") == [
        {"Distributor": "Acme", "Sales": 125}
    ]
