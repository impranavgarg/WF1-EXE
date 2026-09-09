from datetime import date
from io import BytesIO

import pandas as pd
import pytest
from openpyxl import Workbook

from src.excel_reader import (
    clean_distributor_name,
    extract_unique_distributor_names,
    filter_distributor_rows,
    suggest_current_month_sheet,
)
from src.excel_writer import build_output_filename
from src.validators import (
    ValidationError,
    validate_destination_sheet,
    validate_destination_workbook,
    validate_distributor_column,
    validate_selected_distributor,
    validate_source_dataframe,
)


def _workbook_bytes(*sheet_names: str) -> bytes:
    workbook = Workbook()
    workbook.active.title = sheet_names[0]
    for sheet_name in sheet_names[1:]:
        workbook.create_sheet(sheet_name)

    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("  ABC Distributors  ", "ABC Distributors"),
        ("   ", None),
        (None, None),
        (float("nan"), None),
        (123, "123"),
    ],
)
def test_clean_distributor_name(value: object, expected: str | None) -> None:
    assert clean_distributor_name(value) == expected


def test_extract_unique_distributor_names_cleans_deduplicates_and_sorts() -> None:
    dataframe = pd.DataFrame(
        {
            "Distributor": [
                " Zebra Supply ",
                None,
                "Acme Foods",
                "",
                "Zebra Supply",
                "beta Wholesale",
            ]
        }
    )

    assert extract_unique_distributor_names(dataframe, "Distributor") == [
        "Acme Foods",
        "beta Wholesale",
        "Zebra Supply",
    ]


def test_filter_distributor_rows_uses_cleaned_values() -> None:
    dataframe = pd.DataFrame(
        {
            "Distributor": ["Acme", " Acme ", "Other", None],
            "Sales": [10, 20, 30, 40],
        }
    )

    result = filter_distributor_rows(dataframe, "Distributor", "Acme")

    assert result["Sales"].tolist() == [10, 20]


@pytest.mark.parametrize(
    ("distributor", "expected"),
    [
        (
            ' ACME / North:West * "Priority" ',
            "ACME_North_West_Priority_Updated_2026-09-09.xlsx",
        ),
        ("...", "Distributor_Updated_2026-09-09.xlsx"),
    ],
)
def test_build_output_filename(distributor: str, expected: str) -> None:
    assert build_output_filename(distributor, date(2026, 9, 9)) == expected


def test_validate_source_dataframe_rejects_empty_data() -> None:
    with pytest.raises(ValidationError, match="No data was found"):
        validate_source_dataframe(pd.DataFrame(columns=["Distributor"]))


def test_validate_distributor_column_accepts_existing_column() -> None:
    dataframe = pd.DataFrame({"Distributor": ["Acme"]})

    validate_distributor_column(dataframe, "Distributor")


def test_validate_distributor_column_rejects_missing_column() -> None:
    dataframe = pd.DataFrame({"Company": ["Acme"]})

    with pytest.raises(ValidationError, match="not available"):
        validate_distributor_column(dataframe, "Distributor")


def test_validate_selected_distributor_rejects_unknown_name() -> None:
    with pytest.raises(ValidationError, match="not available"):
        validate_selected_distributor("Unknown", ["Acme", "Other"])


def test_validate_destination_workbook_accepts_xlsx_bytes() -> None:
    validate_destination_workbook(_workbook_bytes("Summary"))


def test_validate_destination_workbook_rejects_invalid_bytes() -> None:
    with pytest.raises(ValidationError, match="Unable to open"):
        validate_destination_workbook(b"not an xlsx workbook")


def test_validate_destination_sheet_rejects_missing_sheet() -> None:
    with pytest.raises(ValidationError, match="could not be found"):
        validate_destination_sheet(["Summary", "September 2026"], "August 2026")


def test_suggest_current_month_sheet_prefers_year_specific_match() -> None:
    result = suggest_current_month_sheet(
        ["Summary", "September", "Sep 2026", "September 2026"],
        current_date=date(2026, 9, 9),
    )

    assert result == "September 2026"


def test_suggest_current_month_sheet_falls_back_to_first_sheet() -> None:
    assert suggest_current_month_sheet(
        ["Summary", "YTD"], current_date=date(2026, 9, 9)
    ) == "Summary"
