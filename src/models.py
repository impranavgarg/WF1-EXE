from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ReportContext:
    """Selections required to generate a distributor workbook."""

    source_filename: str
    source_sheet: str
    distributor_column: str
    distributor_name: str
    destination_filename: str
    destination_sheet: str

