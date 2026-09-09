import hashlib
import logging
from collections.abc import Hashable, Sequence
from dataclasses import dataclass

import pandas as pd
import streamlit as st

from src.excel_reader import (
    ExcelReadError,
    extract_unique_distributor_names,
    filter_distributor_rows,
    read_sheet_names,
    read_source_sheet,
    suggest_current_month_sheet,
)
from src.excel_writer import build_output_filename
from src.models import ReportContext
from src.transformer import TransformationError, transform_distributor_report
from src.validators import (
    ValidationError,
    validate_destination_sheet,
    validate_destination_workbook,
    validate_distributor_column,
    validate_selected_distributor,
    validate_source_dataframe,
)


LOGGER = logging.getLogger(__name__)
INVALID_SOURCE_MESSAGE = (
    "Unable to read this Excel file. Please upload a valid .xlsx workbook."
)


@dataclass(frozen=True, slots=True)
class SourceWorkbookState:
    """Validated source workbook choices needed by the distributor step."""

    filename: str
    workbook_bytes: bytes
    signature: str
    sheet_name: str
    distributor_column: Hashable
    dataframe: pd.DataFrame
    distributor_names: list[str]


@dataclass(frozen=True, slots=True)
class SourceSelection:
    """Selected distributor and its filtered source data."""

    workbook: SourceWorkbookState
    distributor_name: str
    filtered_dataframe: pd.DataFrame


@dataclass(frozen=True, slots=True)
class DestinationSelection:
    """Validated destination workbook and sheet choice."""

    filename: str
    workbook_bytes: bytes
    signature: str
    sheet_name: str


@st.cache_data(show_spinner=False)
def _cached_sheet_names(workbook_bytes: bytes) -> list[str]:
    return read_sheet_names(workbook_bytes)


@st.cache_data(show_spinner=False)
def _cached_source_sheet(
    workbook_bytes: bytes, sheet_name: str
) -> pd.DataFrame:
    return read_source_sheet(workbook_bytes, sheet_name)


def _file_signature(filename: str, content: bytes) -> str:
    digest = hashlib.sha256(content).hexdigest()
    return f"{filename}:{len(content)}:{digest}"


def _clear_state(keys: Sequence[str]) -> None:
    for key in keys:
        st.session_state.pop(key, None)


def _show_workbook_details(filename: str, sheet_names: list[str]) -> None:
    with st.container(border=True):
        filename_column, count_column = st.columns([3, 1])
        filename_column.caption("Workbook")
        filename_column.write(filename)
        count_column.caption("Sheets")
        count_column.write(len(sheet_names))
        st.caption("Available sheets")
        st.text("\n".join(f"• {sheet_name}" for sheet_name in sheet_names))


def _show_review(context: ReportContext, row_count: int) -> None:
    with st.container(border=True):
        left_column, right_column = st.columns(2)
        with left_column:
            st.caption("Source File")
            st.write(context.source_filename)
            st.caption("Source Sheet")
            st.write(context.source_sheet)
            st.caption("Distributor")
            st.write(context.distributor_name)
        with right_column:
            st.caption("Rows Found")
            st.write(row_count)
            st.caption("Destination File")
            st.write(context.destination_filename)
            st.caption("Destination Sheet")
            st.write(context.destination_sheet)


def _selection_signature(
    source_signature: str,
    source_sheet: str,
    distributor_column: Hashable,
    distributor_name: str,
    destination_signature: str,
    destination_sheet: str,
) -> str:
    values = (
        source_signature,
        source_sheet,
        repr(distributor_column),
        distributor_name,
        destination_signature,
        destination_sheet,
    )
    return hashlib.sha256(repr(values).encode("utf-8")).hexdigest()


def _render_source_workbook() -> SourceWorkbookState | None:
    st.subheader("1. Upload today's distributor file")
    source_upload = st.file_uploader(
        "Daily distributor Excel file",
        type=["xlsx"],
        key="source_upload",
    )
    if source_upload is None:
        st.info("Upload a .xlsx workbook to begin.")
        return

    source_bytes = source_upload.getvalue()
    source_signature = _file_signature(source_upload.name, source_bytes)
    if st.session_state.get("_source_signature") != source_signature:
        _clear_state(
            (
                "source_sheet",
                "distributor_column",
                "selected_distributor",
                "_last_source_sheet",
                "_last_distributor_column",
                "generated_workbook",
                "generated_filename",
            )
        )
        st.session_state["_source_signature"] = source_signature

    try:
        source_sheet_names = _cached_sheet_names(source_bytes)
    except ExcelReadError:
        LOGGER.exception("Unable to read source workbook %s", source_upload.name)
        st.error(INVALID_SOURCE_MESSAGE)
        return

    if not source_sheet_names:
        st.error("No data was found in the selected sheet.")
        return

    _show_workbook_details(source_upload.name, source_sheet_names)
    if st.session_state.get("source_sheet") not in source_sheet_names:
        st.session_state["source_sheet"] = source_sheet_names[0]
    source_sheet = st.selectbox(
        "Source sheet",
        source_sheet_names,
        key="source_sheet",
    )

    if st.session_state.get("_last_source_sheet") != source_sheet:
        _clear_state(("distributor_column", "selected_distributor"))
        st.session_state["_last_source_sheet"] = source_sheet

    try:
        source_df = _cached_source_sheet(source_bytes, source_sheet)
        validate_source_dataframe(source_df)
    except ExcelReadError:
        LOGGER.exception(
            "Unable to read source sheet %s from %s",
            source_sheet,
            source_upload.name,
        )
        st.error(INVALID_SOURCE_MESSAGE)
        return
    except ValidationError as exc:
        st.error(str(exc))
        return

    source_columns = list(source_df.columns)
    if st.session_state.get("distributor_column") not in source_columns:
        st.session_state["distributor_column"] = source_columns[0]
    distributor_column = st.selectbox(
        "Which column contains the distributor name?",
        source_columns,
        format_func=str,
        key="distributor_column",
    )

    if st.session_state.get("_last_distributor_column") != distributor_column:
        _clear_state(("selected_distributor",))
        st.session_state["_last_distributor_column"] = distributor_column

    try:
        validate_distributor_column(source_df, distributor_column)
        distributors = extract_unique_distributor_names(
            source_df, distributor_column
        )
    except (ValidationError, KeyError) as exc:
        st.error(str(exc))
        return

    if not distributors:
        st.error("No distributors were found in the selected column.")
        return

    return SourceWorkbookState(
        filename=source_upload.name,
        workbook_bytes=source_bytes,
        signature=source_signature,
        sheet_name=source_sheet,
        distributor_column=distributor_column,
        dataframe=source_df,
        distributor_names=distributors,
    )


def _render_distributor_selection(
    source: SourceWorkbookState,
) -> SourceSelection | None:
    st.subheader("2. Select distributor")
    if st.session_state.get("selected_distributor") not in [
        None,
        *source.distributor_names,
    ]:
        st.session_state.pop("selected_distributor", None)
    selected_distributor = st.selectbox(
        "Select distributor",
        source.distributor_names,
        index=None,
        placeholder="Choose a distributor",
        key="selected_distributor",
    )
    if selected_distributor is None:
        st.info("Select a distributor to continue.")
        return

    try:
        validate_selected_distributor(
            selected_distributor, source.distributor_names
        )
    except ValidationError as exc:
        st.error(str(exc))
        return

    filtered_source_df = filter_distributor_rows(
        source.dataframe,
        source.distributor_column,
        selected_distributor,
    )
    st.success(
        f"Found {len(filtered_source_df)} rows for {selected_distributor}."
    )
    with st.expander("View source data"):
        st.dataframe(filtered_source_df, use_container_width=True)

    return SourceSelection(
        workbook=source,
        distributor_name=selected_distributor,
        filtered_dataframe=filtered_source_df,
    )


def _render_destination_selection() -> DestinationSelection | None:
    st.subheader("3. Upload workbook to update")
    destination_upload = st.file_uploader(
        "Existing distributor workbook",
        type=["xlsx"],
        key="destination_upload",
    )
    if destination_upload is None:
        st.info("Upload the existing distributor workbook to continue.")
        return

    destination_bytes = destination_upload.getvalue()
    destination_signature = _file_signature(
        destination_upload.name, destination_bytes
    )
    if st.session_state.get("_destination_signature") != destination_signature:
        _clear_state(
            (
                "destination_sheet",
                "generated_workbook",
                "generated_filename",
            )
        )
        st.session_state["_destination_signature"] = destination_signature

    try:
        validate_destination_workbook(destination_bytes)
        destination_sheet_names = _cached_sheet_names(destination_bytes)
    except (ValidationError, ExcelReadError):
        LOGGER.exception(
            "Unable to open destination workbook %s", destination_upload.name
        )
        st.error("Unable to open the destination workbook.")
        return

    _show_workbook_details(destination_upload.name, destination_sheet_names)

    st.subheader("4. Select sheet to update")
    preferred_sheet = suggest_current_month_sheet(destination_sheet_names)
    if st.session_state.get("destination_sheet") not in destination_sheet_names:
        st.session_state["destination_sheet"] = preferred_sheet
    destination_sheet = st.selectbox(
        "Destination sheet",
        destination_sheet_names,
        key="destination_sheet",
    )

    try:
        validate_destination_sheet(destination_sheet_names, destination_sheet)
    except ValidationError as exc:
        st.error(str(exc))
        return

    return DestinationSelection(
        filename=destination_upload.name,
        workbook_bytes=destination_bytes,
        signature=destination_signature,
        sheet_name=destination_sheet,
    )


def _render_review_and_generation(
    source: SourceSelection,
    destination: DestinationSelection,
) -> None:
    context = ReportContext(
        source_filename=source.workbook.filename,
        source_sheet=source.workbook.sheet_name,
        distributor_column=str(source.workbook.distributor_column),
        distributor_name=source.distributor_name,
        destination_filename=destination.filename,
        destination_sheet=destination.sheet_name,
    )

    st.subheader("5. Review")
    _show_review(context, len(source.filtered_dataframe))
    with st.expander("Preview distributor data"):
        st.dataframe(source.filtered_dataframe, use_container_width=True)

    generation_signature = _selection_signature(
        source.workbook.signature,
        source.workbook.sheet_name,
        source.workbook.distributor_column,
        source.distributor_name,
        destination.signature,
        destination.sheet_name,
    )
    if st.session_state.get("_generation_signature") != generation_signature:
        _clear_state(("generated_workbook", "generated_filename"))
        st.session_state["_generation_signature"] = generation_signature

    st.subheader("6. Generate")
    ready_to_generate = all(
        (
            source.workbook.workbook_bytes,
            source.workbook.sheet_name,
            source.workbook.distributor_column is not None,
            source.distributor_name,
            destination.workbook_bytes,
            destination.sheet_name,
        )
    )
    if st.button(
        "Generate Updated Excel",
        type="primary",
        disabled=not ready_to_generate,
        use_container_width=True,
    ):
        try:
            with st.spinner("Generating updated workbook..."):
                st.session_state["generated_workbook"] = (
                    transform_distributor_report(
                        source_df=source.filtered_dataframe,
                        selected_distributor=source.distributor_name,
                        destination_workbook_bytes=destination.workbook_bytes,
                        destination_sheet_name=destination.sheet_name,
                    )
                )
                st.session_state["generated_filename"] = build_output_filename(
                    source.distributor_name
                )
        except TransformationError as exc:
            LOGGER.exception(
                "Workbook generation failed for distributor %s",
                source.distributor_name,
            )
            st.error(str(exc))

    generated_workbook = st.session_state.get("generated_workbook")
    generated_filename = st.session_state.get("generated_filename")
    if generated_workbook and generated_filename:
        st.success("✓ Updated workbook generated successfully.")
        st.download_button(
            "Download Updated Excel",
            data=generated_workbook,
            file_name=generated_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )


def main() -> None:
    st.set_page_config(
        page_title="Distributor Report Generator",
        page_icon="📊",
        layout="centered",
    )
    st.title("Distributor Report Generator")
    st.caption(
        "Upload today's distributor data, select a distributor, and generate "
        "an updated distributor workbook."
    )

    source_workbook = _render_source_workbook()
    if source_workbook is None:
        return

    source_selection = _render_distributor_selection(source_workbook)
    if source_selection is None:
        return

    destination_selection = _render_destination_selection()
    if destination_selection is None:
        return

    _render_review_and_generation(source_selection, destination_selection)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        LOGGER.exception("Unexpected application error")
        st.error(
            "Something went wrong while processing the workbook. "
            "Please check the uploaded files and try again."
        )
