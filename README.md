# Distributor Report Generator

A small Streamlit utility that prepares distributor-specific Excel workbooks
entirely in memory. It provides the complete upload, selection, preview,
generation, and download workflow without sending corporate files to any
external service.

## What the app does

1. Upload today's `.xlsx` workbook and choose its source sheet.
2. Choose the column that contains distributor names.
3. Select a distributor extracted directly from the uploaded data.
4. Preview the matching source rows.
5. Upload the distributor's existing `.xlsx` workbook.
6. Choose an actual destination sheet from that workbook.
7. Review the selections, generate an updated workbook, and download it.

The generated workbook preserves its original sheets, data, formulas, and
basic workbook formatting. For now, it adds or replaces an
`Automation Preview` sheet containing the selected distributor's source rows.
This proves that the complete pipeline works before real mapping rules are
introduced.

## Requirements

- Python 3.12 or newer
- Local access to the Excel workbooks being processed

## Setup

macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-dev.txt
streamlit run app.py
```

Windows PowerShell:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
streamlit run app.py
```

Windows Command Prompt:

```bat
py -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements-dev.txt
streamlit run app.py
```

Streamlit prints a local URL after startup, usually
`http://localhost:8501`.

## Tests

With the virtual environment active, run:

```bash
python -m pytest
```

The tests cover data cleaning and filtering, input validation, filename
sanitization, month-based sheet suggestions, Excel reading, and preservation
of the existing destination workbook during placeholder generation.

## Deploy on Streamlit Community Cloud

This repository is organized for Streamlit Community Cloud deployment. The
production dependencies are pinned in `requirements.txt`, the entrypoint is
`app.py`, and headless server settings live in `.streamlit/config.toml`.

After pushing the repository to GitHub:

1. Sign in to [Streamlit Community Cloud](https://share.streamlit.io/) with
   GitHub and click **Create app**.
2. Select repository `impranavgarg/WF1-EXE`.
3. Select branch `main` and entrypoint file `app.py`.
4. In **Advanced settings**, keep Python `3.12` selected.
5. Leave the secrets field empty; this application does not require secrets.
6. Choose whether the app should be public or private, then deploy it.

Important: a cloud deployment processes uploaded workbooks on Streamlit's
servers. Because this app has no application-level authentication, confirm the
repository and app visibility are appropriate before uploading corporate
files. For strictly local-only processing, continue to run the app locally.

## Current Status

The application currently implements the complete file-selection and
workbook-generation workflow. Distributor-specific mapping and calculation
rules are intentionally left as a placeholder until real sample workbooks are
available.

## Future Mapping Logic

The future distributor-specific calculations and destination-cell mappings
belong in [`src/transformer.py`](src/transformer.py). The Streamlit interface
already passes the filtered source data, selected distributor, destination
workbook bytes, and destination sheet into that service, so the UI does not
need to be rewritten when real sample workbooks and mapping rules arrive.

## Data handling

- Uploaded workbooks are read and generated in memory.
- The app does not use a database, authentication, AI, external APIs, or
  third-party file services.
- Uploaded workbooks are not intentionally written to disk.
