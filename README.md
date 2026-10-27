# Doc.Bot — VS Code + Python + Excel local edition

This folder is a ready-to-run local version of Doc.Bot. It serves the standalone web app from Python and stores user-entered medical reports, imported spreadsheet rows, and Doc.Bot query/care-plan records in:

```text
data/docbot_users.xlsx
```

OpenRouter remains the primary optional AI path. Settings now loads the current free OpenRouter models (with a built-in fallback list), and you choose one concrete model directly; there is no automatic model-selection route.

The workbook is created automatically on first start. It has these sheets:

- **Users** — one opaque browser-local user ID and timestamps
- **Reports** — saved medical-condition/report entries
- **ImportedSheets** — rows imported from Excel or CSV
- **QueryLog** — Doc.Bot intake questions, urgency, care-plan text, and replies

## Run it in VS Code

1. Open this **`DocBot_Localhost_Python`** folder in VS Code.
2. Open the integrated terminal in this folder.
3. Optional but recommended: create a virtual environment.

### Windows PowerShell

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python server.py
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python server.py
```

4. Open **http://127.0.0.1:8000** in your browser.
5. Leave the terminal running while using Doc.Bot.

There are also shortcuts:

- Windows: double-click `run.bat`
- macOS/Linux: run `./run.sh`

If port 8000 is busy:

```bash
python server.py --port 8001
```

Then open `http://127.0.0.1:8001`.

## Excel storage behavior

The browser sends only user-entered reports, imported health-sheet data, and saved Doc.Bot records to the local Python API. The API writes those rows to `data/docbot_users.xlsx`. The OpenRouter key is **not** sent to Python and is not written to Excel; it remains in the browser's local settings.

The app still works when opened as a standalone HTML file or when the Python server is unavailable. In that case it uses browser storage and shows the Excel-sync status as offline. When it is opened through the Python server, the sidebar status changes to **Python server connected · Excel sync enabled** and saved data is restored from the workbook on the next session.

You can download the current workbook at:

```text
http://127.0.0.1:8000/api/download.xlsx
```

You can also open `data/docbot_users.xlsx` directly in Excel or LibreOffice. Close the workbook before editing records in Doc.Bot, especially on Windows, because an open Excel file may lock writes.

## Privacy and security note

This is a local development/storage setup. It has no login, encryption, or multi-user authentication. Anyone who can access the computer or the bound server can access the workbook. Keep the server bound to `127.0.0.1` and do not expose it to the public internet. Treat the workbook as sensitive medical data and protect or delete it when no longer needed.

Doc.Bot provides general health information only; it does not diagnose or replace a qualified clinician.
