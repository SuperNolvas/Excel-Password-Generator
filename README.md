# Password Generator

Small utility that generates a readable but strong password from two words, shows it for acceptance, writes the accepted password into an Excel workbook (cell A1, 36pt), and optionally password-protects the workbook using Excel COM on Windows.

## Key features
- Generates a 10–32 character password with upper/lowercase letters, digit(s) and 1–2 symbols.
- Interactive acceptance prompt before saving.
- Writes the password to an `.xlsx` workbook (first cell) and adjusts formatting for readability.
- Optional Excel workbook protection via Windows COM (`pywin32`) when Microsoft Excel is installed.
- Attach the generated `.xlsx` to an existing `.eml` template from `EXTERNAL EMAIL TEMPLATES/` and save a ready-to-send `.eml`.
  - The script lists available templates, lets you choose one, prompts for a recipient email (optional), attaches the workbook and writes a `READY_TO_SEND_{templateName}_{workbookName}.eml` output.

## Quick start

1. Place `dictionary.txt` (one word per line) in the repository root.
2. Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

3. Run the script:

```powershell
python password_generator.py
```

Follow the prompts to accept a password, choose an output filename, optionally protect the file, and optionally attach it to an `.eml` template.

## Email template attachment
- Put `.eml` templates inside `EXTERNAL EMAIL TEMPLATES/`.
- After creating the workbook the script can:
  - List templates and let you choose one.
  - Prompt for a recipient email address (pasteable in the CLI). If left blank, the template's `To` header is kept.
  - Attach the generated `.xlsx` to the chosen template and save `READY_TO_SEND_{templateName}_{workbookName}.eml` in the repo root.

Notes:
- The script uses Python's standard `email` library; no extra packages are required for this feature.
- If a template already contains attachments, the workbook will be appended as an additional attachment.

## Git / ignored files
- The repository `.gitignore` excludes generated files:
  - `*.xlsx` and general `.eml` files are ignored by default.
  - Runtime folders such as `READY EMAILS/` and `TEMP EXCEL PASSWORDS/` are ignored, but each contains a `.gitkeep` placeholder so the empty folder is tracked.

If you want to track output `.eml` files like `READY_TO_SEND_...eml`, remove or adjust the `*.eml` rule in `.gitignore`.

## Notes on Excel protection
- Excel protection relies on the Windows COM interface via `pywin32` and requires Excel to be installed. If COM fails or `pywin32` is unavailable the workbook is still written unprotected and the script prints a warning.

## Running in a virtual environment (recommended)

Windows (PowerShell):
```powershell
python -m venv .venv
& .\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python password_generator.py
```

macOS / Linux:
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python password_generator.py
```

## Troubleshooting
- If `openpyxl` is missing, install via `pip install -r requirements.txt`.
- If Excel protection fails, ensure Excel is installed and `pywin32` is available in the environment.
- If templates are not found, confirm `.eml` files are present in `EXTERNAL EMAIL TEMPLATES/` (note `.gitignore` may hide `.eml` files from commits but they still exist locally).

---

If you'd like, I can commit these docs changes and push them for you.
