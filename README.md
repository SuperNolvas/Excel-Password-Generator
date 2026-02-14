# Password Generator

Simple Python script that picks two words from `dictionary.txt`, builds a password (contains uppercase, lowercase, symbol and number, 10-32 chars), shows it in the console for acceptance, writes the accepted password into the first cell of an Excel workbook in 36pt font and attempts to password-protect the workbook using Excel COM on Windows.

## Usage

1. Put your `dictionary.txt` (one word per line) next to the script.
2. Install dependencies:

```bash
python -m pip install -r requirements.txt
```

3. Run the script (PowerShell/Cmd):

```powershell
python password_generator.py
```

## Notes
- Excel password protection uses the Windows COM interface (`pywin32`) and requires Microsoft Excel to be installed. If COM protection fails or `pywin32` is not available, the script will still save the workbook unprotected.
- The script writes the password to cell A1 in 36pt font and adjusts the column width and row height approximately to fit.

## Run in a virtual environment

It's recommended to run the script inside a Python virtual environment to keep dependencies isolated. Example commands:

Windows (PowerShell):
```powershell
python -m venv .venv
& .\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python password_generator.py
```

Windows (Cmd.exe):
```cmd
python -m venv .venv
.\.venv\Scripts\activate.bat
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

Notes:
- If you want Excel encryption on Windows, ensure Microsoft Excel is installed and `pywin32` is available in the venv.
- The repository `.gitignore` already excludes `.xlsx` files so generated workbooks won't be committed.
