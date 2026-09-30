# Payroll Desk

Offline Windows desktop payroll: **Next.js + Tailwind CSS**, **PySide6 WebEngine**, **FastAPI**, and **SQLite**. The Next.js frontend is statically exported; there is no Node.js server in the installed application.

## Run

Double-click **Run Payroll.cmd** in this workspace. To distribute the application, use `dist/PayrollDesk-Setup.exe` after running the build. The installer bundles Python, Qt WebEngine, libraries, and all frontend assets. End users do not need Python, Node.js, internet, or terminal commands. Supported target: Windows 10 (1809+) / Windows 11 x64 with working graphics drivers.

## Data and migration

Data is stored in `%LOCALAPPDATA%\PayrollDesk\payroll.db`. On first source launch, the app copies the adjacent legacy `payroll.db` using SQLite's backup API and creates a backup before schema upgrades. It never edits the original database. Installed copies start empty unless a database is already present; use **Settings > Backup & restore** to bring records to another PC. The installer deliberately does not include this college's private database or salary workbook.

Only one app instance runs per Windows user/data directory. The backend binds an automatically chosen `127.0.0.1` port, checks origin/host and a per-launch secret, and shuts down with the window. Desktop file dialogs issue one-use capabilities so web requests cannot choose arbitrary local file paths.

The original Tkinter application remains in `payroll.py`. Do not continue entering payroll in both applications after migrating: their databases are independent. Use `Run Legacy Payroll.cmd` only if you intentionally need the original application.

## Everyday workflow

1. Add/import employees; confirm BPS scales and college settings.
2. Choose a month. Click an employee to adjust that month's salary, days or bank details.
3. Save, review totals, then **Review & lock month**. Locked rows reject changes at the API as well as the UI.
4. Download all or selected reports using **Save PDF / Export Excel** or the Reports screen. Search filtering does not silently change export scope.
5. Back up regularly to a separate drive. Restore and workbook import create a safety backup first.

Master edits affect unlocked months wherever no monthly override exists. Unlocking deletes the final snapshot and recalculates from current masters/overrides; there is no revision/audit subsystem. Deactivation excludes a person from unlocked payrolls; locked historical payrolls stay unchanged. Email sends only from a locked month and always requires user confirmation. Do not retry an ambiguous email failure without checking delivery.

Amounts retain the original whole-rupee calculation rules. Adhoc is a fixed amount, not an automatic percentage. CPF college contribution is additional to the employee deduction. Tax, attendance integration, bank payment execution, and loan-balance tracking are outside this release.

Database backups contain settings (including SMTP credentials) and must be kept private. PDFs, Excel/Word exports and `logo.png` are separate files. Set the logo again after moving to a different computer. Settings hides the saved SMTP password; blank preserves it. This local release has no login or data encryption.

## Development

Build machine only: Python 3.12, Node.js 22 and, for the installer, Inno Setup 6.

```powershell
uv venv --python 3.12 .desktop-venv
uv pip install --python .desktop-venv\Scripts\python.exe -r requirements-dev.txt
cd frontend
npm ci
npm run build
cd ..
.desktop-venv\Scripts\python.exe desktop.py
```

`powershell -NoProfile -ExecutionPolicy Bypass -File scripts/build.ps1 -Installer` builds the static frontend, runs tests, packages the desktop application with PyInstaller, collects dependency notices, and compiles the installer. The execution-policy option applies only to that build process. `frontend/package-lock.json` pins JavaScript dependencies; requirements files pin direct Python dependencies. No college database is included in the build.

## Verification

```powershell
.desktop-venv\Scripts\python.exe -m pytest -q
$env:PAYROLL_DATA_DIR = "$PWD\artifacts\smoke-data"
.desktop-venv\Scripts\python.exe desktop.py --smoke artifacts\desktop.png
```

Tests cover formula behavior, monthly isolation, locking, host/origin/token protection, PDF/Excel/Word generation, backup/restore, staged import rejection for locked periods, migration parity, and 400 employees. The smoke command captures the actual Qt window and rendered text, then exits. It uses an isolated data directory; omit `PAYROLL_DATA_DIR` for normal use.

An independently clean Windows machine/VM without Python or Node.js is still needed for final portability acceptance. Building and running the bundled executable on the development PC is not equivalent to that test. Email delivery requires actual SMTP credentials and internet; automated tests must not send real employee messages.
