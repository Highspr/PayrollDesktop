# Third-party components

Payroll Desk bundles Python, PySide6/Qt (including Qt WebEngine/Chromium), FastAPI, Starlette, Uvicorn, Pydantic, SQLite, ReportLab, openpyxl, python-docx, Next.js, React, Tailwind CSS, Lucide icons, and their dependencies.

These components retain their respective licenses and copyright notices. Qt/PySide6 libraries are dynamically linked in the distributed directory and must remain replaceable under the applicable LGPL terms. Qt WebEngine includes Chromium third-party notices; its `qtwebengine_resources` files and licensing information must accompany distribution. Review the bundled licenses and upstream source notices before redistribution.

Official component sources:

- https://www.python.org/
- https://code.qt.io/pyside/pyside-setup.git/
- https://code.qt.io/qt/qtwebengine.git/
- https://github.com/fastapi/fastapi
- https://github.com/encode/uvicorn
- https://github.com/vercel/next.js
- https://github.com/facebook/react
- https://github.com/tailwindlabs/tailwindcss
- https://github.com/lucide-icons/lucide

The Windows installer is produced with Inno Setup. The desktop executable is packaged with PyInstaller, whose bootloader exception allows distribution of the application.
