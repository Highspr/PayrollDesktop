# Build with: .desktop-venv\Scripts\python.exe -m PyInstaller PayrollDesk.spec
from pathlib import Path

root = Path(SPECPATH)
datas = [(str(root / 'frontend' / 'out'), 'frontend/out'), (str(root / 'README.md'), '.'), (str(root / 'THIRD_PARTY_NOTICES.md'), '.')]
a = Analysis(['desktop.py'], pathex=[str(root)], binaries=[], datas=datas,
             hiddenimports=['uvicorn.logging','uvicorn.loops.auto','uvicorn.protocols.http.h11_impl','uvicorn.lifespan.on'],
             hookspath=[], runtime_hooks=[], excludes=['tkinter','pytest','httpx','PySide6.QtQml','PySide6.QtQuick','PySide6.Qt3DCore'], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='PayrollDesk', debug=False, bootloader_ignore_signals=False, strip=False, upx=False, console=False)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name='PayrollDesk')
