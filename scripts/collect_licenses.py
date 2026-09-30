"""Copy dependency notices alongside the distributable (no build-time network)."""
from importlib import metadata
from pathlib import Path
import shutil
import sys

root = Path(__file__).resolve().parents[1]
destination = root / 'dist' / 'PayrollDesk' / 'licenses'
destination.mkdir(parents=True, exist_ok=True)
for source in (root / 'licenses').glob('*'):
    shutil.copy2(source, destination / source.name)
excluded = {'pymupdf','pytest','httpx','httpcore','pygments','iniconfig','pluggy','pyinstaller','pyinstaller-hooks-contrib','pefile','altgraph'}
for distribution in metadata.distributions():
    name = distribution.metadata['Name']
    if name.lower().replace('_','-') in excluded:
        continue
    for item in distribution.files or []:
        basename = Path(str(item)).name.lower()
        if ('.dist-info/' in str(item) and ('license' in basename or 'copying' in basename or 'notice' in basename)):
            folder = destination / 'python' / name
            folder.mkdir(parents=True, exist_ok=True)
            shutil.copy2(distribution.locate_file(item), folder / Path(str(item)).name)
    if name.lower().startswith(('pyside','shiboken')):
        folder = destination / 'python' / name
        folder.mkdir(parents=True, exist_ok=True)
        (folder / 'PACKAGE-METADATA.txt').write_text(distribution.read_text('METADATA') or '',encoding='utf-8')
for package in (root / 'frontend' / 'node_modules').iterdir():
    packages = list(package.iterdir()) if package.name.startswith('@') else [package]
    for entry in packages:
        if entry.is_dir():
            for source in entry.iterdir():
                if source.is_file() and source.name.lower().startswith(('license','copying','notice')):
                    folder = destination / 'javascript' / entry.relative_to(root / 'frontend' / 'node_modules')
                    folder.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, folder / source.name)
for name in ('LICENSE.txt','LICENSE'):
    source=Path(sys.base_prefix)/name
    if source.exists():
        shutil.copy2(source,destination/'Python-LICENSE.txt')
print('Dependency notices copied to', destination)
