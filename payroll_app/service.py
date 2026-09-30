"""Local data ownership, migration and native-dialog file capabilities."""
from contextlib import contextmanager, closing
from datetime import datetime
from pathlib import Path
import os
import secrets
import sqlite3
import threading

from . import core


class Store:
    def __init__(self, directory=None, legacy=None):
        self.directory = Path(directory or Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'PayrollDesk')
        self.directory.mkdir(parents=True, exist_ok=True)
        self.path = self.directory / 'payroll.db'
        self.guard = threading.RLock()
        self.files = {}
        if not self.path.exists() and legacy and Path(legacy).exists():
            self.copy_database(Path(legacy), self.path)
            self.backup()
        db = core.connect(str(self.path))
        db.close()
        core.APP_DIR = str(self.directory)

    @staticmethod
    def copy_database(source, target):
        with closing(sqlite3.connect(Path(source).resolve().as_uri() + '?mode=ro', uri=True)) as src:
            with closing(sqlite3.connect(str(target))) as dst:
                src.backup(dst)

    @contextmanager
    def db(self):
        with self.guard:
            db = sqlite3.connect(self.path)
            db.row_factory = sqlite3.Row
            try:
                yield db
            finally:
                db.close()

    def backup(self, target=None):
        with self.guard:
            folder = self.directory / 'backups'
            folder.mkdir(exist_ok=True)
            target = Path(target or folder / (datetime.now().strftime('%Y%m%d-%H%M%S-%f') + '.db'))
            if target.resolve() == self.path.resolve():
                raise ValueError('Choose a different file from the live database.')
            self.copy_database(self.path, target)
            return str(target)

    def ticket(self, path, purpose):
        key = secrets.token_urlsafe(24)
        self.files[key] = (str(path), purpose)
        return key

    def consume(self, key, purpose):
        item = self.files.pop(key, None)
        if not item or item[1] != purpose:
            raise ValueError('Choose the file again using the desktop dialog.')
        if Path(item[0]).resolve() == self.path.resolve():
            raise ValueError('The live database cannot be overwritten or imported directly.')
        return item[0]
