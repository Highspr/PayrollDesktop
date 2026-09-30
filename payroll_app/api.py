"""Private, same-origin API served only inside the desktop application."""
from datetime import date
from contextlib import closing
from pathlib import Path
import calendar
import json
import re
import secrets
import shutil
import sqlite3
import tempfile
import logging

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field, ConfigDict, model_validator

from . import core
from .service import Store


class Employee(BaseModel):
    model_config = ConfigDict(extra='ignore')
    name: str = Field(min_length=1, max_length=200)
    dt_appt: str = ''
    bps: int | None = None
    basic: int = Field(default=0, ge=0)
    special: int = Field(default=0, ge=0)
    designation: str = ''
    section: str = ''
    account_no: str = ''
    bank: str = ''
    email: str = ''
    cpf_no: str = ''
    category: str = ''
    cpf_member: int = Field(default=1, ge=0, le=1)
    active: int = Field(default=1, ge=0, le=1)

    @model_validator(mode='before')
    @classmethod
    def legacy_nulls(cls, values):
        values = dict(values)
        for key in ('dt_appt','designation','section','account_no','bank','email','cpf_no','category'):
            if values.get(key) is None:
                values[key] = ''
        return values

    @model_validator(mode='after')
    def valid_name(self):
        self.name = self.name.strip()
        if not self.name:
            raise ValueError('Employee name is required.')
        return self


class PayLine(BaseModel):
    name: str = Field(min_length=1)
    dt: str = ''
    bps: int | None = None
    des: str = ''
    section: str = ''
    acct: str = ''
    bank: str = ''
    email: str = ''
    cpf_no: str = ''
    cpf_member: int = Field(default=1, ge=0, le=1)
    wd: int = Field(gt=0, le=31)
    worked: int = Field(ge=0, le=31)
    basic: int = Field(ge=0)
    hr: int = Field(ge=0)
    conv: int = Field(ge=0)
    medl: int = Field(ge=0)
    adhoc: int = Field(ge=0)
    spec: int = Field(ge=0)
    cpf: int = Field(ge=0)
    cp_loan: int = Field(ge=0)
    arrear: int = 0
    other: int = 0

    @model_validator(mode='after')
    def valid_days(self):
        if self.worked > self.wd:
            raise ValueError('Days worked cannot exceed working days.')
        if not self.name.strip():
            raise ValueError('Employee name is required.')
        return self


class Scale(BaseModel):
    bps: int = Field(ge=1, le=30)
    hr: int = Field(default=0, ge=0)
    conv: int = Field(default=0, ge=0)
    medl: int = Field(default=0, ge=0)
    adhoc: int = Field(default=0, ge=0)
    cpf: int = Field(default=0, ge=0)


def create_app(store: Store, secret: str, assets=None):
    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)

    @app.middleware('http')
    async def protect(request: Request, call_next):
        host = request.headers.get('host', '')
        origin = request.headers.get('origin')
        if host.split(':')[0] not in ('127.0.0.1', 'testserver') or (origin and origin != f'http://{host}'):
            return JSONResponse({'detail': 'Origin not allowed.'}, status_code=403)
        if request.url.path.startswith('/api/') and not secrets.compare_digest(request.headers.get('x-payroll-token', ''), secret):
            return JSONResponse({'detail': 'Desktop session required.'}, status_code=401)
        response = await call_next(request)
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['Cache-Control'] = 'no-store'
        return response

    @app.exception_handler(ValueError)
    async def value_error(request, error):
        return JSONResponse({'detail': str(error)}, status_code=400)

    @app.exception_handler(OSError)
    async def file_error(request, error):
        return JSONResponse({'detail': 'The file could not be accessed. Close it in other applications, check the selected folder and try again. ' + str(error)}, status_code=400)

    @app.exception_handler(sqlite3.DatabaseError)
    async def database_error(request, error):
        return JSONResponse({'detail': 'The database could not be read. Choose a valid Payroll Desk backup or restart the application. ' + str(error)}, status_code=400)

    @app.exception_handler(Exception)
    async def unexpected_error(request, error):
        logging.getLogger('payroll').error('Operation failed: %s (%s)', request.url.path, type(error).__name__)
        return JSONResponse({'detail': 'This operation could not be completed. Check the selected file or settings, then try again.'}, status_code=500)

    def month_key(month):
        if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])', month):
            raise ValueError('Choose a valid month.')
        return month

    def editable(db, month):
        month_key(month)
        if db.execute('SELECT 1 FROM locked WHERE month=?', (month,)).fetchone():
            raise HTTPException(409, 'This month is locked. Unlock it before editing.')

    @app.get('/api/health')
    def health():
        return {'ready': True}

    @app.get('/api/bootstrap')
    def bootstrap():
        with store.db() as db:
            settings = core.get_settings(db)
            settings['smtp_password'] = ''
            return {'settings': settings, 'labels': core.SETTING_LABELS,
                    'months': [r[0] for r in db.execute('SELECT month FROM monthly UNION SELECT month FROM locked ORDER BY month DESC')],
                    'data_directory': str(store.directory)}

    @app.get('/api/employees')
    def employees():
        with store.db() as db:
            return [dict(r) for r in db.execute('SELECT * FROM employees ORDER BY name COLLATE NOCASE')]

    @app.post('/api/employees')
    def add_employee(body: Employee):
        values = body.model_dump()
        with store.db() as db:
            cur = db.execute(f"INSERT INTO employees ({','.join(values)}) VALUES ({','.join('?' for _ in values)})", tuple(values.values()))
            db.commit()
            return {'id': cur.lastrowid}

    @app.put('/api/employees/{eid}')
    def update_employee(eid: int, body: Employee):
        values = body.model_dump()
        with store.db() as db:
            cur = db.execute(f"UPDATE employees SET {','.join(k+'=?' for k in values)} WHERE id=?", (*values.values(), eid))
            if not cur.rowcount:
                raise HTTPException(404, 'Employee not found.')
            db.commit()
        return {'ok': True}

    @app.get('/api/payroll/{month}')
    def payroll(month: str):
        with store.db() as db:
            rows, locked = core.month_rows(db, month_key(month))
            return {'rows': rows, 'locked': locked, 'totals': {k: sum(r[k] for r in rows) for k in ('gross', 'net', 'cpf')}}

    @app.put('/api/payroll/{month}/{eid}')
    def update_line(month: str, eid: int, body: PayLine):
        with store.db() as db:
            editable(db, month)
            if not db.execute('SELECT 1 FROM employees WHERE id=? AND active=1', (eid,)).fetchone():
                raise HTTPException(404, 'Active employee not found.')
            row = body.model_dump()
            row['emp_id'] = eid
            core.save_monthly_employee(db, month, row)
        return {'ok': True}

    @app.post('/api/payroll/{month}/lock')
    def lock(month: str):
        with store.db() as db:
            editable(db, month)
            rows, _ = core.month_rows(db, month)
            if not rows:
                raise ValueError('Add employees first.')
            for row in rows:
                PayLine.model_validate(dict(row, bps=row['bps'] or None))
            if any(r['net'] < 0 for r in rows):
                raise ValueError('Resolve negative net salaries before locking.')
            core.lock_month(db, month, rows)
        return {'ok': True}

    @app.post('/api/payroll/{month}/unlock')
    def unlock(month: str):
        with store.db() as db:
            core.unlock_month(db, month_key(month))
        return {'ok': True}

    @app.get('/api/scales')
    def scales():
        with store.db() as db:
            return list(core.get_scales(db).values())

    @app.post('/api/scales')
    def save_scale(body: Scale):
        with store.db() as db:
            db.execute('INSERT OR REPLACE INTO bps_scale VALUES (?,?,?,?,?,?)', tuple(body.model_dump().values()))
            db.commit()
        return {'ok': True}

    @app.delete('/api/scales/{bps}')
    def delete_scale(bps: int):
        with store.db() as db:
            db.execute('DELETE FROM bps_scale WHERE bps=?', (bps,))
            db.commit()
        return {'ok': True}

    @app.put('/api/settings')
    def settings(body: dict):
        values = {k: type(core.DEFAULTS[k])(v) for k, v in body.items() if k in core.DEFAULTS}
        if not 1 <= values.get('default_working_days', 30) <= 31:
            raise ValueError('Working days must be between 1 and 31.')
        if values.get('cpf_college_pct', 0) < 0 or not 1 <= values.get('smtp_port', 587) <= 65535:
            raise ValueError('Invalid CPF percentage or SMTP port.')
        if not values.get('smtp_password'):
            values.pop('smtp_password', None)
        with store.db() as db:
            core.save_settings(db, values)
        return {'ok': True}

    @app.post('/api/backup')
    def backup(body: dict):
        return {'path': store.backup(store.consume(body['ticket'], 'backup'))}

    @app.post('/api/restore')
    def restore(body: dict):
        source = Path(store.consume(body['ticket'], 'restore'))
        with store.guard, tempfile.TemporaryDirectory() as temp:
            staged = Path(temp) / 'restore.db'
            Store.copy_database(source, staged)
            with closing(sqlite3.connect(staged)) as test:
                if test.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
                    raise ValueError('The backup database failed integrity checks.')
                tables = {r[0] for r in test.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                if not {'employees', 'monthly', 'locked', 'settings', 'bps_scale'} <= tables:
                    raise ValueError('This is not a Payroll Desk backup.')
            migrated = core.connect(str(staged))
            try:
                core.month_rows(migrated, date.today().strftime('%Y-%m'))
            finally:
                migrated.close()
            store.backup()
            Store.copy_database(staged, store.path)
        return {'ok': True}

    @app.post('/api/logo')
    def logo(body: dict):
        from PIL import Image
        source = store.consume(body['ticket'], 'logo')
        with Image.open(source) as image:
            image.verify()
        shutil.copyfile(source, store.directory / 'logo.png')
        return {'ok': True}

    @app.post('/api/import')
    def import_excel(body: dict):
        source = store.consume(body['ticket'], 'import')
        with store.guard, tempfile.TemporaryDirectory() as temp:
            staged = Path(temp) / 'import.db'
            Store.copy_database(store.path, staged)
            db = core.connect(str(staged))
            try:
                result = core.import_workbook(db, source)
                month = f"{result['year']}-{result['month']:02}"
                editable(db, month)
            finally:
                db.close()
            store.backup()
            Store.copy_database(staged, store.path)
        return result

    @app.post('/api/export')
    def export(body: dict):
        kind = body.get('kind')
        if kind not in ('pdf', 'salary', 'bank', 'cpf', 'word'):
            raise ValueError('Unknown report type.')
        target = store.consume(body['ticket'], 'export')
        with store.db() as db:
            month = month_key(body['month'])
            rows, locked = core.month_rows(db, month)
            settings = core.get_settings(db)
        ids = body.get('ids', [])
        if ids:
            rows = [r for r in rows if r['emp_id'] in ids]
        if not rows:
            raise ValueError('No employees selected for export.')
        year, number = map(int, month.split('-'))
        title = f"{calendar.month_name[number]} {year}" + ('' if locked else ' — DRAFT')
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / ('report' + Path(target).suffix)
            if kind == 'pdf':
                core.make_payslips(str(output), rows, settings, title)
            elif kind == 'salary':
                core.export_xlsx(str(output), settings['college_name'] + ' — ' + title, rows)
            elif kind == 'bank':
                core.export_bank_xlsx(str(output), 'Salary Transfer — ' + title, rows)
            elif kind == 'cpf':
                if not locked:
                    settings['college_name'] += ' — DRAFT'
                core.export_cpf_xlsx(str(output), settings, rows, date.today())
            else:
                core.export_salary_word(str(output), 'Bank Details — ' + title, rows)
            shutil.copyfile(output, target)
        return {'path': target, 'open_ticket': store.ticket(target, 'open')}

    @app.post('/api/email')
    def email(body: dict):
        with store.db() as db:
            rows, locked = core.month_rows(db, month_key(body['month']))
            settings = core.get_settings(db)
        if not locked:
            raise ValueError('Lock this month before emailing payslips.')
        with tempfile.TemporaryDirectory() as temp:
            sent, skipped = core.send_payslips_email(rows, settings, body['month'], temp)
        return {'sent': sent, 'skipped': skipped}

    if assets:
        app.mount('/', StaticFiles(directory=str(assets), html=True), name='frontend')
    return app
