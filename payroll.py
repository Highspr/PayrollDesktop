"""College Payroll - standalone Windows desktop app (works offline).

Run:    python payroll.py
Needs:  pip install openpyxl reportlab python-docx
Build:  pip install pyinstaller
        pyinstaller --onefile --noconsole payroll.py
Data is kept in payroll.db, next to this script (or next to the .exe).
Back up that one file to back up everything.

Allowances and the CP Fund are fixed amounts for each BPS grade (BPS Scales
tab). Staff with no BPS (contract) get their running basic only. Special
allowance is entered per person. Use Employees > Import to load a monthly
salary workbook in the college's own layout.
Optional: put your college crest as logo.png next to this file and it will
appear at the top of every payslip.
"""
import calendar
import json
import os
import re
import sqlite3
import sys
import smtplib
from email.message import EmailMessage
from datetime import date

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
except ImportError:  # lets the calculation code be tested without a display
    tk = None

APP_DIR = os.path.dirname(
    sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
)
DB_FILE = os.path.join(APP_DIR, "payroll.db")

# Placeholder values - set your own in the Settings tab.
DEFAULTS = {
    "college_name": "College Name",
    "default_working_days": 30,
    "adhoc_label": "10% of Basic Pay",
    "cpf_college_pct": 100.0,
    "default_bank": "ABL",
    "voucher_prefix": "BCZC",
    "slip_header": "COLLEGE NAME, CAMPUS ADDRESS  Tel: 000-0000000",
    "signatory_name": "",
    "signatory_title": "",
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_username": "",
    "smtp_password": "",
    "smtp_from": "",
}
SETTING_LABELS = {
    "college_name": "College name (shown on the salary sheet)",
    "default_working_days": "Working days in a month (default)",
    "adhoc_label": "Payslip wording for the adhoc / increase line",
    "cpf_college_pct": "CPF: college contribution, % of employee's deduction",
    "default_bank": "Bank name (used when an employee has none entered)",
    "voucher_prefix": "CPF statement voucher prefix (e.g. BCZC)",
    "slip_header": "Payslip heading (college, campus, phone)",
    "signatory_name": "Payslip signatory name",
    "signatory_title": "Payslip signatory designation",
    "smtp_server": "Email SMTP server (e.g. smtp.gmail.com)",
    "smtp_port": "Email SMTP port (usually 587)",
    "smtp_username": "Email username / login",
    "smtp_password": "Email password / app password",
    "smtp_from": "From email address (leave blank to use username)",
}

HEADS = ["S.NO", "Dt of Appt", "Name of Incumbent", "BPS", "Wrkg Days",
         "With Work", "Encash/ Arrear", "R.Basic Pay", "H.Rent Allow",
         "Conv Allow", "Medl Allow", "Adhoc I/Basic", "Special Allow",
         "Gross Salary", "CP Fund", "CP Loan Instalmt", "L.W.O Pay/Any",
         "Net Salary"]
KEYS = ["sno", "dt", "name", "bps", "wd", "worked", "arrear", "basic", "hr",
        "conv", "medl", "adhoc", "spec", "gross", "cpf", "cp_loan", "lwop",
        "net"]


# ----------------------------------------------------------------- database
def connect(path=DB_FILE):
    db = sqlite3.connect(path)
    db.row_factory = sqlite3.Row
    db.executescript("""
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE IF NOT EXISTS employees(
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
        dt_appt TEXT, bps INTEGER, basic INTEGER NOT NULL DEFAULT 0,
        active INTEGER NOT NULL DEFAULT 1,
        designation TEXT, section TEXT, account_no TEXT,
        bank TEXT, email TEXT, cpf_no TEXT, special INTEGER NOT NULL DEFAULT 0,
        category TEXT, cpf_member INTEGER NOT NULL DEFAULT 1);
    CREATE TABLE IF NOT EXISTS bps_scale(
        bps INTEGER PRIMARY KEY, hr INTEGER NOT NULL DEFAULT 0,
        conv INTEGER NOT NULL DEFAULT 0, medl INTEGER NOT NULL DEFAULT 0,
        adhoc INTEGER NOT NULL DEFAULT 0, cpf INTEGER NOT NULL DEFAULT 0);
    CREATE TABLE IF NOT EXISTS monthly(
        month TEXT, emp_id INTEGER, wd INTEGER, worked INTEGER,
        arrear INTEGER DEFAULT 0, cp_loan INTEGER DEFAULT 0,
        other INTEGER DEFAULT 0, basic INTEGER, hr INTEGER, conv INTEGER,
        medl INTEGER, adhoc INTEGER, special INTEGER, cpf INTEGER,
        name TEXT, dt_appt TEXT, bps INTEGER, designation TEXT, section TEXT,
        account_no TEXT, bank TEXT, email TEXT, cpf_no TEXT, cpf_member INTEGER,
        PRIMARY KEY(month, emp_id));
    CREATE TABLE IF NOT EXISTS locked(
        month TEXT, emp_id INTEGER, data TEXT, PRIMARY KEY(month, emp_id));
    """)
    have = {r["name"] for r in db.execute("PRAGMA table_info(employees)")}
    for col, decl in (("designation", "TEXT"), ("section", "TEXT"),
                      ("account_no", "TEXT"), ("bank", "TEXT"), ("email", "TEXT"), ("cpf_no", "TEXT"),
                      ("special", "INTEGER NOT NULL DEFAULT 0"),
                      ("category", "TEXT"),
                      ("cpf_member", "INTEGER NOT NULL DEFAULT 1")):
        if col not in have:
            db.execute(f"ALTER TABLE employees ADD COLUMN {col} {decl}")
    # Monthly overrides are nullable: NULL means "use the employee/BPS master value".
    monthly_cols = {r["name"] for r in db.execute("PRAGMA table_info(monthly)")}
    for col, decl in (("basic", "INTEGER"), ("hr", "INTEGER"), ("conv", "INTEGER"),
                      ("medl", "INTEGER"), ("adhoc", "INTEGER"), ("special", "INTEGER"),
                      ("cpf", "INTEGER"), ("name", "TEXT"), ("dt_appt", "TEXT"),
                      ("bps", "INTEGER"), ("designation", "TEXT"), ("section", "TEXT"),
                      ("account_no", "TEXT"), ("bank", "TEXT"), ("email", "TEXT"),
                      ("cpf_no", "TEXT"), ("cpf_member", "INTEGER")):
        if col not in monthly_cols:
            db.execute(f"ALTER TABLE monthly ADD COLUMN {col} {decl}")
    db.commit()
    return db


def get_settings(db):
    s = dict(DEFAULTS)
    for r in db.execute("SELECT key, value FROM settings"):
        if r["key"] in DEFAULTS:
            try:
                s[r["key"]] = type(DEFAULTS[r["key"]])(r["value"])
            except ValueError:
                pass
    return s


def save_settings(db, s):
    for k, v in s.items():
        db.execute("INSERT OR REPLACE INTO settings VALUES (?, ?)", (k, str(v)))
    db.commit()


# -------------------------------------------------------------- calculation
def get_scales(db):
    return {r["bps"]: dict(r) for r in db.execute("SELECT * FROM bps_scale")}


def calculate(basic, scale, special, member, s, wd, worked,
              arrear=0, cp_loan=0, other=0):
    """Pay for one employee.

    scale  - the BPS grade's fixed amounts (hr, conv, medl, adhoc, cpf), or
             None for staff with no BPS (contract): running basic only.
    member - True if the person contributes to the CP Fund.
    """
    def rs(x):
        return int(x + 0.5)

    sc = scale or {}
    hr, conv, medl, adhoc = (sc.get(k, 0) for k in ("hr", "conv", "medl", "adhoc"))
    spec = special or 0
    allowances = hr + conv + medl + adhoc + spec
    gross = arrear + basic + allowances
    cpf = sc.get("cpf", 0) if member else 0
    unpaid_days = max(0, wd - worked)
    lwop_pay = rs((basic + allowances) / wd * unpaid_days) if wd > 0 else 0
    lwop = lwop_pay + other
    net = gross - cpf - cp_loan - lwop
    return dict(basic=basic, hr=hr, conv=conv, medl=medl, adhoc=adhoc,
                spec=spec, gross=gross, cpf=cpf, cp_loan=cp_loan,
                lwop=lwop, lwop_pay=lwop_pay, net=net,
                cpf_college=rs(cpf * s["cpf_college_pct"] / 100))


def month_rows(db, month):
    """Return the editable current-month view and whether it is locked."""
    if db.execute("SELECT 1 FROM locked WHERE month=?", (month,)).fetchone():
        rows = [json.loads(r["data"]) for r in db.execute(
            "SELECT data FROM locked WHERE month=? ORDER BY emp_id", (month,))]
        for i, r in enumerate(rows, 1):
            r["sno"] = i
        return rows, True

    s = get_settings(db)
    scales = get_scales(db)
    rows = []
    emps = db.execute("SELECT * FROM employees WHERE active=1 ORDER BY id")
    for i, e in enumerate(emps.fetchall(), 1):
        m = db.execute("SELECT * FROM monthly WHERE month=? AND emp_id=?",
                       (month, e["id"])).fetchone()
        def mv(col, master):
            return (m[col] if m is not None and m[col] is not None else master)
        name = mv("name", e["name"])
        dt = mv("dt_appt", e["dt_appt"] or "")
        bps = mv("bps", e["bps"])
        basic = int(mv("basic", e["basic"]) or 0)
        special = int(mv("special", e["special"] or 0) or 0)
        member = int(mv("cpf_member", e["cpf_member"]) if mv("cpf_member", e["cpf_member"]) is not None else 1)
        scale_master = scales.get(bps) if bps is not None else None
        custom_scale = dict(scale_master or {})
        for k in ("hr", "conv", "medl", "adhoc", "cpf"):
            val = mv(k, None)
            if val is not None:
                custom_scale[k] = int(val or 0)
        wd = int(mv("wd", s["default_working_days"]) or 0)
        worked = int(mv("worked", wd) or 0)
        arrear = int(mv("arrear", 0) or 0)
        cp_loan = int(mv("cp_loan", 0) or 0)
        other = int(mv("other", 0) or 0)
        row = calculate(basic, custom_scale if custom_scale else None, special, member, s,
                        wd, worked, arrear, cp_loan, other)
        row.update(emp_id=e["id"], sno=i, name=name, dt=dt, bps=bps or "", wd=wd,
                   worked=worked, arrear=arrear, other=other,
                   des=mv("designation", e["designation"] or "") or "",
                   section=mv("section", e["section"] or "") or "",
                   acct=mv("account_no", e["account_no"] or "") or "",
                   bank=mv("bank", e["bank"] or s["default_bank"]) or s["default_bank"],
                   email=mv("email", e["email"] or "") or "",
                   cpf_no=mv("cpf_no", e["cpf_no"] or "") or "",
                   cpf_member=member, adhoc_label=s["adhoc_label"])
        rows.append(row)
    return rows, False


def save_monthly_employee(db, month, row):
    """Save every editable employee/pay field for this month."""
    db.execute("""INSERT OR REPLACE INTO monthly
        (month, emp_id, wd, worked, arrear, cp_loan, other, basic, hr, conv, medl,
         adhoc, special, cpf, name, dt_appt, bps, designation, section, account_no,
         bank, email, cpf_no, cpf_member)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (month, row["emp_id"], row["wd"], row["worked"], row["arrear"], row["cp_loan"],
         row["other"], row["basic"], row["hr"], row["conv"], row["medl"], row["adhoc"],
         row["spec"], row["cpf"], row["name"], row["dt"], row["bps"] or None,
         row.get("des", ""), row.get("section", ""), row.get("acct", ""),
         row.get("bank", ""), row.get("email", ""), row.get("cpf_no", ""),
         row.get("cpf_member", 1)))
    db.commit()


def lock_month(db, month, rows):
    for r in rows:
        db.execute("INSERT OR REPLACE INTO locked VALUES (?, ?, ?)",
                   (month, r["emp_id"], json.dumps(r)))
    db.commit()


def unlock_month(db, month):
    db.execute("DELETE FROM locked WHERE month=?", (month,))
    db.commit()


# --------------------------------------------------------- import from Excel
BANK_ABBR = {"allied bank": "ABL", "bank al habib": "BAH"}
TITLES = {"mr", "mrs", "ms", "miss", "mis", "dr", "prof"}
MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}


def _num(v):
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).strip().replace(",", ""))
    except ValueError:
        return 0


def _norm(name):
    n = " ".join(str(name).lower().replace(".", " ").split())
    return "".join(ch for ch in n if ch.isalpha() or ch == " ")


def _given(name):
    words = [w for w in _norm(name).split() if w not in TITLES]
    return words[0] if words else ""


def _similar(a, b):
    wa = {w for w in _norm(a).split() if w not in TITLES}
    wb = {w for w in _norm(b).split() if w not in TITLES}
    return bool(wa and wb) and (_given(a) == _given(b) or wa <= wb or wb <= wa)


def _text_date(v):
    if hasattr(v, "strftime"):
        return v.strftime("%d.%m.%Y")
    return str(v).strip() if v else ""


def _text_acct(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def import_workbook(db, path):
    """Load staff, BPS scale, bank details and one month's inputs from the
    college's monthly salary workbook (salary sheet + bank advice sheets).
    Safe to run again: people are matched by name + appointment date."""
    from collections import Counter
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True)
    sal_ws, banks = None, []
    for ws in wb.worksheets:
        if "Calculation of Staff Salary" in str(ws["A1"].value or ""):
            sal_ws = ws
            continue
        top = [str(c.value) for row in ws.iter_rows(min_row=1, max_row=14)
               for c in row if c.value]
        if any("BANK ADVICE" in t.upper() for t in top):
            name = next((" ".join(t.split()).replace(" Limited", "").replace(" Ltd", "")
                         for t in top if "Bank" in t and ("Limited" in t or "Ltd" in t)), "")
            banks.append((ws, BANK_ABBR.get(name.lower(), name)))
    if sal_ws is None:
        raise ValueError("Could not find the salary sheet (its title should "
                         "contain 'Calculation of Staff Salary').")
    title = str(sal_ws["A1"].value)
    m = re.search(r"Month of\s+([A-Za-z]+)\s+(\d{4})", title)
    if not m or m.group(1).lower() not in MONTHS:
        raise ValueError("Could not read the month from the salary sheet title.")
    month_no, year = MONTHS[m.group(1).lower()], int(m.group(2))
    mkey = f"{year}-{month_no:02d}"

    # ---- salary rows
    rows, cat = [], "Teaching"
    for r in sal_ws.iter_rows(values_only=True):
        r = list(r) + [None] * 20
        head = " ".join(str(v) for v in r[:4] if isinstance(v, str))
        if "Total Pay of Teaching" in head:
            cat = "Admin"
            continue
        if not (isinstance(r[0], (int, float)) and isinstance(r[2], str)
                and r[2].strip() and r[7] is not None):
            continue
        bps = int(_num(r[3])) if str(r[3]).strip().isdigit() or isinstance(r[3], (int, float)) else None
        wd = int(_num(r[4])) or 31
        worked = int(_num(r[5])) if r[5] not in (None, "") else wd
        rows.append(dict(
            name=" ".join(r[2].split()), dt=_text_date(r[1]), bps=bps, wd=wd,
            worked=worked, arrear=int(_num(r[6])), basic=int(_num(r[7])),
            hr=int(_num(r[8])), conv=int(_num(r[9])), medl=int(_num(r[10])),
            adhoc=int(_num(r[11])), spec=int(_num(r[12])), cpf=int(_num(r[14])),
            loan=int(_num(r[15])), lwo=int(_num(r[16])), cat=cat))
    if not rows:
        raise ValueError("No staff rows found in the salary sheet.")
    for r in rows:
        r["net"] = (r["arrear"] + r["basic"] + r["hr"] + r["conv"] + r["medl"]
                    + r["adhoc"] + r["spec"] - r["cpf"] - r["loan"] - r["lwo"])

    # ---- BPS scale, taken from the sheet itself
    by_bps = {}
    for r in rows:
        if r["bps"] is not None:
            by_bps.setdefault(r["bps"], []).append(r)
    for bps, grp in by_bps.items():
        hr, conv, medl, adhoc = Counter(
            (g["hr"], g["conv"], g["medl"], g["adhoc"]) for g in grp).most_common(1)[0][0]
        db.execute("INSERT OR REPLACE INTO bps_scale VALUES (?,?,?,?,?,?)",
                   (bps, hr, conv, medl, adhoc, max(g["cpf"] for g in grp)))
    scales = get_scales(db)

    # ---- bank advice sheets: designation, account number, bank
    brows = []
    for ws, bank in banks:
        for r in ws.iter_rows(min_row=15, values_only=True):
            r = list(r) + [None] * 8
            if (isinstance(r[0], (int, float)) and isinstance(r[1], str)
                    and r[3] and isinstance(r[5], (int, float))):
                brows.append(dict(name=" ".join(r[1].split()), des=str(r[2] or "").strip(),
                                  acct=_text_acct(r[3]), net=int(r[5]), bank=bank))
    used, loose = set(), []
    for r in rows:                                   # pass 1: exact name
        c = [i for i, b in enumerate(brows)
             if i not in used and _norm(b["name"]) == _norm(r["name"])]
        if len(c) > 1:
            c = [i for i in c if brows[i]["net"] == r["net"]] or c
        if c:
            used.add(c[0])
            r["b"] = brows[c[0]]
    for r in rows:                                   # pass 2: same first name + net
        if "b" in r:
            continue
        c = [i for i, b in enumerate(brows) if i not in used
             and b["net"] == r["net"] and _similar(b["name"], r["name"])]
        if c:
            used.add(c[0])
            r["b"] = brows[c[0]]
            loose.append((r["name"], brows[c[0]]["name"]))

    # ---- write employees and this month's inputs
    s = get_settings(db)
    if s["college_name"] == DEFAULTS["college_name"]:
        s["college_name"] = title.split(", Calculation")[0].strip()
        save_settings(db, s)
    existing = {(_norm(e["name"]), e["dt_appt"] or ""): e["id"]
                for e in db.execute("SELECT id, name, dt_appt FROM employees")}
    added = updated = 0
    adjustments, model_gaps = [], []
    for r in rows:
        sc = scales.get(r["bps"]) if r["bps"] is not None else None
        member = 1 if (sc and r["cpf"] > 0) else 0
        b = r.get("b") or {}
        key = (_norm(r["name"]), r["dt"])
        if key in existing:
            eid = existing[key]
            db.execute("UPDATE employees SET bps=?, basic=?, special=?, category=?, "
                       "cpf_member=?, active=1, designation=COALESCE(?, designation), "
                       "account_no=COALESCE(?, account_no), bank=COALESCE(?, bank) "
                       "WHERE id=?", (r["bps"], r["basic"], r["spec"], r["cat"], member,
                                      b.get("des"), b.get("acct"), b.get("bank"), eid))
            updated += 1
        else:
            cur = db.execute(
                "INSERT INTO employees(name, dt_appt, bps, basic, designation, "
                "account_no, bank, special, category, cpf_member) "
                "VALUES (?,?,?,?,?,?,?,?,?,?)",
                (r["name"], r["dt"], r["bps"], r["basic"], b.get("des"),
                 b.get("acct"), b.get("bank"), r["spec"], r["cat"], member))
            eid = cur.lastrowid
            existing[key] = eid
            added += 1
        calc = calculate(r["basic"], sc, r["spec"], member, s, r["wd"],
                         r["worked"], r["arrear"], r["loan"], 0)
        gross_sheet = r["arrear"] + r["basic"] + r["hr"] + r["conv"] + r["medl"] + r["adhoc"] + r["spec"]
        if calc["gross"] != gross_sheet:
            model_gaps.append(r["name"])
        other = r["lwo"] - calc["lwop_pay"]   # security deduction, rounding, etc.
        if other:
            adjustments.append((r["name"], other))
        db.execute("""INSERT OR REPLACE INTO monthly
            (month, emp_id, wd, worked, arrear, cp_loan, other, basic, hr, conv, medl,
             adhoc, special, cpf, name, dt_appt, bps, designation, account_no, bank, cpf_member)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                   (mkey, eid, r["wd"], r["worked"], r["arrear"], r["loan"], other,
                    r["basic"], r["hr"], r["conv"], r["medl"], r["adhoc"], r["spec"],
                    r["cpf"], r["name"], r["dt"], r["bps"], b.get("des"),
                    b.get("acct"), b.get("bank"), member))
    db.commit()
    return dict(year=year, month=month_no, total=len(rows), added=added,
                updated=updated, scales=sorted(by_bps), loose=loose,
                unmatched=[r["name"] for r in rows if "b" not in r],
                model_gaps=model_gaps, adjustments=adjustments)


# ------------------------------------------------------------- excel export
def export_xlsx(path, title, rows):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L

    wb = Workbook()
    ws = wb.active
    ws.title = "Salary Sheet"
    n_cols = len(HEADS)
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    base = Font(name="Arial", size=9)
    bold = Font(name="Arial", size=9, bold=True)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=12, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    hdr = 3
    for c, h in enumerate(HEADS, 1):
        cell = ws.cell(hdr, c, h)
        cell.font = bold
        cell.fill = PatternFill("solid", start_color="D9E1F2")
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = box
    ws.row_dimensions[hdr].height = 30

    first = hdr + 1
    for i, r in enumerate(rows):
        n = first + i
        vals = [r["sno"], r["dt"], r["name"], r["bps"], r["wd"], r["worked"],
                r["arrear"], r["basic"], r["hr"], r["conv"], r["medl"],
                r["adhoc"], r["spec"], f"=SUM(G{n}:M{n})", r["cpf"],
                r["cp_loan"], r["lwop"], f"=N{n}-O{n}-P{n}-Q{n}"]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(n, c, v)
            cell.font = base
            cell.border = box
            if c >= 7:
                cell.number_format = "#,##0"
    last = first + len(rows) - 1

    if rows:
        t = last + 1
        ws.cell(t, 3, "TOTAL")
        for c in range(7, n_cols + 1):
            ws.cell(t, c, f"=SUM({L(c)}{first}:{L(c)}{last})")
            ws.cell(t, c).number_format = "#,##0"
        for c in range(1, n_cols + 1):
            ws.cell(t, c).font = bold
            ws.cell(t, c).border = box

    widths = [6, 12, 28, 6, 8, 8, 11, 12, 11, 11, 11, 12, 11, 13, 10, 12,
              12, 13]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[L(c)].width = w
    ws.freeze_panes = ws.cell(first, 4)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)


# ------------------------------------------- bank statement and CPF statement
def _sheet_styles():
    from openpyxl.styles import Border, Font, PatternFill, Side
    thin = Side(style="thin")
    return (Border(left=thin, right=thin, top=thin, bottom=thin),
            Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True),
            PatternFill("solid", start_color="D9E1F2"))


def export_bank_xlsx(path, title, rows):
    """Name of Teacher | Account number | Bank | Net Salary."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font

    box, base, bold, fill = _sheet_styles()
    wb = Workbook()
    ws = wb.active
    ws.title = "Bank Statement"
    ws.merge_cells("A1:D1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=12, bold=True)
    for c, h in enumerate(["Name of Teacher", "Account number", "Bank",
                           "Net Salary"], 1):
        cell = ws.cell(3, c, h)
        cell.font, cell.border, cell.fill = bold, box, fill
        cell.alignment = Alignment(horizontal="center")
    lst = [r for r in rows if r["net"]]
    first = 4
    for i, r in enumerate(lst):
        n = first + i
        for c, v in enumerate([r["name"], r.get("acct") or "",
                               r.get("bank") or "", r["net"]], 1):
            cell = ws.cell(n, c, v)
            cell.font, cell.border = base, box
        ws.cell(n, 2).number_format = "@"
        ws.cell(n, 4).number_format = "#,##0"
    if lst:
        t = first + len(lst)
        ws.cell(t, 1, "TOTAL")
        ws.cell(t, 4, f"=SUM(D{first}:D{t - 1})").number_format = "#,##0"
        for c in range(1, 5):
            ws.cell(t, c).font, ws.cell(t, c).border = bold, box
    for col, w in zip("ABCD", (34, 28, 12, 16)):
        ws.column_dimensions[col].width = w
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)


def export_cpf_xlsx(path, s, rows, on_date):
    """Contributory Provident Fund statement (one line per CPF member)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, Side

    box, base, bold, fill = _sheet_styles()
    wb = Workbook()
    ws = wb.active
    ws.title = "CPF Statement"
    ws.merge_cells("A1:H1")
    ws["A1"] = f"CONTRIBUTORY PROVIDENT FUND- {s['college_name'].upper()}"
    ws["A1"].font = Font(name="Arial", size=12, bold=True, underline="single")

    pre = s["voucher_prefix"].strip()
    line = Border(bottom=Side(style="thin"))
    for r, text in ((3, f"{pre} S & W F Voucher No.".strip()),
                    (4, f"{pre} CPF Voucher No.".strip())):
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        ws.cell(r, 5, text).font = base
        ws.cell(r, 8).border = line

    lst = [r for r in rows if r["cpf"] or r["cp_loan"]]
    hdr, first = 8, 9
    last = first + len(lst) - 1
    ws.merge_cells("A6:C6")
    ws["A6"] = (f'="Rs. "&TEXT(SUM(H{first}:H{last}),"#,##0")&"/-"'
                if lst else "Rs. 0/-")
    ws["A6"].font = Font(name="Arial", size=11, bold=True, underline="single")
    ws["E6"] = "Dated:"
    ws["E6"].font = base
    ws.merge_cells("F6:H6")
    ws["F6"] = f"{on_date.day} {calendar.month_name[on_date.month]} {on_date.year}"
    ws["F6"].font = Font(name="Arial", size=10, underline="single")

    heads = ["S. No", "Name", "Designation", "A/C No",
             "Deduction from Individual pay", "From College Salary Fund",
             "Loan Install-ment", "Total"]
    for c, h in enumerate(heads, 1):
        cell = ws.cell(hdr, c, h)
        cell.font, cell.border, cell.fill = bold, box, fill
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    ws.row_dimensions[hdr].height = 44
    for i, r in enumerate(lst):
        n = first + i
        vals = [i + 1, r["name"], r.get("des") or "", r.get("cpf_no") or "",
                r["cpf"], r.get("cpf_college", r["cpf"]),
                r["cp_loan"] or None, f"=SUM(E{n}:G{n})"]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(n, c, v)
            cell.font, cell.border = base, box
            if c >= 5:
                cell.number_format = "#,##0"
    if lst:
        t = last + 1
        ws.cell(t, 2, "TOTAL")
        for c in range(5, 9):
            L = "EFGH"[c - 5]
            ws.cell(t, c, f"=SUM({L}{first}:{L}{last})").number_format = "#,##0"
        for c in range(1, 9):
            ws.cell(t, c).font, ws.cell(t, c).border = bold, box
    for col, w in zip("ABCDEFGH", (6, 28, 16, 9, 15, 15, 12, 13)):
        ws.column_dimensions[col].width = w
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)


# -------------------------------------------------------------- payslip PDF
def make_payslips(path, rows, s, month_label):
    """One A4 payslip per employee, all in a single PDF."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas
    from reportlab.platypus import Table, TableStyle

    W, H = A4
    left, right = 40, W - 40
    c = canvas.Canvas(path, pagesize=A4)
    logo = os.path.join(APP_DIR, "logo.png")

    def money(v):
        return f"{v:,}"

    def nil(v):
        return money(v) if v else "Nil"

    def segs(x, y, parts, size=10.5):
        """Draw (text, bold) pieces in a line; bold pieces are underlined."""
        for text, bold in parts:
            font = "Helvetica-Bold" if bold else "Helvetica"
            c.setFont(font, size)
            c.drawString(x, y, text)
            w = stringWidth(text, font, size)
            if bold:
                c.line(x, y - 1.5, x + w, y - 1.5)
            x += w

    for r in rows:
        y = H - 40
        if os.path.exists(logo):
            c.drawImage(logo, W / 2 - 45, y - 90, 90, 90,
                        preserveAspectRatio=True, mask="auto")
        y -= 115
        head = s["slip_header"]
        c.setFont("Helvetica-Bold", 10.5)
        c.drawCentredString(W / 2, y, head)
        hw = stringWidth(head, "Helvetica-Bold", 10.5)
        c.line(W / 2 - hw / 2, y - 2, W / 2 + hw / 2, y - 2)

        y -= 38
        segs(left, y, [("Pay slip of ", False), (r["name"], True),
                       (" for the Month: ", False), (month_label, True),
                       ("  Section: ", False), (r.get("section") or "-", True)])
        y -= 20
        segs(left, y, [("Designation: ", False), (r.get("des") or "-", True)])
        segs(300, y, [("BPS-", False), (str(r["bps"] or "-"), True)])
        segs(390, y, [("Working Days ", False),
                      (f"{r['worked']}/{r['wd']}", True)])

        acct = "".join(ch for ch in (r.get("acct") or "") if not ch.isspace())[:24]
        if acct:
            y -= 30
            c.setFont("Helvetica-Bold", 10.5)
            label = f"{r.get('bank') or ''} Account No".strip()
            c.drawString(left, y, label)
            x = max(130, left + stringWidth(label, "Helvetica-Bold", 10.5) + 12)
            for ch in acct:
                c.rect(x, y - 5, 15, 18)
                c.drawCentredString(x + 7.5, y, ch)
                x += 15
        y -= 22

        cr = [("R. Basic Pay", money(r["basic"])),
              ("House Rent Allowance", money(r["hr"])),
              ("Conveyance Allowance", money(r["conv"])),
              ("Medical Allowance", money(r["medl"])),
              (r.get("adhoc_label", "Adhoc Allowance"), money(r["adhoc"]))]
        if r["spec"]:
            cr.append(("Special Allowance", money(r["spec"])))
        cr.append(("Arrear/ Leave Encashment", money(r["arrear"])))
        dr = [("Leave Without Pay", nil(r.get("lwop_pay", r["lwop"] - r.get("other", 0)))),
              ("CPF Monthly Contribution", nil(r["cpf"])),
              ("Installment of CPF Loan", nil(r["cp_loan"])),
              ("Other Deduction", nil(r.get("other", 0)))]
        n = max(len(cr), len(dr))
        data = [["Cr", "Rs.", "Dr", "Rs."]]
        for i in range(n):
            a = cr[i] if i < len(cr) else ("", "")
            b = dr[i] if i < len(dr) else ("", "")
            data.append([a[0], a[1], b[0], b[1]])
        data.append(["Gross Pay", money(r["gross"]), "Total Deductions",
                     money(r["gross"] - r["net"])])
        data.append(["Net Pay", "", "", f"{r['net']:,.2f}"])
        gi, ni = n + 1, n + 2
        t = Table(data, colWidths=[175, 80, 170, 90])
        t.setStyle(TableStyle([
            ("FONT", (0, 0), (-1, -1), "Helvetica", 10.5),
            ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 10.5),
            ("FONT", (0, gi), (-1, ni), "Helvetica-Bold", 10.5),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("ALIGN", (3, 0), (3, -1), "RIGHT"),
            ("BOX", (0, 0), (-1, -1), 1, colors.black),
            ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.black),
            ("LINEABOVE", (0, gi), (-1, gi), 0.8, colors.black),
            ("LINEABOVE", (0, ni), (-1, ni), 0.8, colors.black),
            ("LINEAFTER", (1, 0), (1, -1), 0.5, colors.black),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        _, th = t.wrap(0, 0)
        t.drawOn(c, left, y - th)

        if s["signatory_name"] or s["signatory_title"]:
            sy = y - th - 90
            c.setFont("Helvetica-Bold", 11)
            c.drawString(390, sy, s["signatory_name"].upper())
            c.drawString(390, sy - 14, s["signatory_title"])
        c.showPage()
    c.save()



# -------------------------------------------------------------- Word / email

def export_salary_word(path, title, rows):
    """Create an editable Word document with teacher bank-transfer details."""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt
    doc = Document()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(14)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    headers = ["Teacher Name", "Account Number", "Net Salary", "Bank"]
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
    for r in rows:
        cells = table.add_row().cells
        cells[0].text = str(r.get("name", ""))
        cells[1].text = str(r.get("acct", ""))
        cells[2].text = f"{int(r.get('net', 0) or 0):,}"
        cells[3].text = str(r.get("bank", ""))
    doc.add_paragraph()
    p = doc.add_paragraph(f"Total Net Salary: Rs {sum(int(r.get('net', 0) or 0) for r in rows):,}")
    p.runs[0].bold = True
    doc.save(path)


def send_payslips_email(rows, s, month_label, temp_dir):
    """Generate one PDF per teacher and email it through configured SMTP."""
    username = str(s.get("smtp_username", "")).strip()
    password = str(s.get("smtp_password", ""))
    server = str(s.get("smtp_server", "")).strip()
    port = int(s.get("smtp_port", 587))
    sender = str(s.get("smtp_from", "")).strip() or username
    if not username or not password or not server or not sender:
        raise ValueError("Configure SMTP server, username, password and From email in Settings first.")
    os.makedirs(temp_dir, exist_ok=True)
    sent, skipped = [], []
    with smtplib.SMTP(server, port, timeout=30) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(username, password)
        for r in rows:
            recipient = str(r.get("email", "")).strip()
            if not recipient:
                skipped.append(f"{r.get('name', '')}: no email address")
                continue
            try:
                safe = re.sub(r"[^A-Za-z0-9._-]+", "_", r.get("name", "teacher"))[:60]
                pdf = os.path.join(temp_dir, f"Payslip_{safe}_{month_label.replace(' ', '_')}.pdf")
                make_payslips(pdf, [r], s, month_label)
                msg = EmailMessage()
                msg["Subject"] = f"Salary Payslip - {month_label} - {r.get('name', '')}"
                msg["From"] = sender
                msg["To"] = recipient
                msg.set_content(
                    f"Dear {r.get('name', '')},\n\nPlease find attached your salary payslip for {month_label}.\n\n"
                    "Regards,\n" + (s.get("signatory_name") or s.get("college_name", "Payroll Office")))
                with open(pdf, "rb") as f:
                    msg.add_attachment(f.read(), maintype="application", subtype="pdf",
                                       filename=os.path.basename(pdf))
                smtp.send_message(msg)
                sent.append(r.get("name", ""))
            except Exception as ex:
                skipped.append(f"{r.get('name', '')}: {ex}")
    return sent, skipped


# ---------------------------------------------------------------------- GUI
# Colors, loosely matching a clean "payroll desk" web dashboard look.
BG = "#f4efe7"          # content background
CARD = "#ffffff"
SIDEBAR = "#1c2531"
SIDEBAR_ACTIVE = "#2a3646"
ACCENT = "#dd8b2e"       # orange
TEXT = "#1c2531"
MUTED = "#6b7280"
BORDER = "#e4ddd0"
GREEN = "#2f8f5b"


def fmt(v):
    return f"{v:,}" if isinstance(v, int) else str(v)


EMP_FIELDS = ("name", "dt", "bps", "basic", "special", "designation", "section",
              "account", "bank", "email", "cpf_no", "category", "cpf_member")
EMP_LABELS = [("name", "Name (with Mr/Mrs)", 26), ("dt", "Appointed", 12),
              ("bps", "BPS (blank = contract)", 8), ("basic", "Running basic (Rs)", 12),
              ("special", "Special allowance (Rs)", 10), ("designation", "Designation", 22),
              ("section", "Section", 12), ("account", "Bank account no.", 24),
              ("bank", "Bank", 10), ("email", "Email", 28), ("cpf_no", "CPF A/C no.", 8),
              ("category", "Staff group", 10), ("cpf_member", "CPF member (Y/N)", 4)]
BPS_LABELS = [("bps", "BPS"), ("hr", "House Rent"), ("conv", "Conveyance"),
              ("medl", "Medical"), ("adhoc", "Adhoc / Increase"), ("cpf", "CPF deduction")]
NAV = [("overview", "\u2317", "Overview"), ("payroll", "\U0001F4C4", "Payroll desk"),
       ("teachers", "\U0001F465", "Teachers"), ("details", "\u25A3", "Employee details"), ("bps", "\U0001F4CA", "BPS Scales"),
       ("settings", "\u2699", "Settings")]


def style_setup(root):
    s = ttk.Style(root)
    try:
        s.theme_use("clam")
    except tk.TclError:
        pass
    root.configure(bg=BG)
    s.configure(".", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
    s.configure("TFrame", background=CARD)
    s.configure("Content.TFrame", background=BG)
    s.configure("Card.TFrame", background=CARD)
    s.configure("TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
    s.configure("Content.TLabel", background=BG, foreground=TEXT)
    s.configure("Muted.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 9))
    s.configure("CardMuted.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))
    s.configure("H1.TLabel", background=BG, foreground=TEXT, font=("Georgia", 22))
    s.configure("Eyebrow.TLabel", background=BG, foreground=ACCENT,
               font=("Segoe UI", 9, "bold"))
    s.configure("KPI.TLabel", background=CARD, foreground=TEXT, font=("Georgia", 20))
    s.configure("TEntry", fieldbackground="white", padding=4)
    s.configure("TCombobox", fieldbackground="white", padding=4)
    s.configure("Big.TCombobox", fieldbackground="white", padding=4,
               font=("Segoe UI", 14))
    s.configure("Big.TSpinbox", fieldbackground="white", padding=4,
               font=("Segoe UI", 14))
    s.configure("TButton", padding=(10, 6), font=("Segoe UI", 9))
    s.configure("Accent.TButton", background=ACCENT, foreground="white",
               font=("Segoe UI", 9, "bold"))
    s.map("Accent.TButton", background=[("active", "#c87a24")])
    s.configure("Treeview", rowheight=26, font=("Segoe UI", 10),
               background="white", fieldbackground="white", borderwidth=0)
    s.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
               background="#faf7f2", foreground=MUTED, relief="flat")
    s.map("Treeview", background=[("selected", "#fbe8cf")],
          foreground=[("selected", TEXT)])
    s.configure("TNotebook", background=CARD, borderwidth=0)
    s.configure("TNotebook.Tab", padding=(12, 6))


class Card(tk.Frame):
    """A plain white panel with a border, used for KPI tiles and sections."""
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1, bd=0, **kw)


class NavButton(tk.Frame):
    def __init__(self, parent, icon, text, command):
        super().__init__(parent, bg=SIDEBAR, cursor="hand2")
        self.command = command
        self.bar = tk.Frame(self, bg=SIDEBAR, width=3)
        self.bar.pack(side="left", fill="y")
        self.inner = tk.Frame(self, bg=SIDEBAR)
        self.inner.pack(side="left", fill="both", expand=True, padx=(10, 4), pady=9)
        self.icon_lbl = tk.Label(self.inner, text=icon, bg=SIDEBAR, fg="#cfd6e0",
                                 font=("Segoe UI", 11))
        self.icon_lbl.pack(side="left")
        self.text_lbl = tk.Label(self.inner, text=text, bg=SIDEBAR, fg="#cfd6e0",
                                 font=("Segoe UI", 10))
        self.text_lbl.pack(side="left", padx=8)
        for w in (self, self.bar, self.inner, self.icon_lbl, self.text_lbl):
            w.bind("<Button-1>", lambda e: self.command())

    def set_active(self, active):
        bg = SIDEBAR_ACTIVE if active else SIDEBAR
        self.configure(bg=bg)
        self.bar.configure(bg=ACCENT if active else SIDEBAR)
        self.inner.configure(bg=bg)
        self.icon_lbl.configure(bg=bg, fg="white" if active else "#cfd6e0")
        self.text_lbl.configure(bg=bg, fg="white" if active else "#cfd6e0",
                                font=("Segoe UI", 10, "bold" if active else "normal"))


class App(tk.Tk if tk else object):
    def __init__(self):
        super().__init__()
        self.title("Payroll Desk")
        self.geometry("1360x760")
        self.minsize(1080, 640)
        self.db = connect()
        self.rows, self.locked = [], False
        self.section = "payroll"
        style_setup(self)

        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True)
        self.sidebar = tk.Frame(outer, bg=SIDEBAR, width=220)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self.content = tk.Frame(outer, bg=BG)
        self.content.pack(side="left", fill="both", expand=True)

        self.pages = {}
        for key, _, _ in NAV:
            p = tk.Frame(self.content, bg=BG)
            self.pages[key] = p
        self.build_sidebar()
        self.build_overview()
        self.build_payroll()
        self.build_employees()
        self.build_details()
        self.build_bps()
        self.build_settings()

        self.load_employees()
        self.load_bps()
        self.load_payroll()
        self.show("payroll")

        # Windows sometimes leaves a freshly-built window blank until it is
        # repainted. Force one shortly after startup so nothing needs a
        # manual resize to appear.
        self.update_idletasks()
        self.after(60, self._nudge_redraw)

    def _nudge_redraw(self):
        try:
            self.geometry(self.geometry())
            self.update_idletasks()
        except tk.TclError:
            pass

    # ---- sidebar / navigation
    def build_sidebar(self):
        head = tk.Frame(self.sidebar, bg=SIDEBAR)
        head.pack(fill="x", pady=(22, 18), padx=16)
        badge = tk.Label(head, text="\u25A3", bg=ACCENT, fg="white",
                         font=("Segoe UI", 13, "bold"), width=2, height=1)
        badge.pack(side="left")
        txt = tk.Frame(head, bg=SIDEBAR)
        txt.pack(side="left", padx=10)
        tk.Label(txt, text="Payroll Desk", bg=SIDEBAR, fg="white",
                 font=("Georgia", 13)).pack(anchor="w")
        s = get_settings(self.db)
        tk.Label(txt, text=s["college_name"].upper(), bg=SIDEBAR, fg="#8b95a5",
                 font=("Segoe UI", 7, "bold")).pack(anchor="w")

        tk.Label(self.sidebar, text="WORKSPACE", bg=SIDEBAR, fg="#5d6779",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=20, pady=(6, 4))
        self.navbtns = {}
        for key, icon, text in NAV:
            b = NavButton(self.sidebar, icon, text, lambda k=key: self.show(k))
            b.pack(fill="x", padx=10, pady=1)
            self.navbtns[key] = b

        foot = tk.Frame(self.sidebar, bg="#161d27")
        foot.pack(side="bottom", fill="x")
        tk.Frame(foot, bg="#2a3646", height=1).pack(fill="x")
        finner = tk.Frame(foot, bg="#161d27")
        finner.pack(fill="x", padx=16, pady=12)
        self.foot_name = tk.Label(finner, text="", bg="#161d27", fg="white",
                                  font=("Segoe UI", 9, "bold"), anchor="w")
        self.foot_name.pack(fill="x")
        self.foot_title = tk.Label(finner, text="", bg="#161d27", fg="#8b95a5",
                                   font=("Segoe UI", 8), anchor="w")
        self.foot_title.pack(fill="x")
        self.refresh_sidebar_footer()

    def refresh_sidebar_footer(self):
        s = get_settings(self.db)
        self.foot_name.config(text=s["signatory_name"] or "Set up your name")
        self.foot_title.config(text=s["signatory_title"] or "in the Settings page")

    def show(self, key):
        self.section = key
        for k, b in self.navbtns.items():
            b.set_active(k == key)
        for k, p in self.pages.items():
            if k == key:
                p.pack(fill="both", expand=True)
            else:
                p.pack_forget()
        if key == "overview":
            self.load_overview()
        self.update_idletasks()

    # ---- page header helper
    def page_header(self, parent, eyebrow, title, subtitle=""):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x", padx=28, pady=(24, 10))
        left = tk.Frame(head, bg=BG)
        left.pack(side="left")
        tk.Label(left, text=eyebrow, bg=BG, fg=ACCENT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(left, text=title, bg=BG, fg=TEXT, font=("Georgia", 24)).pack(anchor="w")
        if subtitle:
            tk.Label(left, text=subtitle, bg=BG, fg=MUTED,
                     font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))
        return head

    def kpi_card(self, parent, label, value, note, accent):
        c = Card(parent)
        top = tk.Frame(c, bg=CARD)
        top.pack(fill="x", padx=16, pady=(14, 0))
        tk.Label(top, text=label.upper(), bg=CARD, fg=MUTED,
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        dot = tk.Frame(top, bg=accent, width=8, height=8)
        dot.pack(side="right")
        val = tk.Label(c, text=value, bg=CARD, fg=TEXT, font=("Georgia", 21))
        val.pack(anchor="w", padx=16, pady=(4, 0))
        tk.Label(c, text=note, bg=CARD, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(0, 14))
        return c, val

    # ---- overview page
    def build_overview(self):
        f = self.pages["overview"]
        self.page_header(f, "WORKSPACE", "Overview", "A quick look at your college payroll.")
        wrap = tk.Frame(f, bg=BG)
        wrap.pack(fill="x", padx=28)
        self.ov_cards = {}
        for i, key in enumerate(("staff", "gross", "net")):
            c, val = self.kpi_card(wrap, {"staff": "Active staff", "gross": "Latest gross payroll",
                                          "net": "Latest net payroll"}[key], "-", "", ACCENT)
            c.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 12, 0))
            wrap.grid_columnconfigure(i, weight=1)
            self.ov_cards[key] = val
        tip = Card(f)
        tip.pack(fill="x", padx=28, pady=20)
        tk.Label(tip, text="Getting around", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16, pady=(14, 4))
        for line in ("Payroll desk - open a month, edit days and deductions, lock it, "
                    "and print payslips, the bank statement and the CPF statement.",
                    "Teachers - add or import staff and their bank details.",
                    "BPS Scales - set the fixed allowances for each pay grade.",
                    "Settings - your college name, payslip heading and signatory."):
            tk.Label(tip, text="\u2022  " + line, bg=CARD, fg=MUTED, wraplength=760,
                     justify="left", font=("Segoe UI", 9)).pack(
                anchor="w", padx=16, pady=2)
        tk.Frame(tip, bg=CARD, height=10).pack()

    def load_overview(self):
        n = len(self.rows)
        self.ov_cards["staff"].config(text=str(n))
        self.ov_cards["gross"].config(text=f"Rs {sum(r['gross'] for r in self.rows):,}")
        self.ov_cards["net"].config(text=f"Rs {sum(r['net'] for r in self.rows):,}")

    # ---- payroll desk (main working page)
    def build_payroll(self):
        f = self.pages["payroll"]
        head = tk.Frame(f, bg=BG)
        head.pack(fill="x", padx=28, pady=(24, 6))
        left = tk.Frame(head, bg=BG)
        left.pack(side="left", fill="x", expand=True)
        tk.Label(left, text="PAYROLL DESK", bg=BG, fg=ACCENT,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        row = tk.Frame(left, bg=BG)
        row.pack(anchor="w", pady=(2, 0))
        self.month = ttk.Combobox(row, values=list(calendar.month_name)[1:],
                                  state="readonly", width=12,
                                  style="Big.TCombobox")
        self.month.current(date.today().month - 1)
        self.month.pack(side="left")
        self.month.bind("<<ComboboxSelected>>", lambda e: self.load_payroll())
        self.year = tk.IntVar(value=date.today().year)
        ttk.Spinbox(row, from_=2020, to=2100, textvariable=self.year, width=6,
                    command=self.load_payroll, style="Big.TSpinbox").pack(
            side="left", padx=6)
        self.lock_pill = tk.Label(row, text="", bg="#fdecd2", fg="#9a5a12",
                                  font=("Segoe UI", 8, "bold"), padx=8, pady=2)
        self.lock_pill.pack(side="left", padx=10)

        right = tk.Frame(head, bg=BG)
        right.pack(side="right")
        ttk.Button(right, text="\u2193 Bank statement", command=self.do_bank).pack(
            side="left", padx=3)
        ttk.Button(right, text="\u2193 CPF statement", command=self.do_cpf).pack(
            side="left", padx=3)
        ttk.Button(right, text="\u2193 Salary sheet (Excel)", command=self.do_export).pack(
            side="left", padx=3)
        ttk.Button(right, text="\u2193 Bank details (Word)", command=self.do_word).pack(
            side="left", padx=3)
        ttk.Button(right, text="\U0001F4C4 Payslips (all)", style="Accent.TButton",
                  command=lambda: self.do_payslips(False)).pack(side="left", padx=3)
        ttk.Button(right, text="\u2709 Email payslips", command=self.do_email_payslips).pack(
            side="left", padx=3)

        row2 = tk.Frame(f, bg=BG)
        row2.pack(fill="x", padx=28, pady=(0, 4))
        self.lock_btn = ttk.Button(row2, text="Lock month", command=self.do_lock)
        self.lock_btn.pack(side="left")
        self.unlock_btn = ttk.Button(row2, text="Unlock month", command=self.do_unlock)
        self.unlock_btn.pack(side="left", padx=6)
        ttk.Button(row2, text="Payslip (selected)",
                  command=lambda: self.do_payslips(True)).pack(side="left", padx=6)
        ttk.Button(row2, text="Reload", command=self.load_payroll).pack(side="left", padx=6)

        cards = tk.Frame(f, bg=BG)
        cards.pack(fill="x", padx=28, pady=14)
        _, self.kpi_gross = self.kpi_card(cards, "Gross payroll", "Rs 0", "", ACCENT)
        self.kpi_gross.master.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        _, self.kpi_net = self.kpi_card(cards, "Net salary", "Rs 0", "Ready after deductions", GREEN)
        self.kpi_net.master.grid(row=0, column=1, sticky="ew", padx=(0, 12))
        _, self.kpi_ded = self.kpi_card(cards, "Deductions", "Rs 0", "CPF and other deductions", "#c0392b")
        self.kpi_ded.master.grid(row=0, column=2, sticky="ew")
        for i in range(3):
            cards.grid_columnconfigure(i, weight=1)

        table_card = Card(f)
        table_card.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        tbar = tk.Frame(table_card, bg=CARD)
        tbar.pack(fill="x", padx=14, pady=10)
        tk.Label(tbar, text="Find a teacher:", bg=CARD, fg=MUTED,
                font=("Segoe UI", 8)).pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar()
        se = ttk.Entry(tbar, textvariable=self.search_var, width=28)
        se.pack(side="left")
        self.search_var.trace_add("write", lambda *a: self.refresh_table())
        self.row_count_lbl = tk.Label(tbar, text="0 rows", bg=CARD, fg=MUTED,
                                      font=("Segoe UI", 8, "bold"))
        self.row_count_lbl.pack(side="left", padx=12)
        tk.Label(tbar, text="Double-click a row to edit days / arrears / deductions",
                bg=CARD, fg=MUTED, font=("Segoe UI", 8)).pack(side="right")

        wrap = tk.Frame(table_card, bg=CARD)
        wrap.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        ids = [f"c{i}" for i in range(len(HEADS))]
        self.ptree = ttk.Treeview(wrap, columns=ids, show="headings")
        for cid, h in zip(ids, HEADS):
            self.ptree.heading(cid, text=h.replace("\n", " "))
            self.ptree.column(cid, width=140 if h.startswith("Name") else 84,
                              anchor="w" if h.startswith(("Name", "Dt")) else "e")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.ptree.yview)
        hsb = ttk.Scrollbar(table_card, orient="horizontal", command=self.ptree.xview)
        self.ptree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.ptree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(fill="x")
        self.ptree.bind("<Double-1>", lambda e: self.edit_inputs())
        self.ptree.bind("<ButtonRelease-1>", lambda e: self.open_selected_employee())
        self.ptree.tag_configure("odd", background="#faf8f4")

    def month_key(self):
        return f"{self.year.get()}-{self.month.current() + 1:02d}"

    def month_title(self):
        s = get_settings(self.db)
        return (f"{s['college_name']} - Calculation of Staff Salary for the "
                f"Month of {self.month.get()} {self.year.get()}")

    def load_payroll(self):
        self.rows, self.locked = month_rows(self.db, self.month_key())
        self.refresh_table()
        g = sum(r["gross"] for r in self.rows)
        n = sum(r["net"] for r in self.rows)
        d = g - n
        self.kpi_gross.config(text=f"Rs {g:,}")
        self.kpi_net.config(text=f"Rs {n:,}")
        self.kpi_ded.config(text=f"Rs {d:,}")
        self.lock_pill.config(text="ARCHIVED - READ ONLY" if self.locked
                              else "DRAFT - recalculates live")
        self.lock_btn.state(["disabled" if self.locked else "!disabled"])
        self.unlock_btn.state(["!disabled" if self.locked else "disabled"])
        if self.section == "overview":
            self.load_overview()

    def refresh_table(self):
        self.ptree.delete(*self.ptree.get_children())
        q = self.search_var.get().strip().lower() if hasattr(self, "search_var") else ""
        shown = [r for r in self.rows if q in r["name"].lower()] if q else self.rows
        for i, r in enumerate(shown):
            self.ptree.insert("", "end", iid=str(r["emp_id"]),
                              values=[fmt(r[k]) for k in KEYS],
                              tags=("odd",) if i % 2 else ())
        self.row_count_lbl.config(text=f"{len(shown)} rows")

    def open_selected_employee(self):
        sel = self.ptree.selection()
        if sel:
            self.open_employee_details(int(sel[0]))

    def edit_inputs(self):
        self.open_selected_employee()

    def build_details(self):
        f = self.pages["details"]
        self.page_header(f, "WORKSPACE", "Employee details",
                         "Each tab is the editable current-month record for one teacher.")
        bar = tk.Frame(f, bg=BG)
        bar.pack(fill="x", padx=28, pady=(0, 8))
        ttk.Button(bar, text="Refresh current month", command=self.refresh_detail_tabs).pack(side="left")
        tk.Label(bar, text="Changes here are saved for the selected month only; master employee data remains unchanged.",
                 bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(side="left", padx=12)
        card = Card(f)
        card.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        self.detail_nb = ttk.Notebook(card)
        self.detail_nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.detail_tabs = {}

    def open_employee_details(self, eid):
        if self.locked:
            messagebox.showinfo("Archived", "This month is locked. Unlock it before editing employee details.")
            return
        row = next((r for r in self.rows if int(r["emp_id"]) == eid), None)
        if not row:
            return
        self.show("details")
        tab_id = self.detail_tabs.get(eid)
        if tab_id and self.detail_nb.exists(tab_id):
            self.detail_nb.select(tab_id)
            return
        tab = tk.Frame(self.detail_nb, bg=BG)
        self.detail_nb.add(tab, text=(row["name"] or "Employee")[:24])
        self.detail_tabs[eid] = str(tab)
        canvas = tk.Canvas(tab, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=BG)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        vars_ = {}
        fields = [
            ("name", "Teacher name", "text"), ("dt", "Appointment date", "text"),
            ("bps", "BPS", "num"), ("designation", "Designation", "text"),
            ("section", "Section", "text"), ("acct", "Account number", "text"),
            ("bank", "Bank (ABL / HBL etc.)", "text"), ("email", "Email ID", "text"),
            ("cpf_no", "CPF account no.", "text"), ("cpf_member", "CPF member (Y/N)", "text"),
            ("wd", "Working days", "num"), ("worked", "Days worked", "num"),
            ("basic", "Running Basic", "money"), ("hr", "House Rent Allowance", "money"),
            ("conv", "Conveyance Allowance", "money"), ("medl", "Medical Allowance", "money"),
            ("adhoc", row.get("adhoc_label", "Adhoc / Basic Increase"), "money"),
            ("spec", "Special Allowance", "money"), ("cpf", "CP Fund / CPF Deduction", "money"),
            ("arrear", "Encashment / Arrear", "money"), ("cp_loan", "CP Loan Instalment", "money"),
            ("other", "Other Deduction", "money")]
        tk.Label(inner, text=f"Current month: {self.month.get()} {self.year.get()}",
                 bg=BG, fg=ACCENT, font=("Segoe UI", 10, "bold")).grid(
                     row=0, column=0, columnspan=4, sticky="w", padx=18, pady=(14, 8))
        for i, (k, lab, kind) in enumerate(fields):
            r, c = divmod(i, 2)
            box = tk.Frame(inner, bg=CARD)
            box.grid(row=r+1, column=c, sticky="ew", padx=8, pady=6)
            ttk.Label(box, text=lab).pack(anchor="w", padx=10, pady=(8, 2))
            vars_[k] = tk.StringVar(value=str(row.get(k, "")))
            ttk.Entry(box, textvariable=vars_[k], width=34).pack(fill="x", padx=10, pady=(0, 8))
        for c in range(2): inner.grid_columnconfigure(c, weight=1)
        result = tk.Frame(inner, bg=CARD)
        result.grid(row=(len(fields)+1)//2+1, column=0, columnspan=2, sticky="ew", padx=8, pady=10)
        calc_lbl = tk.Label(result, text="", bg=CARD, fg=TEXT, font=("Segoe UI", 11, "bold"))
        calc_lbl.pack(anchor="w", padx=12, pady=8)

        def to_int(k):
            raw = vars_[k].get().replace(",", "").strip()
            return int(float(raw or 0))

        def save():
            try:
                vals = {k: to_int(k) for k in ("bps", "wd", "worked", "basic", "hr", "conv", "medl", "adhoc", "spec", "cpf", "arrear", "cp_loan", "other")}
                vals["bps"] = None if not vars_["bps"].get().strip() else vals["bps"]
                member = 0 if vars_["cpf_member"].get().strip().upper().startswith("N") else 1
                vals.update(name=vars_["name"].get().strip(), dt=vars_["dt"].get().strip(),
                            designation=vars_["designation"].get().strip(), section=vars_["section"].get().strip(),
                            acct=vars_["acct"].get().strip(), bank=vars_["bank"].get().strip(),
                            email=vars_["email"].get().strip(), cpf_no=vars_["cpf_no"].get().strip(),
                            cpf_member=member, emp_id=eid)
                if not vals["name"]:
                    raise ValueError("Teacher name is required")
                sc = {"hr": vals["hr"], "conv": vals["conv"], "medl": vals["medl"],
                      "adhoc": vals["adhoc"], "cpf": vals["cpf"]}
                calc = calculate(vals["basic"], sc, vals["spec"], member,
                                 get_settings(self.db), vals["wd"], vals["worked"],
                                 vals["arrear"], vals["cp_loan"], vals["other"])
                row2 = dict(vals, **calc)
                save_monthly_employee(self.db, self.month_key(), row2)
                self.load_payroll()
                messagebox.showinfo("Saved", f"Current-month details saved for {vals['name']}.", parent=tab)
                self.detail_nb.tab(tab, text=vals["name"][:24])
                calc_lbl.config(text=f"Gross: Rs {calc['gross']:,}    Net Salary: Rs {calc['net']:,}")
            except ValueError as ex:
                messagebox.showerror("Invalid value", str(ex), parent=tab)

        ttk.Button(inner, text="Save current-month details", style="Accent.TButton",
                   command=save).grid(row=(len(fields)+1)//2+2, column=0, columnspan=2, pady=(4, 18))
        calc_lbl.config(text=f"Gross: Rs {row['gross']:,}    Net Salary: Rs {row['net']:,}")
        self.detail_nb.select(tab)

    def refresh_detail_tabs(self):
        if not self.detail_tabs:
            return
        self.rows, self.locked = month_rows(self.db, self.month_key())
        for eid in list(self.detail_tabs):
            tab_id = self.detail_tabs[eid]
            if self.detail_nb.exists(tab_id):
                self.detail_nb.forget(tab_id)
        self.detail_tabs.clear()


    def do_lock(self):
        if self.locked or not self.rows:
            return
        if messagebox.askyesno("Lock month", "Lock this month? Figures are saved "
                               "as final and will not change if you edit "
                               "settings or basic pay later."):
            lock_month(self.db, self.month_key(), self.rows)
            self.load_payroll()

    def do_unlock(self):
        if self.locked and messagebox.askyesno(
                "Unlock", "Unlock this month? It will recalculate from current settings."):
            unlock_month(self.db, self.month_key())
            self.load_payroll()

    def save_excel(self, name, build):
        if not self.rows:
            messagebox.showinfo("Nothing to export", "Add employees first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=f"{name}_{self.month_key()}.xlsx")
        if path:
            try:
                build(path)
                messagebox.showinfo("Done", f"Saved:\n{path}")
            except Exception as ex:  # e.g. file open in Excel
                messagebox.showerror("Export failed", str(ex))

    def do_export(self):
        self.save_excel("Salary_Sheet",
                        lambda p: export_xlsx(p, self.month_title(), self.rows))

    def do_bank(self):
        s = get_settings(self.db)
        title = (f"{s['college_name']} - Salary Transfer Statement, "
                 f"{self.month.get()} {self.year.get()}")
        self.save_excel("Bank_Statement",
                        lambda p: export_bank_xlsx(p, title, self.rows))

    def do_cpf(self):
        s = get_settings(self.db)
        self.save_excel("CPF_Statement",
                        lambda p: export_cpf_xlsx(p, s, self.rows, date.today()))

    def do_word(self):
        if not self.rows:
            messagebox.showinfo("Nothing to export", "Add employees first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Word document", "*.docx")],
            initialfile=f"Teacher_Bank_Details_{self.month_key()}.docx")
        if not path:
            return
        try:
            title = f"Teacher Salary Bank Details - {self.month.get()} {self.year.get()}"
            export_salary_word(path, title, self.rows)
            messagebox.showinfo("Done", f"Word file saved:\n{path}")
        except ImportError:
            messagebox.showerror("Missing package", "Run once: pip install python-docx")
        except Exception as ex:
            messagebox.showerror("Word export failed", str(ex))

    def do_email_payslips(self):
        if not self.rows:
            messagebox.showinfo("Nothing to send", "Add employees first.")
            return
        missing = [r["name"] for r in self.rows if not str(r.get("email", "")).strip()]
        if missing:
            if not messagebox.askyesno("Missing email IDs",
                    f"{len(missing)} teacher(s) have no email ID. Continue with the others?"):
                return
        if not messagebox.askyesno("Email payslips",
                f"Send {len(self.rows)-len(missing)} payslip(s) for {self.month.get()} {self.year.get()} by email?"):
            return
        import tempfile
        try:
            sent, skipped = send_payslips_email(
                self.rows, get_settings(self.db), f"{self.month.get()} {self.year.get()}",
                tempfile.mkdtemp(prefix="payroll_payslips_"))
            msg = f"Sent: {len(sent)}\nSkipped: {len(skipped)}"
            if skipped:
                msg += "\n\n" + "\n".join(skipped[:15])
                if len(skipped) > 15: msg += "\n..."
            messagebox.showinfo("Email result", msg)
        except ImportError:
            messagebox.showerror("Missing package", "Run once: pip install reportlab")
        except Exception as ex:
            messagebox.showerror("Email failed", str(ex))

    def do_payslips(self, only_selected):
        rows = self.rows
        if only_selected:
            sel = self.ptree.selection()
            if not sel:
                messagebox.showinfo("Select", "Select an employee row first.")
                return
            rows = [r for r in rows if str(r["emp_id"]) in sel]
        if not rows:
            messagebox.showinfo("Nothing to print", "Add employees first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialfile=f"Payslips_{self.month_key()}.pdf")
        if not path:
            return
        try:
            make_payslips(path, rows, get_settings(self.db),
                          f"{self.month.get()} {self.year.get()}")
            messagebox.showinfo("Done", f"Saved:\n{path}")
        except ImportError:
            messagebox.showerror("Missing package",
                                 "Run this once in Command Prompt:\n\npip install reportlab")
        except Exception as ex:  # e.g. file open in a PDF viewer
            messagebox.showerror("Payslip failed", str(ex))

    # ---- teachers (employees)
    def build_employees(self):
        f = self.pages["teachers"]
        self.page_header(f, "WORKSPACE", "Teachers", "Staff records used to build every month's payroll.")
        card = Card(f)
        card.pack(fill="x", padx=28, pady=(4, 12))
        tk.Label(card, text="Add or edit a teacher", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16, pady=(14, 6))
        form = tk.Frame(card, bg=CARD)
        form.pack(fill="x", padx=16, pady=(0, 6))
        self.ev = {k: tk.StringVar() for k in EMP_FIELDS}
        for i, (k, lab, w) in enumerate(EMP_LABELS):
            r, c = divmod(i, 4)
            ttk.Label(form, text=lab).grid(row=r * 2, column=c, padx=6, pady=(4, 0), sticky="w")
            ttk.Entry(form, textvariable=self.ev[k], width=w).grid(
                row=r * 2 + 1, column=c, padx=6, pady=(0, 8), sticky="w")
        bar = tk.Frame(card, bg=CARD)
        bar.pack(fill="x", padx=16, pady=(0, 14))
        ttk.Button(bar, text="Add", style="Accent.TButton", command=self.emp_add).pack(
            side="left", padx=(0, 6))
        ttk.Button(bar, text="Update selected", command=self.emp_update).pack(
            side="left", padx=3)
        ttk.Button(bar, text="Delete selected", command=self.emp_delete).pack(
            side="left", padx=3)
        ttk.Button(bar, text="Clear form", command=self.emp_clear).pack(side="left", padx=3)
        ttk.Button(bar, text="Import from salary workbook...",
                  command=self.emp_import).pack(side="right")

        list_card = Card(f)
        list_card.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        cols = ("sno", "name", "designation", "dt", "bps", "basic", "special", "category")
        self.etree = ttk.Treeview(list_card, columns=cols, show="headings")
        for c, h, w in zip(cols, ("S.NO", "Name", "Designation", "Dt of Appt", "BPS",
                                  "Running Basic", "Special", "Group"),
                           (50, 240, 160, 100, 50, 110, 80, 80)):
            self.etree.heading(c, text=h)
            self.etree.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(list_card, orient="vertical", command=self.etree.yview)
        self.etree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", pady=1)
        self.etree.pack(fill="both", expand=True, padx=1, pady=1)
        self.etree.bind("<<TreeviewSelect>>", self.emp_select)

    def emp_values(self):
        g = lambda k: self.ev[k].get().strip()
        if not g("name"):
            messagebox.showerror("Missing", "Name is required.")
            return None
        try:
            basic = int(float(g("basic").replace(",", "") or 0))
            special = int(float(g("special").replace(",", "") or 0))
            bps = int(g("bps")) if g("bps") else None
        except ValueError:
            messagebox.showerror("Invalid", "BPS, basic pay and special allowance "
                                            "must be numbers.")
            return None
        member = 0 if g("cpf_member").upper().startswith("N") else 1
        return (g("name"), g("dt"), bps, basic, g("designation"), g("section"),
                g("account"), g("bank"), g("email"), g("cpf_no"), special, g("category"), member)

    def emp_add(self):
        v = self.emp_values()
        if v:
            self.db.execute("INSERT INTO employees(name, dt_appt, bps, basic, "
                            "designation, section, account_no, bank, email, cpf_no, "
                            "special, category, cpf_member) "
                            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", v)
            self.db.commit()
            self.emp_clear()
            self.after_emp_change()

    def emp_selected_id(self):
        sel = self.etree.selection()
        return int(sel[0]) if sel else None

    def emp_update(self):
        eid, v = self.emp_selected_id(), self.emp_values()
        if eid and v:
            self.db.execute("UPDATE employees SET name=?, dt_appt=?, bps=?, "
                            "basic=?, designation=?, section=?, account_no=?, "
                            "bank=?, email=?, cpf_no=?, special=?, category=?, cpf_member=? "
                            "WHERE id=?", v + (eid,))
            self.db.commit()
            self.after_emp_change()
        elif not eid:
            messagebox.showinfo("Select", "Select an employee in the list first.")

    def emp_delete(self):
        eid = self.emp_selected_id()
        if eid and messagebox.askyesno(
                "Delete", "Remove this employee from future payrolls?\n"
                          "Locked past months are not affected."):
            self.db.execute("UPDATE employees SET active=0 WHERE id=?", (eid,))
            self.db.commit()
            self.emp_clear()
            self.after_emp_change()

    def emp_clear(self):
        for v in self.ev.values():
            v.set("")

    def emp_select(self, _):
        eid = self.emp_selected_id()
        if eid:
            e = self.db.execute("SELECT * FROM employees WHERE id=?", (eid,)).fetchone()
            self.ev["name"].set(e["name"])
            self.ev["dt"].set(e["dt_appt"] or "")
            self.ev["bps"].set(e["bps"] if e["bps"] is not None else "")
            self.ev["basic"].set(e["basic"])
            self.ev["special"].set(e["special"] or 0)
            self.ev["designation"].set(e["designation"] or "")
            self.ev["section"].set(e["section"] or "")
            self.ev["account"].set(e["account_no"] or "")
            self.ev["bank"].set(e["bank"] or "")
            self.ev["email"].set(e["email"] or "")
            self.ev["cpf_no"].set(e["cpf_no"] or "")
            self.ev["category"].set(e["category"] or "")
            self.ev["cpf_member"].set("Y" if e["cpf_member"] else "N")

    def load_employees(self):
        self.etree.delete(*self.etree.get_children())
        emps = self.db.execute("SELECT * FROM employees WHERE active=1 ORDER BY id")
        for i, e in enumerate(emps.fetchall(), 1):
            self.etree.insert("", "end", iid=str(e["id"]), values=(
                i, e["name"], e["designation"] or "", e["dt_appt"] or "",
                e["bps"] if e["bps"] is not None else "", fmt(e["basic"]),
                fmt(e["special"] or 0), e["category"] or ""))

    def after_emp_change(self):
        self.load_employees()
        self.load_payroll()

    def emp_import(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            rep = import_workbook(self.db, path)
        except Exception as ex:
            messagebox.showerror("Import failed", str(ex))
            return
        self.year.set(rep["year"])
        self.month.current(rep["month"] - 1)
        self.load_employees()
        self.load_bps()
        self.load_payroll()
        lines = [f"{calendar.month_name[rep['month']]} {rep['year']}: {rep['total']} staff "
                 f"read ({rep['added']} new, {rep['updated']} updated).",
                 "BPS scale updated for grades: " + ", ".join(map(str, rep["scales"])) + "."]
        if rep["model_gaps"]:
            lines.append("Allowances differ from the BPS scale for: "
                         + ", ".join(rep["model_gaps"]))
        if rep["loose"]:
            lines.append("Bank details matched by similar name and same net pay - "
                         "please check:\n  " + "\n  ".join(
                             f"{a}  <->  {b}" for a, b in rep["loose"]))
        if rep["unmatched"]:
            lines.append("No bank details found for: " + ", ".join(rep["unmatched"]))
        lines.append(f"{len(rep['adjustments'])} people have an amount in 'Other "
                     "deduction' (e.g. security deposits, small rounding).")
        messagebox.showinfo("Import finished", "\n\n".join(lines))
        self.show("payroll")

    # ---- BPS scales
    def build_bps(self):
        f = self.pages["bps"]
        self.page_header(f, "WORKSPACE", "BPS Scales",
                         "Fixed monthly allowances for each pay grade.")
        card = Card(f)
        card.pack(fill="x", padx=28, pady=(4, 12))
        tk.Label(card, text="Add or edit a grade", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16, pady=(14, 6))
        form = tk.Frame(card, bg=CARD)
        form.pack(padx=16, pady=(0, 6), anchor="w")
        self.bv = {k: tk.StringVar() for k, _ in BPS_LABELS}
        for i, (k, lab) in enumerate(BPS_LABELS):
            ttk.Label(form, text=lab).grid(row=0, column=i, padx=6, sticky="w")
            ttk.Entry(form, textvariable=self.bv[k], width=10).grid(
                row=1, column=i, padx=6, pady=(0, 8))
        bar = tk.Frame(card, bg=CARD)
        bar.pack(fill="x", padx=16, pady=(0, 14))
        ttk.Button(bar, text="Save (add / update)", style="Accent.TButton",
                  command=self.bps_save).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Delete selected", command=self.bps_delete).pack(
            side="left", padx=3)
        tk.Label(f, text="Staff with no BPS (contract) get their running basic only. "
                        "The running basic itself is entered per person on the "
                        "Teachers page.", bg=BG, fg=MUTED,
                font=("Segoe UI", 9)).pack(anchor="w", padx=30, pady=(0, 8))

        list_card = Card(f)
        list_card.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        cols = [k for k, _ in BPS_LABELS]
        self.btree = ttk.Treeview(list_card, columns=cols, show="headings", height=14)
        for k, lab in BPS_LABELS:
            self.btree.heading(k, text=lab)
            self.btree.column(k, width=140, anchor="e")
        self.btree.pack(fill="both", expand=True, padx=1, pady=1)
        self.btree.bind("<<TreeviewSelect>>", self.bps_select)

    def load_bps(self):
        self.btree.delete(*self.btree.get_children())
        for r in self.db.execute("SELECT * FROM bps_scale ORDER BY bps"):
            self.btree.insert("", "end", iid=str(r["bps"]),
                              values=[fmt(r[k]) for k, _ in BPS_LABELS])

    def bps_select(self, _):
        sel = self.btree.selection()
        if sel:
            r = self.db.execute("SELECT * FROM bps_scale WHERE bps=?", (int(sel[0]),)).fetchone()
            for k, _ in BPS_LABELS:
                self.bv[k].set(r[k])

    def bps_save(self):
        try:
            v = [int(float(self.bv[k].get().replace(",", "") or 0)) for k, _ in BPS_LABELS]
        except ValueError:
            messagebox.showerror("Invalid", "Enter numbers only.")
            return
        if not self.bv["bps"].get().strip():
            messagebox.showerror("Missing", "Enter the BPS grade.")
            return
        self.db.execute("INSERT OR REPLACE INTO bps_scale VALUES (?,?,?,?,?,?)", v)
        self.db.commit()
        self.load_bps()
        self.load_payroll()

    def bps_delete(self):
        sel = self.btree.selection()
        if sel and messagebox.askyesno("Delete", "Delete this grade? Staff on it will "
                                                 "get no allowances until it is re-added."):
            self.db.execute("DELETE FROM bps_scale WHERE bps=?", (int(sel[0]),))
            self.db.commit()
            self.load_bps()
            self.load_payroll()

    # ---- settings
    def build_settings(self):
        f = self.pages["settings"]
        self.page_header(f, "WORKSPACE", "Settings",
                         "College identity, payslip heading and CPF options.")
        card = Card(f)
        card.pack(fill="x", padx=28, pady=(4, 20))
        s = get_settings(self.db)
        self.sv = {}
        form = tk.Frame(card, bg=CARD)
        form.pack(padx=16, pady=16, anchor="w")
        for i, (k, lab) in enumerate(SETTING_LABELS.items()):
            ttk.Label(form, text=lab).grid(row=i, column=0, sticky="w", padx=6, pady=6)
            self.sv[k] = tk.StringVar(value=str(s[k]))
            ent = ttk.Entry(form, textvariable=self.sv[k],
                      width=44 if k in ("college_name", "slip_header") else 28,
                      show="*" if k == "smtp_password" else "")
            ent.grid(
                row=i, column=1, padx=6)
        ttk.Button(form, text="Save settings", style="Accent.TButton",
                  command=self.save_set).grid(
            row=len(SETTING_LABELS), column=0, columnspan=2, pady=(10, 0), sticky="w")

    def save_set(self):
        try:
            s = {k: type(DEFAULTS[k])(v.get().strip()) for k, v in self.sv.items()}
        except ValueError:
            messagebox.showerror("Invalid", "Numeric settings must be numbers.")
            return
        save_settings(self.db, s)
        self.refresh_sidebar_footer()
        self.load_payroll()
        messagebox.showinfo("Saved", "Settings saved.")


if __name__ == "__main__":
    if tk is None:
        print("tkinter is not available in this Python installation.\n"
              "Reinstall Python from python.org and tick 'tcl/tk and IDLE' "
              "in the installer's optional features.")
        input("Press Enter to close...")
        sys.exit(1)
    try:
        App().mainloop()
    except Exception:
        import traceback
        err = traceback.format_exc()
        log_path = os.path.join(APP_DIR, "payroll_error.log")
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(err + "\n")
        except OSError:
            pass
        print(err)
        print(f"(Also saved to {log_path})")
        input("Press Enter to close...")
        sys.exit(1)
