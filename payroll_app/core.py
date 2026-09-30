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
    from .reports import build_payslips
    build_payslips(path, rows, s, month_label, APP_DIR)


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


