import sys, json
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from payroll_app import core
from openpyxl import load_workbook
p='outputs/payroll-testing-50/Payroll_Test_Data_50_Employees.xlsx'
db=core.connect(':memory:')
result=core.import_workbook(db,p)
rows,locked=core.month_rows(db,'2026-09')
expected=json.load(open('tmp/dummy-sheet/expected.json'))
assert len(rows)==50
assert not result['unmatched'] and not result['model_gaps']
assert all(r['net']==next(e['net'] for e in expected if e['name']==r['name']) for r in rows)
assert db.execute('select count(*) from employees where category=?',('Teaching',)).fetchone()[0]==30
assert db.execute('select count(*) from employees where category=?',('Admin',)).fetchone()[0]==20
assert all(r['bank']=='Demo Bank' and r['account_no'].startswith('TEST-ACCT-') for r in db.execute('select bank, account_no from employees'))
again=core.import_workbook(db,p)
assert again['added']==0 and again['updated']==50
print(result)
print('Verified 50 records, all net salaries, bank details, 30 Teaching / 20 Admin, and duplicate-free reimport.')
print('Net total:',sum(r['net'] for r in rows))
print('First appointment date:',load_workbook(p,data_only=True).worksheets[0]['B6'].value)
