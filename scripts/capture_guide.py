"""Capture real application screens in an isolated training workspace."""
import json
from pathlib import Path
import socket
import sys
import threading
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import uvicorn
from PySide6.QtCore import QTimer, Qt, QRect
from PySide6.QtWidgets import QApplication, QFileDialog
from desktop import Window, ROOT
from payroll_app.service import Store
from payroll_app.api import create_app
from payroll_app import core

out = ROOT / 'tmp' / 'pdfs' / 'guide'
out.mkdir(parents=True, exist_ok=True)
store = Store(out / 'training-data')
with store.db() as db:
    # This database belongs only to the booklet capture script.
    for table in ('employees','monthly','locked','bps_scale','settings'):
        db.execute(f'DELETE FROM {table}')
    db.execute("DELETE FROM sqlite_sequence WHERE name='employees'")
    names=['Amina Example','Bilal Example','Dania Example','Faisal Example','Hina Example','Omar Example','Sara Example','Zain Example']
    for i,name in enumerate(names,1):
        db.execute('INSERT INTO employees(name,dt_appt,bps,basic,special,designation,section,account_no,bank,email,cpf_no,category) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
                   (name,'01.08.2024',17,30000+(i-1)*2000,2000,'Lecturer','Senior College',f'DEMO-{i:03}','Sample Bank',f'teacher{i}@example.invalid',f'CPF-{i:03}','Teaching'))
    db.execute("INSERT INTO employees(name,basic,active) VALUES ('Inactive Example',25000,0)")
    db.execute('INSERT INTO bps_scale VALUES (17,5000,4000,1000,3000,3000)')
    db.execute('INSERT INTO bps_scale VALUES (16,4000,3000,800,2500,2500)')
    db.commit()
    core.save_settings(db,dict(core.DEFAULTS,college_name='Sample College - Training',slip_header='Sample College | Training copy',adhoc_label='Adhoc allowance (fixed amount)',signatory_name='Accounts Officer',signatory_title='Finance Office',smtp_server='smtp.example.invalid',smtp_username='payroll@example.invalid',smtp_from='payroll@example.invalid'))
    rows,_=core.month_rows(db,'2026-08')
    for row in rows:
        core.save_monthly_employee(db,'2026-08',row)
    core.lock_month(db,'2026-08',rows)
    rows,_=core.month_rows(db,'2026-09')
    row=rows[0];row.update(worked=28,arrear=2000,cp_loan=1000,other=500)
    core.save_monthly_employee(db,'2026-09',row)

secret='guide-capture-only'
sock=socket.socket();sock.bind(('127.0.0.1',0));port=sock.getsockname()[1]
server=uvicorn.Server(uvicorn.Config(create_app(store,secret,ROOT/'frontend'/'out'),log_config=None,access_log=False))
thread=threading.Thread(target=server.run,kwargs={'sockets':[sock]},daemon=True);thread.start()
for _ in range(100):
    if server.started: break
    time.sleep(.05)
app=QApplication([])
app.setApplicationName('PayrollGuideCapture')
window=Window(f'http://127.0.0.1:{port}',secret,store)
window.setWindowTitle('Payroll Desk - sample screenshots')
window.resize(1360,1040)
window.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen,True)
window.show()
manifest={}
failures=[]
QFileDialog.getSaveFileName=lambda *args,**kwargs:(str(out/'Sample_Payslips.pdf'),'PDF document (*.pdf)')

def js(code,callback=None):
    window.view.page().runJavaScript(code,callback or (lambda _:None))

def button(text):
    return "[...document.querySelectorAll('button')].find(b=>b.textContent.trim()==="+json.dumps(text)+")"

def click(text): return button(text)+"?.click()"

def field(label):
    return "[...document.querySelectorAll('label.field')].find(x=>x.querySelector('span')?.textContent==="+json.dumps(label)+")"

def capture(name,action='true',targets=None,crop=None,delay=550):
    def measure():
        expressions={str(i):e for i,e in enumerate(targets or [],1)}
        code="(()=>{const r={};"+''.join(f"{{const e=({expr});if(e){{const b=e.getBoundingClientRect();r[{json.dumps(key)}]={{x:b.x+b.width/2,y:b.y+b.height/2,w:b.width,h:b.height}}}}}}" for key,expr in expressions.items())+"return JSON.stringify({points:r,width:innerWidth,height:innerHeight,text:document.body.innerText});})()"
        js(code,save)
    def save(raw):
        data=json.loads(raw)
        rect=crop or (0,0,data['width'],data['height'])
        window.view.grab(QRect(*map(int,rect))).save(str(out/(name+'.png')))
        data['crop']=rect
        for value in data['points'].values():
            value['x']-=rect[0];value['y']-=rect[1]
        if len(data['points'])!=len(targets or []): failures.append(name)
        manifest[name]=data
        next_step()
    js(action)
    QTimer.singleShot(delay,measure)

S=lambda selector:f"document.querySelector({json.dumps(selector)})"
steps=[
 lambda:capture('dashboard',"document.documentElement.dataset.theme='light';window.confirm=()=>true;window.scrollTo(0,0)",[
     S('nav'),S('.month-picker'),S('.stat-grid'),button('Open payroll'),S('.overview-row:nth-child(4)'),button('Dark appearance')]),
 lambda:capture('dashboard-actions','window.scrollTo(0,250)',[S('.quick-action:nth-of-type(1)'),S('.quick-action:nth-of-type(2)'),S('.quick-action:nth-of-type(3)'),S('.periods')],crop=(240,220,1100,770)),
 lambda:capture('employees',click('Employees'),[S('.search'),S('.checkbox-label'),button('Import Excel'),button('Add employee'),S('.employee-link')]),
 lambda:capture('employee-top',S('.employee-link')+'.click()',[
     field('Full name'),field('Designation'),field('Appointment date'),field('BPS (blank for contract)'),field('Section'),field('Category / group')],crop=(770,0,590,1040)),
 lambda:capture('employee-bottom',S('.drawer-content')+'.scrollTop=9999',[
     field('Bank account number'),field('Email address'),field('CPF member'),field('Employee status'),field('Running basic'),button('Save changes')],crop=(770,0,590,1040)),
 lambda:capture('payroll',click('Close')+';'+click('Payroll'),[
     S('.month-picker'),S('.search'),S('.table-subbar select'),button('Export Excel'),button('Save PDF'),button('Review & lock month')]),
 lambda:capture('payroll-table','window.scrollTo(0,260)',[
     S('thead input'),S('tbody .employee-link'),S('tbody .days'),S('tbody .deduction'),S('tbody .net'),button('Refresh')],crop=(240,215,1100,730)),
 lambda:capture('monthly-top',S('tbody .employee-link')+'.click()',[
     S('.form-note'),field('Bank account number'),field('Bank'),field('Email address'),field('CPF member')],crop=(770,0,590,1040)),
 lambda:capture('monthly-inputs',S('.drawer-content')+'.scrollTop=510',[
     field('Working days'),field('Days worked'),field('Running basic'),field('Adhoc allowance'),field('CPF deduction'),field('Other deduction')],crop=(770,0,590,1040)),
 lambda:capture('monthly-preview',S('.drawer-content')+'.scrollTop=9999',[
     field('Arrears / encashment'),field('Loan instalment'),S('.salary-preview'),button('Save changes')],crop=(770,0,590,1040)),
 lambda:capture('payroll-locked',click('Close')+';'+click('Review & lock month'),[
     S('.toolbar-actions .badge'),button('Unlock month'),button('Save PDF')],delay=950),
 lambda:capture('reports',click('Reports'),[
     S('.report-banner select'),S('.report-card:nth-child(1) button'),S('.report-card:nth-child(2) button'),S('.report-card:nth-child(3) button'),S('.report-card:nth-child(4) button'),S('.report-card:nth-child(5) button')]),
 lambda:capture('report-email','window.scrollTo(0,230)',[
     S('.report-card:nth-child(6) button'),S('.hint:last-of-type')],crop=(240,360,1100,620)),
 lambda:capture('export-success',"window.scrollTo(0,0);"+S('.report-card:nth-child(1) button')+'.click()',[
     S('.notice'),button('Open saved file'),S('.report-banner select')],delay=1800),
 lambda:capture('settings-general',click('Settings'),[
     button('Set logo'),field('College name (shown on the salary sheet)'),field('Working days in a month (default)'),field("CPF: college contribution, % of employee's deduction"),button('Save settings')]),
 lambda:capture('bps',click('BPS scales'),[button('Add grade'),button('Edit grade'),S('thead')]),
 lambda:capture('bps-editor',button('Edit grade')+'.click()',[
     field('BPS grade'),field('House rent'),field('Adhoc allowance'),field('CPF deduction'),button('Delete grade'),button('Save changes')],crop=(770,0,590,1040)),
 lambda:capture('email-settings',click('Close')+';'+click('Email'),[
     field('Email SMTP server (e.g. smtp.gmail.com)'),field('Email SMTP port (usually 587)'),field('Email username / login'),field('Email password / app password'),field('From email address (leave blank to use username)'),button('Save settings')]),
 lambda:capture('backup',click('Backup & restore'),[button('Create backup'),button('Restore backup'),S('.data-location')]),
 lambda:capture('dark',click('Dashboard')+';'+click('Dark appearance'),[button('Light appearance')]),
]

def next_step():
    if steps:
        QTimer.singleShot(180,steps.pop(0))
    else:
        (out/'manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8')
        print(json.dumps({'captured':len(manifest),'missing_targets':failures}))
        app.quit()

QTimer.singleShot(4000,next_step)
QTimer.singleShot(90000,app.quit)
app.exec()
server.should_exit=True;thread.join(timeout=5)
sys.exit(1 if failures or steps else 0)
