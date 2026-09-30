"""Exercise the actual Qt/Next.js window against isolated synthetic data."""
import json
from pathlib import Path
import socket
import sys
import threading
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import uvicorn
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QFileDialog
from desktop import Window, ROOT
from payroll_app.service import Store
from payroll_app.api import create_app

out = ROOT / 'artifacts' / 'ui-verification'
out.mkdir(parents=True, exist_ok=True)
store = Store(out / 'data')
with store.db() as db:
    if not db.execute('SELECT 1 FROM employees').fetchone():
        db.executemany('INSERT INTO employees(name,designation,bps,basic,account_no,bank) VALUES (?,?,?,?,?,?)',
                       [(f'Teacher {i:03}', 'Lecturer', 17, 45000+i, f'00123456{i:04}', 'ABL') for i in range(400)])
        db.execute('INSERT INTO bps_scale VALUES (17,5000,3000,2000,4500,2500)')
        db.commit()
secret='isolated-ui-test-secret'
sock=socket.socket();sock.bind(('127.0.0.1',0));port=sock.getsockname()[1]
server=uvicorn.Server(uvicorn.Config(create_app(store,secret,ROOT/'frontend'/'out'),log_config=None,access_log=False))
thread=threading.Thread(target=server.run,kwargs={'sockets':[sock]},daemon=True);thread.start()
for _ in range(100):
    if server.started: break
    time.sleep(.05)
app=QApplication([])
window=Window(f'http://127.0.0.1:{port}',secret,store)
window.show()
results=[]
failed=[]

def js(code, callback=None):
    window.view.page().runJavaScript(code,callback or (lambda _:None))

def click(text):
    return "[...document.querySelectorAll('button')].find(b=>b.textContent.trim()==="+json.dumps(text)+")?.click()"

def snapshot(name, code='true', condition='true'):
    def check(result):
        if not result: failed.append(name)
        window.grab().save(str(out/(name+'.png')))
        results.append({'check':name,'passed':bool(result)})
        next_step()
    js(code)
    QTimer.singleShot(650,lambda:js(condition,check))

steps=[
 lambda:snapshot('dashboard-light',"document.documentElement.dataset.theme='light';window.scrollTo(0,0)","document.body.innerText.includes('400 employees') && !!window.desktop"),
 lambda:snapshot('payroll',click('Payroll'),"document.querySelectorAll('tbody tr').length===400"),
 lambda:snapshot('search',"(()=>{const i=document.querySelector('input[aria-label=\"Search employees\"]');Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set.call(i,'Teacher 015');i.dispatchEvent(new Event('input',{bubbles:true}));})()","document.querySelectorAll('tbody tr').length===1"),
 lambda:snapshot('employee-editor',"document.querySelector('.employee-link').click()","!!document.querySelector('[role=dialog]') && document.body.innerText.includes('Estimated net pay')"),
 lambda:snapshot('editor-closed',click('Close'),"!document.querySelector('[role=dialog]')"),
 lambda:snapshot('reports',click('Reports'),"document.querySelectorAll('.report-card').length===6"),
 lambda:snapshot('reports-dark',click('Dark appearance'),"document.documentElement.dataset.theme==='dark'"),
 lambda:snapshot('settings',click('Settings'),"document.querySelectorAll('.settings-panel input').length>=8"),
 lambda:snapshot('scales',click('BPS scales'),"document.body.innerText.includes('Fixed allowances by grade')"),
 lambda:snapshot('native-save',"window.__nativeSaved=false;window.desktop.choose('backup','ui-backup.db',async raw=>{const t=JSON.parse(raw);const r=await fetch('/api/backup',{method:'POST',headers:{'Content-Type':'application/json','X-Payroll-Token':window.payrollToken},body:JSON.stringify({ticket:t.ticket})});window.__nativeSaved=r.ok;});","window.__nativeSaved===true"),
]

# Keep the native bridge and API intact; avoid waiting on a human in this test.
QFileDialog.getSaveFileName=lambda *args,**kwargs:(str(out/'native-backup.db'),'Payroll database (*.db)')

def next_step():
    if steps:
        QTimer.singleShot(250,steps.pop(0))
    else:
        (out/'results.json').write_text(json.dumps(results,indent=2),encoding='utf-8')
        print(json.dumps({'checks':results,'backup_exists':(out/'native-backup.db').exists()}))
        app.quit()

QTimer.singleShot(5000,next_step)
QTimer.singleShot(60000,app.quit)
app.exec()
server.should_exit=True;thread.join(timeout=5)
sys.exit(1 if failed or steps else 0)
