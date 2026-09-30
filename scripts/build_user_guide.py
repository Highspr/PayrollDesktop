"""Build the illustrated user booklet from real, isolated app screenshots."""
from pathlib import Path
import json
import math
import os
from xml.sax.saxutils import escape

from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph

ROOT=Path(__file__).resolve().parents[1]
SRC=ROOT/'tmp/pdfs/guide'
OUT=ROOT/'output/pdf'
OUT.mkdir(parents=True,exist_ok=True)
PDF=OUT/'Payroll_Desk_Visual_User_Guide.pdf'
M=json.loads((SRC/'manifest.json').read_text(encoding='utf-8'))
FONTDIR=Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'
for name,file in [('Guide','segoeui.ttf'),('GuideBold','segoeuib.ttf')]:
    pdfmetrics.registerFont(TTFont(name,str(FONTDIR/file)))
W,H=landscape(A4)
NAVY=colors.HexColor('#142B46');TEAL=colors.HexColor('#087F75');MUTED=colors.HexColor('#5F7085')
LINE=colors.HexColor('#DFE7EF');PALE=colors.HexColor('#F3F7FB');MINT=colors.HexColor('#EAF6F2');GOLD=colors.HexColor('#D77918')
C=canvas.Canvas(str(PDF),pagesize=(W,H),pageCompression=1)
C.setTitle('Payroll Desk - Visual User Guide')
C.setAuthor('Payroll Desk')
C.setSubject('Simple English instructions with annotated application screenshots')
PAGE=0
pages=[]

def text(x,y,s,size=10,font='Guide',color=NAVY):
    C.setFont(font,size);C.setFillColor(color);C.drawString(x,y,s)

def para(x,top,width,s,size=10,leading=None,color=NAVY,bold=False,max_height=None):
    style=ParagraphStyle('p',fontName='GuideBold' if bold else 'Guide',fontSize=size,leading=leading or size*1.4,textColor=color,spaceAfter=0)
    p=Paragraph(s,style);_,height=p.wrap(width,1000)
    if max_height and height>max_height:
        raise ValueError(f'Paragraph exceeds available area: {s[:70]} ({height} > {max_height})')
    p.drawOn(C,x,top-height)
    return height

def number(x,y,n,r=9):
    C.setFillColor(colors.white);C.circle(x,y,r+1.4,fill=1,stroke=0)
    C.setFillColor(NAVY);C.circle(x,y,r,fill=1,stroke=0)
    C.setFillColor(colors.white);C.setFont('GuideBold',9);C.drawCentredString(x,y-3.1,str(n))

def start(title,section,subtitle='',bookmark=None):
    global PAGE
    PAGE+=1;pages.append(title)
    C.setFillColor(colors.white);C.rect(0,0,W,H,fill=1,stroke=0)
    C.setFillColor(TEAL);C.rect(0,H-7,W,7,fill=1,stroke=0)
    key=bookmark or f'page{PAGE}'
    C.bookmarkPage(key);C.addOutlineEntry(title,key,0,False)
    text(32,H-32,section.upper(),8,'GuideBold',TEAL)
    text(32,H-64,title,25,'GuideBold')
    if subtitle: para(32,H-76,W-64,subtitle,10,14,MUTED,max_height=32)

def finish():
    C.setStrokeColor(LINE);C.setLineWidth(.6);C.line(32,33,W-32,33)
    text(32,20,'PAYROLL DESK  /  VISUAL GUIDE',7,'GuideBold',MUTED)
    text(290,20,'Screenshots use sample data.',7,'Guide',MUTED)
    if PAGE>2:
        text(W-153,20,'Back to contents',7,'Guide',TEAL)
        C.linkRect('', 'contents', (W-157,15,W-60,30),relative=0,thickness=0)
    text(W-46,20,f'{PAGE:02}',8,'GuideBold')
    C.showPage()

def shot(name,box,crop=None,arrows=None):
    """Place an unmodified screenshot; annotations are PDF vector objects."""
    image=SRC/(name+'.png')
    iw,ih=Image.open(image).size
    rx,ry,rw,rh=crop or (0,0,iw,ih)
    bx,by,bw,bh=box
    scale=min(bw/rw,bh/rh)
    dw,dh=rw*scale,rh*scale
    x=bx;y=by+bh-dh
    C.setFillColor(PALE);C.roundRect(x-2,y-2,dw+4,dh+4,6,fill=1,stroke=0)
    C.saveState();path=C.beginPath();path.rect(x,y,dw,dh);C.clipPath(path,stroke=0)
    C.drawImage(str(image),x-rx*scale,y-(ih-ry-rh)*scale,width=iw*scale,height=ih*scale)
    C.restoreState()
    C.setStrokeColor(LINE);C.setLineWidth(.65);C.rect(x,y,dw,dh,fill=0,stroke=1)
    positions=[]
    for item in arrows or []:
        n,key,*custom=item
        p=M[name]['points'][str(key)] if isinstance(key,(str,int)) else {'x':key[0],'y':key[1]}
        tx=x+(p['x']-rx)*scale;ty=y+dh-(p['y']-ry)*scale
        if not(x-1<=tx<=x+dw+1 and y-1<=ty<=y+dh+1):
            raise ValueError(f'Arrow {name}/{key} outside screenshot crop')
        offsets=[(-32,22),(32,22),(-32,-22),(32,-22),(0,31),(0,-31)]
        if custom: offsets=[custom[0]]+offsets
        for dx,dy in offsets:
            mx=max(x+11,min(x+dw-11,tx+dx));my=max(y+11,min(y+dh-11,ty+dy))
            if all(math.hypot(mx-a,my-b)>25 for a,b in positions):break
        positions.append((mx,my))
        angle=math.atan2(ty-my,tx-mx)
        sx=mx+10*math.cos(angle);sy=my+10*math.sin(angle)
        C.setStrokeColor(colors.white);C.setLineWidth(4);C.line(sx,sy,tx,ty)
        C.setStrokeColor(GOLD);C.setLineWidth(1.7);C.line(sx,sy,tx,ty)
        head=C.beginPath();head.moveTo(tx,ty)
        head.lineTo(tx-7*math.cos(angle-.42),ty-7*math.sin(angle-.42))
        head.lineTo(tx-7*math.cos(angle+.42),ty-7*math.sin(angle+.42));head.close()
        C.setFillColor(GOLD);C.drawPath(head,fill=1,stroke=0)
        number(mx,my,n)
    return (x,y,dw,dh)

def keys(items,top=174,columns=2,x=32,width=None,row_height=41,size=9.3):
    width=width or W-64
    colw=width/columns
    for i,(title,body) in enumerate(items):
        row=i//columns;col=i%columns
        xx=x+col*colw;yy=top-row*row_height
        number(xx+9,yy-9,i+1,8)
        para(xx+25,yy,colw-41,f'<b>{escape(title)}</b> - {escape(body)}',size,size*1.3,max_height=row_height-3)

def note(s,y=47,color=MINT):
    C.setFillColor(color);C.roundRect(32,y,W-64,34,6,fill=1,stroke=0)
    para(44,y+24,W-88,s,9,12,max_height=26)

def wide(title,section,subtitle,image,crop,arrows,items,tip=None):
    start(title,section,subtitle)
    shot(image,(32,204,W-64,286),crop,arrows)
    keys(items,top=186,row_height=42)
    if tip:note(tip,y=43)
    finish()

def detail(title,section,subtitle,image,crop,arrows,items,tip=None,extra=None):
    start(title,section,subtitle)
    shot(image,(40,94,365,400),crop,arrows)
    if extra:
        im,box,cr,ar=extra;shot(im,box,cr,ar)
    keys(items,top=485,columns=1,x=435,width=W-465,row_height=64,size=10)
    if tip:note(tip,y=44)
    finish()

def card(x,y,w,h,title,body,n=None):
    C.setFillColor(PALE);C.roundRect(x,y,w,h,9,fill=1,stroke=0)
    if n is not None:
        number(x+21,y+h-24,n);offset=42
    else:offset=16
    para(x+offset,y+h-15,w-offset-14,escape(title),13,17,bold=True,max_height=37)
    para(x+16,y+h-60,w-32,body,10,14,color=MUTED,max_height=h-70)

# 01 - Cover.
PAGE+=1;pages.append('Cover')
C.bookmarkPage('cover');C.addOutlineEntry('Cover','cover',0,False)
C.setFillColor(NAVY);C.rect(0,0,W,H,fill=1,stroke=0)
C.setFillColor(TEAL);C.rect(32,H-57,47,6,fill=1,stroke=0)
text(32,H-88,'PAYROLL DESK',12,'GuideBold',colors.HexColor('#79D5C2'))
text(32,H-139,'Your visual',39,'GuideBold',colors.white)
text(32,H-187,'user guide.',39,'GuideBold',colors.white)
para(34,H-211,340,'Simple English. Real screenshots.<br/>Numbered arrows for every main task.',16,23,colors.HexColor('#C3D2E2'))
text(34,69,'DESKTOP EDITION 1.0  |  SEPTEMBER 2026',9,'GuideBold',colors.HexColor('#79D5C2'))
text(34,46,'Open it. Find the button. Follow the steps.',11,'Guide',colors.white)
shot('dashboard',(407,80,404,449),(225,90,1110,930))
C.showPage()

# 02 - Contents.
start('Find what you need','Start here','Click a topic below to jump to it. The same page numbers work in a printed copy.','contents')
topics=[('First launch and the monthly routine',3),('Dashboard: navigation and totals',4),('Dashboard: quick actions',5),('Employees: search, add and edit',6),('Employee master: personal details',7),('Employee master: bank and salary',8),('Import an Excel salary workbook',9),('Payroll: month and main buttons',10),('Read the salary table',11),('Edit one month only',12),('Working days and salary inputs',13),('Understand the salary calculation',14),('Lock and unlock a month',15),('Download the right report',16),('Save, open and print a file',17),('General settings and college logo',18),('BPS scales and allowances',19),('Email setup and sending payslips',20),('Backup, restore and another PC',21),('Light and dark appearance',22),('Salary words in plain English',23),('Monthly checklist and quick fixes',24)]
for i,(label,page) in enumerate(topics):
    col=i//11;row=i%11;x=32+col*398;y=475-row*33
    text(x,y,f'{page:02}',10,'GuideBold',TEAL)
    text(x+32,y,label,10)
    C.setStrokeColor(LINE);C.line(x,y-11,x+365,y-11)
    C.linkRect('',f'page{page}',(x,y-10,x+366,y+13),relative=0,thickness=0)
note('<b>How to read an arrow:</b> match its number to the explanation below or beside the screenshot. All people and account details shown are examples.')
finish()

# 03 - First launch.
start('Start with a simple routine','Getting started','Use the desktop shortcut. You do not need to open a browser or type commands.')
card(32,302,246,186,'Open Payroll Desk','After installation, double-click <b>Payroll Desk</b> on the desktop.<br/><br/>In this project folder, <b>Run Payroll.cmd</b> also opens the new application.',1)
card(298,302,246,186,'Check the month','The app normally opens the current month. Use <b>Payroll period</b> to choose the month you need.<br/><br/>An older saved month is not opened automatically.',2)
card(564,302,246,186,'Prepare the records','Check employee details, BPS scales and college settings. Add staff manually or import the college salary workbook.',3)
card(32,96,246,186,'Review the payroll','Enter this month\'s days, arrears and deductions. Click <b>Save changes</b>, then review gross, deductions and net pay.',4)
card(298,96,246,186,'Lock and download','When the figures are correct, lock the month. Save the payslips and the statements you need.',5)
card(564,96,246,186,'Make a backup','Save a database backup to another drive. Keep the backup private.<br/><br/>Core payroll works offline. Only email needs internet.',6)
finish()

# 04 - Dashboard top.
wide('Dashboard: your starting point','Dashboard','This screen gives a summary of the selected month. It does not pay money to anyone.',
 'dashboard',(0,75,1360,500),[(1,1), (2,2),(3,3),(4,4)],
 [('Navigation menu','Open Dashboard, Employees, Payroll, Reports or Settings.'),('Payroll period','Choose the month and year before working on salary.'),('Salary cards','Gross is before deductions. Deductions reduce pay. Net payable is the amount left.'),('Open payroll','Go to the detailed salary table for this month.')],
 'The totals describe the selected month. They are not a total of every month in the database.')

# 05 - Dashboard lower cards.
wide('Dashboard: useful shortcuts','Dashboard','Use these shortcuts for common jobs. The main pages offer the same features.',
 'dashboard-actions',(0,355,1100,390),[(1,1), (2,2),(3,3),(4,4)],
 [('Download payslips','Opens Reports. Choose the scope, then click Save PDF.'),('Import salary workbook','Choose the supported college Excel workbook. See page 9 before importing.'),('Back up your records','Choose a folder and save a copy of the database.'),('Recent periods','Click a saved month to load it. The summary updates for that period.')],
 '<b>Month at a glance:</b> shows payroll staff, active master records, bank/net-pay issues and the number of saved periods. A review count is a prompt to check the records.')

# 06 - Employee list.
wide('Employees: find, add and edit staff','Employees','These are master records: the usual details and pay defaults for each person.',
 'employees',(235,220,1100,400),[(1,1),(2,2),(3,3),(4,4),(5,5)],
 [('Search','Type a name, designation or BPS grade to narrow the list.'),('Show inactive','Tick this to include people marked inactive.'),('Import Excel','Load staff and salary inputs from the supported workbook.'),('Add employee','Open a blank form for a new staff member.'),('Open a record','Click the name or the arrow at the right. Edit the form, then save.')],
 'Click supported column headings, such as Employee or Basic Pay, to change the sort order. Click again to reverse it.')

# 07 - Employee personal fields.
detail('Employee master: personal details','Employees','Add a new employee or open an existing employee from the list.',
 'employee-top',(0,100,590,315),[(1,1),(2,2),(3,3),(4,4),(5,5),(6,6)],
 [('Full name','Enter the name to show on salary sheets and payslips. This field is required.'),('Designation','Enter the job title, for example Lecturer or Accountant.'),('Appointment date','Enter the date in your college\'s usual format, such as 01.08.2024.'),('BPS','Enter the pay grade. Leave it blank for a person without a BPS grade.'),('Section','Enter the section or department used on the payslip.'),('Category / group','Use a staff group, such as Teaching or Admin, for your records.')],
 'Changes in Employees are master changes. Unlocked months may use these changes if no separate monthly value is saved.')

# 08 - Employee bank and pay.
start('Employee master: bank and default pay','Employees','Keep account details accurate. Amounts here are the usual defaults for payroll.')
shot('employee-bottom',(40,176,365,318),(0,425,590,425),[(1,1),(2,2),(3,3),(4,4),(5,5)])
shot('employee-bottom',(40,100,365,62),(0,955,590,85),[(6,6)])
keys([('Bank account and bank','Enter both the account number and bank name. Keep any leading zeros.'),('Email and CPF account','Enter the employee\'s email for payslips. CPF account number is used on the CPF statement.'),('CPF member','Choose Yes to apply the CPF deduction; choose No for a non-member.'),('Employee status','Inactive people are excluded from unlocked payrolls. To reactivate, show inactive staff and choose Active.'),('Default salary','Running basic is the base salary. Special allowance is the extra amount for this person.'),('Save changes','Saves the record. Close or Discard leaves the form; confirm if asked about unsaved edits.')],top=485,columns=1,x=435,width=W-465,row_height=64,size=10)
note('Deactivation does not remove a person from locked historical payrolls. It can change an unlocked past month, so finalize history before changing staff status.')
finish()

# 09 - Import.
start('Import a salary workbook','Employees / Import Excel','Use the college salary workbook layout supported by the application.')
shot('employees',(32,336,W-64,145),(250,225,1060,150),[(1,3)])
card(32,117,246,199,'Choose the workbook','Click <b>Import Excel</b>, read the confirmation and choose an <b>.xlsx</b> file.<br/><br/>The salary title must include <b>Calculation of Staff Salary</b> and a readable month/year, such as <b>Month of August 2026</b>.',1)
card(298,117,246,199,'Understand the update','The importer reads employees, BPS amounts and one month\'s salary inputs.<br/><br/>People are matched by name and appointment date. Matching records are updated. A safety backup is made before the live data is replaced.',2)
card(564,117,246,199,'Read the result','Check the new/updated counts, missing bank matches, similar-name matches, allowance differences and other deductions.<br/><br/>Open the imported month and review the affected staff before locking it.',3)
note('An import into a locked month is rejected. The staged import is not applied. Unlock only if you intentionally want to change that month.')
finish()

# 10 - Payroll toolbar.
wide('Payroll: month and main buttons','Payroll','Start here for monthly salary work. Check the period before editing or exporting.',
 'payroll',(250,100,1080,455),[(1,1),(2,2),(3,3),(4,4),(5,5),(6,6)],
 [('Payroll period','Choose the month and year.'),('Search','Find a person by name, designation or BPS.'),('Export scope','Choose Entire month or Selected employees.'),('Export Excel','Save the salary register as an Excel file.'),('Save PDF','Save payslips for the chosen scope.'),('Review & lock month','Finalize the figures after you have reviewed them. See page 15.')])

# 11 - Salary table.
start('Read the salary table','Payroll','Amounts are shown in Pakistani rupees (PKR). Scroll sideways if all columns do not fit.')
shot('payroll-table',(32,319,W-64,175),(0,105,1100,225),[(1,1),(2,2),(3,3),(4,4),(5,5)])
shot('payroll-table',(32,265,W-64,39),(0,690,1100,40),[(6,6)])
keys([('Checkboxes','Select individual staff. The top checkbox selects the rows currently visible in the list.'),('Employee name','Click the name or row arrow to open that person\'s monthly form.'),('Days','The first number is days worked; the second is working days for the month.'),('Deductions','The total includes CPF, loan instalment, unpaid-day deduction and other deduction.'),('Net pay','The amount remaining after deductions. Red negative pay needs review.'),('Refresh','Reload saved payroll data. Gross, allowances and arrears columns show the earnings breakdown.')],top=242,row_height=54,size=10)
note('Searching does not automatically limit an export. Tick the people you need, then explicitly choose Selected employees.')
finish()

# 12 - Monthly identity.
detail('Edit one month without changing the master','Payroll / Employee details','Click a person in Payroll. The form heading identifies the selected month.',
 'monthly-top',(0,100,565,680),[(1,1),(2,2),(3,3),(4,4),(5,5)],
 [('Monthly-only notice','Edits here are saved for this month. The Employees master record stays unchanged.'),('Bank account number','A correction here applies to this month\'s reports only.'),('Bank','Check the bank name together with the account number.'),('Email address','This month\'s saved email is used when its payslip is sent.'),('CPF member','Choose whether the CPF deduction applies in this month.')],
 'You can also edit name, designation, appointment date, BPS, section and CPF account here. Scroll down for salary inputs. Locked months open read-only.')

# 13 - Attendance and salary.
detail('Enter days, earnings and deductions','Payroll / Monthly inputs','Scroll inside the employee form to reach Attendance & salary inputs.',
 'monthly-inputs',(0,280,565,510),[(1,1),(2,2),(3,3),(4,4),(5,5),(6,6)],
 [('Working days','Enter the payroll basis for the month: 1 to 31 days.'),('Days worked','Enter 0 up to the working-days number. Fewer days produce an unpaid-day deduction.'),('Basic and allowances','Review running basic, house rent, conveyance, medical and special allowance.'),('Adhoc allowance','Enter a fixed amount. A label mentioning a percentage does not calculate that percentage.'),('CPF deduction','Enter the employee\'s fund deduction. It is applied only when CPF member is Yes.'),('Other deduction','Enter an additional deduction. Arrears add pay; loan instalment reduces pay.')],
 'New months use master/BPS defaults. Last month\'s arrears and loan instalment do not automatically repeat.')

# 14 - Calculation example.
start('Understand the net-pay preview','Payroll / Salary example','This example matches Amina Example in the training screenshots. It is an illustration of the app\'s calculation.')
shot('monthly-preview',(40,97,320,397),(0,505,565,535),[(1,1),(2,2),(3,3),(4,4)])
rows=[('Basic + allowances','30,000 + 15,000 = 45,000'),('Gross salary','45,000 + 2,000 arrears = 47,000'),('Unpaid days','30 working days - 28 worked = 2'),('Leave deduction','45,000 / 30 x 2 = 3,000'),('Total deductions','3,000 CPF + 1,000 loan + 3,000 leave + 500 other = 7,500'),('Net pay','47,000 - 7,500 = 39,500')]
for i,(title,body) in enumerate(rows):
    yy=483-i*49
    para(392,yy,408,escape(title),11,15,bold=True)
    para(392,yy-19,408,escape(body),10,14,color=MUTED,max_height=29)
keys([('Arrears','Extra pay added this month.'),('Loan instalment','This month\'s loan deduction.'),('Preview','Updates as you change inputs.'),('Save changes','Applies the form to payroll.')],top=169,columns=2,x=392,width=408,row_height=42,size=9)
note('Arrears are not included in the unpaid-day deduction basis. The app uses whole rupees. A preview is not saved until you click Save changes.')
finish()

# 15 - Locking.
start('Lock the final figures','Payroll / Lock and unlock','Use locking after you have checked the month, staff list and salary totals.')
shot('payroll-locked',(32,323,W-64,166),(250,200,1080,180),[(1,1),(2,2),(3,3)])
card(32,104,246,194,'Check the status','<b>Draft</b> means payroll can change.<br/><br/><b>Locked</b> means the saved salary snapshot is used. Monthly details become read-only and later master pay changes do not change these locked figures.',1)
card(298,104,246,194,'Unlock only for a correction','Click <b>Unlock month</b> and read the confirmation.<br/><br/>Unlocking removes the final snapshot and recalculates with current master data and monthly inputs. Recheck the amounts before locking again.',2)
card(564,104,246,194,'Use the final reports','Click <b>Save PDF</b> or open Reports.<br/><br/>Locking does not send email, transfer money or create a backup. It preserves the payroll figures. College headings still use current Settings.',3)
note('An empty payroll or a negative net salary cannot be finalized. Fix invalid days and amounts, save the changes, then try again.')
finish()

# 16 - Report gallery.
wide('Choose the right report','Reports','Select the month and export scope first. Each card creates a different file.',
 'reports',(260,467,1050,553),[(1,2),(2,3),(3,4),(4,5),(5,6),(6,(1143,973))],
 [('Employee payslips - PDF','One A4 page per employee, combined in one PDF.'),('Salary register - Excel','Detailed earnings, deductions and totals.'),('Bank transfer statement - Excel','Names, bank accounts, banks and net salaries.'),('CPF statement - Excel','Employee contribution, college contribution and loan instalment.'),('Bank details document - Word','An editable bank-details document.'),('Deliver payslips - Email','Sends individual PDFs. Needs internet, email settings and a locked month.')],
 'A bank statement is a document only. Payroll Desk does not transfer money or connect to your bank account.')

# 17 - Saving.
wide('Save a file, then open or print it','Reports / File downloads','The same save workflow applies to PDF, Excel and Word reports.',
 'export-success',(250,210,1080,275),[(1,1),(2,2),(3,3)],
 [('Saved message','After the report finishes, check the folder and file name shown in the green message.'),('Open saved file','Open the report in the default Windows viewer. To print a PDF, use that viewer\'s Print command.'),('Report scope','Entire month includes everyone. Selected employees uses the checkboxes on Payroll.')],
 'Click a Save button, choose a folder in the Windows save box, enter a file name and click Save. Cancel closes the box without creating a report.')

# 18 - Settings.
wide('Set up your college and payslips','Settings / General','Changes are stored when you click Save settings. Set logo accepts a PNG image.',
 'settings-general',(270,390,1010,625),[(1,1),(2,2),(3,3),(4,4),(5,5)],
 [('Set logo','Choose the college logo for future payslips.'),('College name and heading','Set the salary-sheet name and payslip heading. Add the signatory name and title below.'),('Working days','Set the default payroll days. Check the actual month before calculating salary.'),('College CPF contribution','This percentage is based on the employee\'s CPF deduction, not basic pay. 100% means an equal contribution.'),('Save settings','Save after editing. Default bank and voucher prefix are also set on this page.')],
 'The adhoc/increase field changes payslip wording only. The actual adhoc amount comes from BPS or the monthly salary inputs.')

# 19 - BPS.
start('Manage BPS allowance amounts','Settings / BPS scales','These are fixed amounts for a grade. Basic salary is entered separately for each employee.')
shot('bps',(40,408,365,86),(275,390,1020,205),[(1,1),(2,2)])
shot('bps-editor',(40,172,365,216),(0,105,590,330),[(3,2),(4,4)])
shot('bps-editor',(40,103,365,55),(0,955,590,85),[(5,5),(6,6)])
keys([('Add grade','Create a new grade and enter its monthly amounts.'),('Edit grade','Open an existing grade. Keep its BPS number unless you intend to add a different grade.'),('Allowance fields','Set house rent, conveyance, medical and adhoc. Use amounts, not percentage formulas.'),('CPF deduction','Set the employee deduction for members of this grade.'),('Delete grade','Removes that grade. People using its defaults may lose those allowances in unlocked months.'),('Save changes','Save the grade. Monthly overrides stay as entered; locked salary figures stay fixed.')],top=485,columns=1,x=435,width=W-465,row_height=64,size=10)
note('With no BPS, grade allowances are not added automatically. Running basic, special allowance and any explicitly entered monthly amounts can still apply.')
finish()

# 20 - Email.
wide('Set up email, then send payslips','Settings / Email','Use the details supplied by your email provider or IT team. Example addresses in this guide do not work.',
 'email-settings',(275,400,1010,445),[(1,1),(2,2),(3,3),(4,4),(5,5),(6,6)],
 [('SMTP server','The server used to send email. Enter your provider\'s value.'),('SMTP port','Use the provider\'s STARTTLS port, usually 587.'),('Username / login','The login name for the sending mailbox.'),('Password / app password','Enter the correct credential. A blank field keeps the saved password.'),('From email address','The sender address. Leave blank to use the username.'),('Save settings','Then lock the month, open Reports and choose Email all payslips.')],
 'Confirm the send count. Staff without an email are skipped. Email sends to the whole month, not the report selection. Review sent/failed results before retrying.')

# 21 - Backup.
wide('Back up, restore or move to another PC','Settings / Backup & restore','A database backup contains employees, settings, monthly inputs and locked payroll figures.',
 'backup',(275,390,1040,370),[(1,1),(2,2),(3,3)],
 [('Create backup','Choose another drive or safe folder and save the .db file. Do this regularly.'),('Restore backup','Replaces current data with the chosen backup. A safety copy is made before replacement.'),('Your local data folder','Shows where the application keeps its working database on this computer.')],
 'On another Windows PC: install Payroll Desk, open Settings, then Restore backup. Exported reports and logo.png are separate; copy reports and set the logo again.')

# 22 - Appearance.
start('Choose a comfortable appearance','Light and dark themes','The theme changes the screen colors. It does not change salaries or the report files.')
shot('dashboard',(32,238,377,247),(0,590,790,440),[(1,6)])
shot('dark',(432,238,377,247),(0,590,790,440),[(2,1)])
text(32,214,'LIGHT APPEARANCE',10,'GuideBold',TEAL)
text(432,214,'DARK APPEARANCE',10,'GuideBold',TEAL)
keys([('Dark appearance','Click the moon button at the bottom of the sidebar to switch to dark colors.'),('Light appearance','Click the sun button to return to light colors. Your choice is remembered on this computer.')],top=191,row_height=68,size=11)
note('If a section does not fit, scroll the page or the table. Employee forms have their own scroll area. Resize the window for a wider view.')
finish()

# 23 - Glossary.
start('Salary words in plain English','Quick reference','These meanings explain the fields used by this application.')
terms=[('Master record','An employee\'s usual details and pay defaults, edited in Employees.'),('Monthly record','Details or amounts saved for one selected month.'),('BPS','The grade used to look up fixed allowance and CPF amounts.'),('Running basic','The employee\'s base salary before allowances.'),('Allowance','Extra pay, such as house rent, conveyance, medical or special allowance.'),('Adhoc allowance','A fixed additional amount. The wording does not create a percentage formula.'),('Arrears / encashment','An extra amount added to gross pay for the month.'),('Gross salary','Basic + allowances + arrears, before deductions.'),('CPF / CP Fund','The employee\'s contributory provident fund deduction.'),('College contribution','The college\'s additional CPF amount. It is not deducted again from net salary.'),('Loan instalment','The repayment amount deducted this month; this app does not track a loan balance.'),('Leave without pay','The deduction calculated when worked days are fewer than working days.'),('Other deduction','Another entered deduction, such as a separately approved adjustment.'),('Net salary / net payable','Gross salary minus CPF, loan, unpaid-day and other deductions.'),('Draft / Locked','Draft can recalculate. Locked uses saved final payroll figures.'),('Backup / Export','Backup saves the database. Export creates a PDF, Excel or Word report.')]
for i,(term,definition) in enumerate(terms):
    col=i//8;row=i%8;x=32+col*398;top=484-row*52
    para(x,top,366,escape(term),10.5,14,bold=True)
    para(x,top-17,366,escape(definition),9.5,12.5,color=MUTED,max_height=32)
finish()

# 24 - Monthly checklist and troubleshooting.
start('Your monthly desk checklist','Keep this page nearby','Work through the left side each month. Use the right side when something needs attention.')
text(32,482,'BEFORE YOU FINISH',13,'GuideBold',TEAL)
checks=[('Choose the correct month','Check the period before editing or exporting.'),('Check the staff list','Confirm active staff and bank/account details.'),('Enter the monthly changes','Days worked, arrears, allowances and deductions.'),('Save every edited form','A calculation preview alone does not save data.'),('Review the totals','Check unusual deductions and negative net pay.'),('Lock the month','Preserve the figures after your review.'),('Save the reports','Check Entire month versus Selected employees.'),('Back up the database','Keep a private copy on another drive.')]
for i,(title,body) in enumerate(checks):
    y=452-i*45
    C.setStrokeColor(TEAL);C.setLineWidth(1);C.roundRect(33,y-13,12,12,2,fill=0,stroke=1)
    para(55,y,335,escape(title),10.5,14,bold=True)
    para(55,y-17,335,escape(body),9.5,12.5,color=MUTED,max_height=26)
text(437,482,'QUICK FIXES',13,'GuideBold',TEAL)
fixes=[('The employee list looks empty','Clear the search. On Employees, try Show inactive. Check the month and whether records were imported or restored.'),('I cannot edit the monthly form','The month may be locked. Unlock only if a correction is intended, then review and lock it again.'),('The export is empty or incomplete','Check the export scope. Selected employees needs checked rows. A search alone does not set the report scope.'),('The file will not save or open','Close it in Excel or the PDF viewer, choose a writable folder and retry. A viewer is needed to open the saved file.'),('Email failed','Check internet, SMTP details and employee email addresses. Review results before retrying to avoid duplicate emails.'),('My old and new apps show different data','Use the new Payroll Desk shortcut. Run Legacy Payroll.cmd opens the old app, whose database is separate.')]
for i,(title,body) in enumerate(fixes):
    y=452-i*64
    para(437,y,372,escape(title),10.5,14,bold=True)
    para(437,y-18,372,escape(body),9.5,12.5,color=MUTED,max_height=42)
finish()

if PAGE!=24:raise ValueError(f'Expected 24 pages, got {PAGE}')
C.save()
(SRC/'page-index.json').write_text(json.dumps(pages,indent=2),encoding='utf-8')
print(f'Created {PDF} ({PAGE} pages)')
