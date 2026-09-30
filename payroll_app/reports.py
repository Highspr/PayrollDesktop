"""Printable payslips with wrapping text and consistent page margins."""
from pathlib import Path
import os
from xml.sax.saxutils import escape


def build_payslips(path, rows, settings, month_label, logo_directory):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image

    regular, bold = 'Helvetica', 'Helvetica-Bold'
    fontdir = Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts'
    if (fontdir / 'segoeui.ttf').exists() and (fontdir / 'segoeuib.ttf').exists():
        if 'PayrollRegular' not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont('PayrollRegular', str(fontdir / 'segoeui.ttf')))
            pdfmetrics.registerFont(TTFont('PayrollBold', str(fontdir / 'segoeuib.ttf')))
        regular, bold = 'PayrollRegular', 'PayrollBold'
    navy, teal, muted = colors.HexColor('#142b46'), colors.HexColor('#087f75'), colors.HexColor('#68768a')
    line, soft = colors.HexColor('#dfe6ee'), colors.HexColor('#f5f8fb')
    styles = {
        'body': ParagraphStyle('body', fontName=regular, fontSize=9, leading=13, textColor=navy, wordWrap='CJK'),
        'small': ParagraphStyle('small', fontName=regular, fontSize=8, leading=12, textColor=muted, wordWrap='CJK'),
        'title': ParagraphStyle('title', fontName=bold, fontSize=18, leading=23, textColor=navy, wordWrap='CJK'),
        'bold': ParagraphStyle('bold', fontName=bold, fontSize=9, leading=13, textColor=navy, wordWrap='CJK'),
        'right': ParagraphStyle('right', fontName=regular, fontSize=9, leading=13, alignment=TA_RIGHT, textColor=navy),
    }
    def p(value, style='body'):
        return Paragraph(escape(str(value or '—')), styles[style])
    def rs(value):
        return f'{value:,.0f}'
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=42, rightMargin=42, topMargin=38, bottomMargin=42,
                            title='Salary Payslips - ' + month_label, author=settings['college_name'])
    width = A4[0] - 84
    flow = []
    draft = 'DRAFT' in month_label.upper()
    logo = Path(logo_directory) / 'logo.png'
    for index, row in enumerate(rows):
        if index:
            flow.append(PageBreak())
        heading = [p(settings['college_name'], 'title'), Spacer(1, 5), p(settings['slip_header'], 'small')]
        if logo.exists():
            image = Image(str(logo))
            ratio = min(52 / image.imageWidth, 52 / image.imageHeight)
            image.drawWidth, image.drawHeight = image.imageWidth * ratio, image.imageHeight * ratio
            header = Table([[heading, image]], colWidths=[width-70,70])
        else:
            header = Table([[heading]], colWidths=[width])
        header.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0)]))
        flow += [header, Spacer(1,24), p('SALARY PAYSLIP' + ('  /  DRAFT' if draft else ''),'bold'), Spacer(1,5), p(month_label,'small'), Spacer(1,18)]
        details = [
            [p('Employee','small'),p(row['name'],'bold'),p('BPS','small'),p(row.get('bps') or 'Contract')],
            [p('Designation','small'),p(row.get('des')),p('Days worked','small'),p(f"{row['worked']} / {row['wd']}")],
            [p('Section','small'),p(row.get('section')),p('Appointment','small'),p(row.get('dt'))],
            [p('Bank','small'),p(row.get('bank')),p('Account no.','small'),p(row.get('acct'))],
        ]
        detail_table = Table(details, colWidths=[68, width*.43, 70, width-138-width*.43])
        detail_table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),soft),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),10),('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8)]))
        flow += [detail_table,Spacer(1,22)]
        earnings = [('Running basic pay',row['basic']),('House rent allowance',row['hr']),('Conveyance allowance',row['conv']),('Medical allowance',row['medl']),(row.get('adhoc_label') or 'Adhoc allowance',row['adhoc']),('Special allowance',row['spec']),('Arrears / encashment',row['arrear'])]
        deductions = [('Leave without pay',row.get('lwop_pay',row['lwop']-row.get('other',0))),('CPF contribution',row['cpf']),('CPF loan instalment',row['cp_loan']),('Other deduction',row.get('other',0))]
        values = [[p('EARNINGS','bold'),p('PKR','bold'),p('DEDUCTIONS','bold'),p('PKR','bold')]]
        for i in range(max(len(earnings),len(deductions))):
            e = earnings[i] if i<len(earnings) else ('',None)
            d = deductions[i] if i<len(deductions) else ('',None)
            values.append([p(e[0]) if e[0] else '',p(rs(e[1]),'right') if e[1] is not None else '',p(d[0]) if d[0] else '',p(rs(d[1]),'right') if d[1] is not None else ''])
        values.append([p('Gross salary','bold'),p(rs(row['gross']),'right'),p('Total deductions','bold'),p(rs(row['gross']-row['net']),'right')])
        amounts = Table(values,colWidths=[width*.31,width*.19,width*.31,width*.19])
        amounts.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('BACKGROUND',(0,0),(-1,0),soft),('BACKGROUND',(0,-1),(-1,-1),soft),('LINEBELOW',(0,0),(-1,0),.7,line),('LINEABOVE',(0,-1),(-1,-1),.7,line),('LINEAFTER',(1,0),(1,-1),.7,line),('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),10)]))
        flow += [amounts,Spacer(1,15)]
        net = Table([[p('NET PAYABLE','bold'),p('PKR '+rs(row['net']),'right')]],colWidths=[width*.6,width*.4])
        net.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#eaf7f2')),('TOPPADDING',(0,0),(-1,-1),14),('BOTTOMPADDING',(0,0),(-1,-1),14),('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),10)]))
        flow += [net,Spacer(1,22)]
        if settings.get('signatory_name') or settings.get('signatory_title'):
            flow += [p(settings.get('signatory_name'),'bold'),p(settings.get('signatory_title'),'small')]

    def page_footer(canvas, document):
        canvas.saveState()
        canvas.setStrokeColor(line)
        canvas.line(42,35,A4[0]-42,35)
        canvas.setFillColor(muted)
        canvas.setFont(regular,7)
        canvas.drawString(42,23,'Payroll Desk | '+('Draft - figures may change' if draft else 'Finalized payroll'))
        canvas.drawRightString(A4[0]-42,23,f'Page {document.page}')
        canvas.restoreState()
    doc.build(flow,onFirstPage=page_footer,onLaterPages=page_footer)
