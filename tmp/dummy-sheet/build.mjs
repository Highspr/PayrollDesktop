import fs from 'node:fs/promises';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';
const out='D:/payroll/outputs/payroll-testing-50';
await fs.mkdir(out,{recursive:true});
const wb=Workbook.create();
const salary=wb.worksheets.add('Salary Sheet');
const bank=wb.worksheets.add('Bank Advice');
const first=['Amina','Bilal','Sara','Hamza','Hina','Usman','Zara','Ali','Maryam','Omar'];
const last=['Ahmed','Khan','Malik','Iqbal','Raza'];
const headers=['No.','Appointment date','Employee name','BPS','Working days','Days worked','Arrears','Basic pay','House rent','Conveyance','Medical','Adhoc','Special','Gross pay','CPF deduction','Loan deduction','Unpaid leave + other','Net pay'];
salary.getRange('A1').values=[['Demo College, Calculation of Staff Salary for the Month of September 2026']];
salary.getRange('A2').values=[['TEST DATA ONLY: 50 fictional employees (30 Teaching, 20 Admin). All amounts in PKR.']];
salary.getRange('A3').values=[['Import through Employees > Import Excel. Use a test database: import adds employees and updates BPS scales.']];
salary.getRange('A4').values=[['Keep the title, column order and Teaching subtotal row. Save in Excel after changing formulas before importing.']];
salary.getRange('A5:R5').values=[headers];
const records=[];
for(let i=1;i<=50;i++){
 const r=i<=30?i+5:i+6;
 const bps=[5,9,11,14,16,17,18][(i-1)%7];
 const basic=18000+bps*2400+(i%4)*1000;
 const hr=bps*350,conv=1500+bps*100,med=1000+bps*50,adhoc=bps*400;
 const special=i%6===0?2500:0,arrear=i%7===0?3000:0,loan=i%8===0?1500:0;
 const worked=i%5===0?28:i%9===0?29:30;
 const cpf=i%4===0?0:bps*200;
 const name=`Test ${first[(i-1)%10]} ${last[Math.floor((i-1)/10)]}`;
 const row=[i,new Date(Date.UTC(2018+i%7,i%12,1+i%20)),name,bps,30,worked,arrear,basic,hr,conv,med,adhoc,special,null,cpf,loan,null,null];
 salary.getRange(`A${r}:R${r}`).values=[row];
 salary.getRange(`N${r}`).formulas=[[`=SUM(G${r}:M${r})`]];
 salary.getRange(`Q${r}`).formulas=[[`=ROUND(SUM(H${r}:M${r})/E${r}*(E${r}-F${r}),0)`]];
 salary.getRange(`R${r}`).formulas=[[`=N${r}-SUM(O${r}:Q${r})`]];
 if(i%2===0)salary.getRange(`A${r}:R${r}`).format.fill='#F1F5F9';
 records.push({row:r,name,bps,net:arrear+basic+hr+conv+med+adhoc+special-cpf-loan-Math.floor((basic+hr+conv+med+adhoc+special)/30*(30-worked)+0.5)});
}
salary.getRange('A36').values=[['Total Pay of Teaching']];
salary.getRange('A57').values=[['Total Pay of Admin']];
salary.getRange('A58').values=[['Grand total (50 staff)']];
for(const col of 'GHIJKLMNOPQR'){
 salary.getRange(`${col}36`).formulas=[[`=SUM(${col}6:${col}35)`]];
 salary.getRange(`${col}57`).formulas=[[`=SUM(${col}37:${col}56)`]];
 salary.getRange(`${col}58`).formulas=[[`=SUM(${col}36,${col}57)`]];
}
bank.getRange('A1').values=[['BANK ADVICE - TEST DATA ONLY']];
bank.getRange('A2').values=[['Demo Bank Limited']];
bank.getRange('A3').values=[['September 2026. Fictional accounts; not for payment.']];
bank.getRange('A5').values=[['This sheet supplies designations and bank details during import.']];
bank.getRange('A6').values=[['Keep employee records starting at row 15.']];
bank.getRange('A14:F14').values=[['No.','Employee name','Designation','Account number','Category','Net pay (PKR)']];
for(let i=0;i<50;i++){
 const r=i+15,s=records[i].row;
 bank.getRange(`A${r}:F${r}`).values=[[i+1,records[i].name,i<30?['Lecturer','Teacher','Senior Teacher'][i%3]:['Clerk','Lab Assistant','Librarian','Office Assistant'][i%4],`TEST-ACCT-${String(i+1).padStart(4,'0')}`,i<30?'Teaching':'Admin',null]];
 bank.getRange(`F${r}`).formulas=[[`='Salary Sheet'!R${s}`]];
 if(i%2===1)bank.getRange(`A${r}:F${r}`).format.fill='#F1F5F9';
}
bank.getRange('B65').values=[['Total']];bank.getRange('F65').formulas=[['=SUM(F15:F64)']];
for(const [ws,end,head] of [[salary,'R58','A5:R5'],[bank,'F65','A14:F14']]){
 ws.showGridLines=false;
 ws.getRange(`A1:${end}`).format.font={name:'Arial',size:10,color:'#172B4D'};
 ws.getRange(`A1:${end}`).format.rowHeight=22;
 ws.getRange(`A1:${end}`).format.verticalAlignment='center';
 ws.getRange('A1').format.font={size:14,bold:true};
 ws.getRange(head).format={fill:'#172B4D',font:{color:'#FFFFFF',bold:true},wrapText:true,rowHeight:40,horizontalAlignment:'center'};
}
salary.getRange('A1:R58').format.columnWidth=13;
salary.getRange('A1:A58').format.columnWidth=7;
salary.getRange('B1:B58').format.columnWidth=18;
salary.getRange('C1:C58').format.columnWidth=26;
salary.getRange('D1:F58').format.columnWidth=10;
salary.getRange('Q1:Q58').format.columnWidth=19;
salary.getRange('B6:B56').setNumberFormat('yyyy-mm-dd');
salary.getRange('O1:P58').format.columnWidth=17;
salary.getRange('G6:R58').setNumberFormat('#,##0');
for(const r of [36,57,58])salary.getRange(`A${r}:R${r}`).format={fill:'#DCE7F2',font:{bold:true}};
salary.freezePanes.freezeRows(5);
bank.getRange('A1:F65').format.columnWidth=22;
bank.getRange('A1:A65').format.columnWidth=7;
bank.getRange('B1:C65').format.columnWidth=26;
bank.getRange('D15:D64').setNumberFormat('@');
bank.getRange('F15:F65').setNumberFormat('#,##0');
bank.freezePanes.freezeRows(14);
wb.recalculate();
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:10},summary:'Formula error scan'})).ndjson);
console.log((await wb.inspect({kind:'table',range:'Salary Sheet!N58:R58',include:'values,formulas',tableMaxRows:2,tableMaxCols:5})).ndjson);
await (await SpreadsheetFile.exportXlsx(wb)).save(`${out}/Payroll_Test_Data_50_Employees.xlsx`);
for(const [sheetName,range,file] of [['Salary Sheet','A1:R12','salary'],['Bank Advice','A14:F22','bank']]){
 const img=await wb.render({sheetName,range,scale:1.5,format:'png'});
 await fs.writeFile(`D:/payroll/tmp/dummy-sheet/${file}.png`,new Uint8Array(await img.arrayBuffer()));
}
await fs.writeFile('D:/payroll/tmp/dummy-sheet/expected.json',JSON.stringify(records));
console.log('Created 50 employees. Expected net total: '+records.reduce((s,r)=>s+r.net,0));
