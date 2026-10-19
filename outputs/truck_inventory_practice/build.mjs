import fs from 'node:fs/promises';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';
const wb = Workbook.create();
const sheet = wb.worksheets.add('FINAL LS');
const headers = ['KKG Load #','Retailer PO #','Item #','Qty','Description','DC','DC Name','Delivery Date'];
const rows = [];
const quantities = [[6,8,4,6],[10,12,8,10],[16,20,14,18]];
for (let load=0;load<3;load++) {
  for (let line=0;line<4;line++) {
    const full=line%2===0;
    rows.push([`PRACTICE-${101+load}`,`TEST-PO-${1001+load*2+Math.floor(line/2)}`,full?'681943924':'683054545',quantities[load][line],full?'Practice full pallet display':'Practice half pallet display',`TEST-${load+1}`,['Practice Chicago DC','Practice Dallas DC','Practice Atlanta DC'][load],new Date(Date.UTC(2026,9,22+load))]);
  }
}
sheet.getRange('A1:H13').values=[headers,...rows];
sheet.showGridLines=false;
sheet.getRange('A1:H13').format.font={name:'Arial',size:11,color:'#243247'};
sheet.getRange('A1:H13').format.rowHeight=24;
sheet.getRange('A1:H1').format={fill:'#243247',font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},rowHeight:30};
sheet.getRange('A1:C13').format.columnWidth=22;
sheet.getRange('D1:D13').format.columnWidth=9;
sheet.getRange('E1:E13').format.columnWidth=34;
sheet.getRange('F1:F13').format.columnWidth=13;
sheet.getRange('G1:G13').format.columnWidth=26;
sheet.getRange('H1:H13').format.columnWidth=19;
sheet.getRange('A2:C13').setNumberFormat('@');
sheet.getRange('D2:D13').setNumberFormat('#,##0');
sheet.getRange('H2:H13').setNumberFormat('mm/dd/yy');
for(let start=2; start<=10; start+=8) sheet.getRange(`A${start}:H${Math.min(start+3,13)}`).format.fill='#EEF3F8';
sheet.tables.add('A1:H13',true,'PracticeLoads');
sheet.freezePanes.freezeRows(1);
const guide=wb.worksheets.add('Practice Guide');
const notes=[
 ['Truck Inventory practice',''],
 ['Data','All loads, POs, destinations, and quantities are fictional practice data.'],
 ['1. Upload','In Truck Inventory > Inputs, upload this file under Upload combined load sheet.'],
 ['2. Parse','Parse the input, then select PRACTICE-101 to start.'],
 ['3. Item Setup','Enter the sample dimensions and weights below for each selected load.'],
 ['4. Build','Select Dry Van and build the truck plan. Try the other loads and compare results.'],
 ['Quantity meaning','Qty is the number of physical units with the dimensions entered in Item Setup.'],
 ['Units','Dimensions are inches. Weight is pounds per unit.'],
 ['Sample Item Setup',''],
 ['681943924','Full Pallet: Length 40, Width 48, Height 15.5, Weight 110. Stackable: No. Stack Qty: 1.'],
 ['683054545','Half Pallet: Length 20, Width 48, Height 52.7, Weight 350. Stackable: No. Stack Qty: 1.'],
 ['PRACTICE-101','24 units. Start here for a smaller mixed load.'],
 ['PRACTICE-102','40 units. Compare the larger mixed load with the first load.'],
 ['PRACTICE-103','68 units. Explore capacity handling and truck assignments.'],
 ['Optional exercise','Change stackability and Stack Qty in Item Setup and rebuild to compare results.'],
 ['Editing the file','Keep the FINAL LS column headers. Add or change data rows to create more scenarios.']
];
guide.getRange('A1:B16').values=notes;
guide.showGridLines=false;
guide.getRange('A1:B16').format.font={name:'Arial',size:11,color:'#243247'};
guide.getRange('A1:A16').format.columnWidth=30;
guide.getRange('B1:B16').format.columnWidth=112;
guide.getRange('A1:B16').format.rowHeight=27;
guide.getRange('A1').format.font={name:'Arial',size:15,bold:true};
guide.getRange('A9:B9').format.fill='#EEF3F8';
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:'FINAL LS',range:'A1:D13',maxChars:1800,tableMaxRows:4})).ndjson);
await (await SpreadsheetFile.exportXlsx(wb)).save('outputs/truck_inventory_practice/truck_inventory_practice.xlsx');
const preview=await wb.render({sheetName:'FINAL LS',range:'A1:H13',scale:1,format:'png'});
await fs.writeFile('outputs/truck_inventory_practice/preview.png',new Uint8Array(await preview.arrayBuffer()));
