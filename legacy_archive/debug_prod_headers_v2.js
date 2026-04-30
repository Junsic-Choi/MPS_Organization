const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const tempWB = XLSX.readFile(FILE_PATH, { bookSheets: true });
const prodName = tempWB.SheetNames.find(n => n.includes('배포') && !n.includes('Check') && !n.includes('OLD'));
const workbook = XLSX.readFile(FILE_PATH, { sheets: [prodName] });
const sheet = workbook.Sheets[prodName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

let out = `Sheet: [${prodName}]\n`;
// Print Header rows (usually row 5 or 6)
[4, 5].forEach(r => {
    out += `Row ${r+1}: ${JSON.stringify(raw[r])}\n`;
});

// Print some data rows to see where numbers are
[6, 7, 10, 20, 50].forEach(r => {
    out += `Row ${r+1}: ${JSON.stringify(raw[r])}\n`;
});

fs.writeFileSync('debug_prod_headers_v2.txt', out);
console.log('Done.');
