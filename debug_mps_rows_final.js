const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH);
const mpsName = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const sheet = workbook.Sheets[mpsName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

let out = "--- MPS Row Verification ---\n";
[5, 6].forEach(r => {
    const row = raw[r] || [];
    out += `Row ${r+1}: Code=${row[3]}, Product=${row[4]}, raw=[${row.join('|')}]\n`;
});

fs.writeFileSync('debug_mps_rows_final.txt', out);
console.log('Done.');
