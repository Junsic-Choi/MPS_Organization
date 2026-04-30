const XLSX = require('xlsx');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH);
const mpsName = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const sheet = workbook.Sheets[mpsName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

console.log('--- MPS Header Scan ---');
const headers = raw[4] || []; // Headers are usually on row 5
headers.forEach((h, i) => {
    if (h) console.log(`${i}: ${h}`);
});

console.log('\n--- Sample Row (Row 7) ---');
const row = raw[6] || [];
row.forEach((v, i) => {
    if (v !== undefined) console.log(`${i}: ${v}`);
});
