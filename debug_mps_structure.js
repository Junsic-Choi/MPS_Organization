const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH);
const mpsName = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const sheet = workbook.Sheets[mpsName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

let out = "--- MPS Structure Scan ---\n";
out += "Sheet Name: " + mpsName + "\n\n";

const headers = raw[4] || []; 
out += "[Headers (Line 5)]\n";
headers.forEach((h, i) => {
    out += `${i}: ${h}\n`;
});

out += "\n[Sample Row (Row 7)]\n";
const row = raw[6] || [];
row.forEach((v, i) => {
    if (v !== undefined) out += `${i}: ${v}\n`;
});

fs.writeFileSync('debug_mps_structure.txt', out);
console.log('Done.');
