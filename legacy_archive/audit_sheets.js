const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const target = process.argv[2] || 'MPS2603-1.xlsx';
const filePath = path.join(__dirname, target);

if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
}

const wb = XLSX.readFile(filePath, { bookSheets: true });
console.log('--- SHEET NAMES ---');
console.log(wb.SheetNames);
console.log('-------------------');
