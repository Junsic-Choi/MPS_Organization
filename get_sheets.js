const XLSX = require('xlsx');
const fs = require('fs');
const wb = XLSX.readFile('MPS2603-1.xlsx');
fs.writeFileSync('sheet_names.txt', wb.SheetNames.join('\n'));
