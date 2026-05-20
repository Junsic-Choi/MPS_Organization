const fs = require('fs');
const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

raw.forEach((row, idx) => {
    const rowStr = row.join(' | ');
    if (rowStr.toUpperCase().includes('DVF4000')) {
        console.log(`Row ${idx}:`, row.slice(0, 10));
    }
});
