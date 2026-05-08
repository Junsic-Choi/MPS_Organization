const XLSX = require('xlsx');
const fs = require('fs');

const filename = 'MPS2603-1.xlsx';
const wb = XLSX.readFile(filename);
const sheetNames = wb.SheetNames;
const masterWs = wb.Sheets[sheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS'];

if (masterWs) {
    const raw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
    console.log(`First 5 rows of ${filename} MPS sheet:`);
    for (let r = 0; r < 5; r++) {
        console.log(`Row ${r+1}:`, JSON.stringify(raw[r]));
    }
}
