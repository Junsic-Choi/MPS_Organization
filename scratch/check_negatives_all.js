const XLSX = require('xlsx');
const fs = require('fs');

const filename = 'MPS2603-1.xlsx';
const wb = XLSX.readFile(filename);
const sheetNames = wb.SheetNames;
const masterWs = wb.Sheets[sheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS'];

if (masterWs) {
    const raw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
    const headerRow = raw[4]; // Row 5 is header
    
    console.log(`Checking for negative values in ${filename}...`);
    for (let r = 5; r < raw.length; r++) {
        const row = raw[r] || [];
        row.forEach((cell, c) => {
            if (typeof cell === 'number' && cell < 0) {
                console.log(`Row ${r+1}, Col ${c+1} (${headerRow[c]}): Value ${cell}, Code: ${row[3]}, Product: ${row[4]}`);
            }
        });
    }
}
