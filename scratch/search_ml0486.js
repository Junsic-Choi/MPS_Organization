const XLSX = require('xlsx');
const fs = require('fs');

const filename = 'MPS2604-1.xlsx';
const wb = XLSX.readFile(filename);
const sheetNames = wb.SheetNames;
const masterWs = wb.Sheets[sheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS'];

if (masterWs) {
    const raw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
    console.log(`Searching for ML0486 in ${filename} MPS sheet...`);
    
    // Find header row (usually contains month names)
    let headerRow = [];
    for (let r = 0; r < 20; r++) {
        if ((raw[r] || []).some(cell => /^\d+월/.test((cell || '').toString()))) {
            headerRow = raw[r];
            console.log(`Header Row ${r+1}:`, JSON.stringify(headerRow));
            break;
        }
    }
    
    for (let r = 0; r < raw.length; r++) {
        const row = raw[r] || [];
        if (row.some(cell => (cell || '').toString().includes('ML0486'))) {
            console.log(`Row ${r+1}:`, JSON.stringify(row));
        }
    }
}
