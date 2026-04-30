const XLSX = require('xlsx');
const file = 'c:/Users/i0215099/Desktop/MPS_UPDATE/MPS2603-1.xlsx';
const wb = XLSX.readFile(file);
const ws = wb.Sheets['MPS'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

console.log('--- [MPS] Sheet Dump (Row 6-20) ---');
for (let i = 5; i < 20; i++) {
    const row = data[i];
    if (row) {
        console.log(`Row ${i+1}: D="${row[3]}", E="${row[4]}"`);
    }
}
