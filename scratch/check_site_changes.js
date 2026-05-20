const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

raw.forEach((row, idx) => {
    if (idx <= 5) return;
    const s = (row[0] || '').toString().trim();
    if (s) {
        console.log(`Row ${idx}: "${s}"`);
    }
});
