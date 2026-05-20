const fs = require('fs');
const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

let lastSite = '';
let lastGroup = '';

raw.forEach((row, idx) => {
    if (idx < 2) return;
    const s = (row[0] || '').toString().trim();
    const g = (row[1] || '').toString().trim();
    if (s) lastSite = s;
    if (g) lastGroup = g;

    if (idx >= 225 && idx <= 245) {
        console.log(`Row ${idx}: s="${row[0] || ''}" (lastSite="${lastSite}"), g="${row[1] || ''}" (lastGroup="${lastGroup}"), model="${row[2] || ''}"`);
    }
});
