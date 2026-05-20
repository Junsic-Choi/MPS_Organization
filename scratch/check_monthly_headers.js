const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

console.log('Row 2:', raw[2]);
console.log('Row 3:', raw[3]);
console.log('Row 5:', raw[5]);

let lastSite = '';
const sums = {};

raw.forEach((row, idx) => {
    if (idx <= 5) return;
    const s = (row[0] || '').toString().trim();
    if (s) lastSite = s;

    if (lastSite.includes('성주')) {
        row.forEach((val, colIdx) => {
            const num = parseInt(val) || 0;
            if (num > 0) {
                sums[colIdx] = (sums[colIdx] || 0) + num;
            }
        });
    }
});

console.log('\nColumn sums for Seongju:');
for (const colIdx in sums) {
    console.log(`Col ${colIdx}: headerR2="${raw[2][colIdx] || ''}", headerR3="${raw[3][colIdx] || ''}", headerR5="${raw[5][colIdx] || ''}" -> sum=${sums[colIdx]}`);
}
