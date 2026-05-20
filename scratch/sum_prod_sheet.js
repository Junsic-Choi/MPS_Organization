const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

let lastSite = '';
const siteSums = {};
let grandTotal = 0;

raw.forEach((row, idx) => {
    if (idx <= 5) return;
    const s = (row[0] || '').toString().trim();
    if (s) lastSite = s;

    const cleanSite = lastSite.replace(/^\d+\.\s*/, '').trim();

    // Sum Columns 4, 7, 8, 9, 10, 12 (4월 실적 and 5-9월 계획)
    const colsToSum = [4, 7, 8, 9, 10, 12];
    let rowSum = 0;
    colsToSum.forEach(c => {
        rowSum += parseInt(row[c]) || 0;
    });

    if (rowSum > 0) {
        siteSums[cleanSite] = (siteSums[cleanSite] || 0) + rowSum;
        grandTotal += rowSum;
    }
});

console.log('생산배포용 site sums (Col 4+7+8+9+10+12):');
console.log(siteSums);
console.log('Grand Total in 생산배포용:', grandTotal);
