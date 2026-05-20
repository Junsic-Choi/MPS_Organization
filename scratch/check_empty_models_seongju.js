const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

let lastSite = '';
let totalSum = 0;
let modelSum = 0;
let noModelSum = 0;

raw.forEach((row, idx) => {
    if (idx <= 5) return;
    const s = (row[0] || '').toString().trim();
    if (s) lastSite = s;

    if (lastSite.includes('성주')) {
        const colsToSum = [4, 7, 8, 9, 10, 12];
        let rowSum = 0;
        colsToSum.forEach(c => {
            rowSum += parseInt(row[c]) || 0;
        });

        if (rowSum > 0) {
            totalSum += rowSum;
            const model = (row[2] || '').toString().trim();
            if (model) {
                modelSum += rowSum;
            } else {
                noModelSum += rowSum;
                console.log(`Row ${idx} has no model but sum=${rowSum}:`, row.slice(0, 15));
            }
        }
    }
});

console.log(`Total Seongju Qty (Col 4+7+8+9+10+12): ${totalSum}`);
console.log(`Sum of rows with model: ${modelSum}`);
console.log(`Sum of rows without model: ${noModelSum}`);
