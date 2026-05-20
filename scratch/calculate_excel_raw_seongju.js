const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const headerRow = raw[2];
console.log('Header Row 2:', headerRow);

let lastSite = '';
let seongjuRows = [];
const monthlySums = {};

raw.forEach((row, idx) => {
    if (idx <= 5) return; // Skip headers and metadata
    const s = (row[0] || '').toString().trim();
    if (s) lastSite = s;

    if (lastSite.includes('성주')) {
        seongjuRows.push({ idx, row });
        // Let's sum for each column that has a numeric header or contains '생산'
        headerRow.forEach((colName, colIdx) => {
            if (colIdx >= 4) {
                const val = parseInt(row[colIdx]) || 0;
                const key = `${colIdx}: ${colName || ''}`;
                monthlySums[key] = (monthlySums[key] || 0) + val;
            }
        });
    }
});

console.log('Seongju Monthly Sums in 생산배포용:');
console.log(monthlySums);

let totalSeongjuInExcel = 0;
for (const [key, val] of Object.entries(monthlySums)) {
    if (key.includes('생산')) {
        totalSeongjuInExcel += val;
    }
}
console.log('Total Seongju production sum across all "생산" columns:', totalSeongjuInExcel);
