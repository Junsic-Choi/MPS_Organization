const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets['MPS'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const row2 = raw[2] || [];
const row3 = raw[3] || [];
const row4 = raw[4] || [];

console.log('--- Row 3 (Excel-provided Sums) ---');
for (let c = 8; c < Math.max(row2.length, row3.length, row4.length); c++) {
    if (row4[c] === '생산') {
        let label = '';
        for (let idx = c; idx >= 0; idx--) {
            if (row2[idx]) {
                label = row2[idx];
                break;
            }
        }
        console.log(`Col ${c} (${label} - 생산): Excel Row3 Value = "${row3[c] || ''}", Calculated Sum = ${getColSum(raw, c)}`);
    }
}

function getColSum(raw, col) {
    let sum = 0;
    for (let r = 5; r < raw.length; r++) {
        sum += parseInt((raw[r] || [])[col]) || 0;
    }
    return sum;
}
