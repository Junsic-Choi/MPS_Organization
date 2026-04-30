const XLSX = require('xlsx');
const wb = XLSX.readFile('c:/Users/i0215099/Desktop/MPS_UPDATE/MPS2603-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, {header:1});

const target = 4650;
const colSums = {};
for (let c = 0; c < 50; c++) {
    let sum = 0;
    for (let r = 6; r < raw.length; r++) {
        const row = raw[r] || [];
        const model = (row[2] || '').toString().trim();
        if (!model || model === '총합계') continue;
        sum += parseInt(row[c]) || 0;
    }
    if (sum > 0) colSums[c] = sum;
}

const entries = Object.entries(colSums).map(([c, s]) => ({ col: parseInt(c), sum: s }));

function findSubset(target, items) {
    const n = items.length;
    for (let i = 0; i < (1 << n); i++) {
        let currentSum = 0;
        let subset = [];
        for (let j = 0; j < n; j++) {
            if ((i >> j) & 1) {
                currentSum += items[j].sum;
                subset.push(items[j].col);
            }
        }
        if (currentSum === target) return subset;
    }
    return null;
}

const result = findSubset(target, entries);
console.log('Target:', target);
console.log('Resulting Columns:', result);
if (result) {
    result.forEach(c => console.log(`Col ${c}: ${colSums[c]}`));
}
