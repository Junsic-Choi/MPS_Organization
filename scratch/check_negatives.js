const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets['MPS'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const monthCols = [
    { name: '4월', col: 12 },
    { name: '5월', col: 17 },
    { name: '6월', col: 22 },
    { name: '7월', col: 28 },
    { name: '8월', col: 34 },
    { name: '9월', col: 40 }
];

console.log('Row | Model | Product | Site | 4월 | 5월 | 6월 | 7월 | 8월 | 9월 | Sum');
console.log('---|---|---|---|---|---|---|---|---|---|---');

let seongjuCount = 0;

raw.forEach((row, idx) => {
    if (idx <= 4) return;
    const site = String(row[6] || '').trim();
    if (site === '1842' || site.includes('성주')) {
        let sum = 0;
        const q = monthCols.map(m => {
            const val = parseInt(row[m.col]) || 0;
            sum += val;
            return val;
        });

        if (sum > 0) {
            seongjuCount++;
            console.log(`${idx} | ${row[3] || ''} | ${row[4] || ''} | ${site} | ${q[0]} | ${q[1]} | ${q[2]} | ${q[3]} | ${q[4]} | ${q[5]} | ${sum}`);
        }
    }
});

console.log(`\nTotal rows: ${seongjuCount}`);
