const XLSX = require('xlsx');
const { processMpsFile } = require('../extractor');

const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets['MPS'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const row4 = raw[4] || [];
const monthCols = [
    { name: '3월', col: 8 },
    { name: '4월', col: 12 },
    { name: '5월', col: 17 },
    { name: '6월', col: 22 },
    { name: '7월', col: 28 },
    { name: '8월', col: 34 },
    { name: '9월', col: 40 }
];

console.log('=== RAW EXCEL SUMS BY MONTH ===');
const rawMonthlyTotal = {};
const rawMonthlySeongju = {}; // original site 1842 or contains '성주'
const rawMonthlyNamsan = {};  // original site 1840 or contains '남산'

for (let m of monthCols) {
    rawMonthlyTotal[m.name] = 0;
    rawMonthlySeongju[m.name] = 0;
    rawMonthlyNamsan[m.name] = 0;
}

raw.forEach((row, idx) => {
    if (idx <= 4) return; // skip headers
    const origSite = String(row[6] || '').trim();
    if (!origSite && !row[3] && !row[4]) return; // empty row

    const isSeongju = (origSite === '1842' || origSite.includes('성주'));
    const isNamsan = (origSite === '1840' || origSite.includes('남산'));

    monthCols.forEach(m => {
        const val = parseInt(row[m.col]) || 0;
        rawMonthlyTotal[m.name] += val;
        if (isSeongju) {
            rawMonthlySeongju[m.name] += val;
        }
        if (isNamsan) {
            rawMonthlyNamsan[m.name] += val;
        }
    });
});

console.log('Total Raw monthly sum:');
console.log(rawMonthlyTotal);
let rawGrandTotal = 0;
let raw4to9Total = 0;
monthCols.forEach(m => {
    rawGrandTotal += rawMonthlyTotal[m.name];
    if (m.name !== '3월') raw4to9Total += rawMonthlyTotal[m.name];
});
console.log(`Grand Total (3-9월): ${rawGrandTotal}`);
console.log(`Plan Total (4-9월): ${raw4to9Total}`);

console.log('\nSeongju (Original Site 1842) Raw monthly sum:');
console.log(rawMonthlySeongju);
let rawSeongjuTotal = 0;
let rawSeongju4to9Total = 0;
monthCols.forEach(m => {
    rawSeongjuTotal += rawMonthlySeongju[m.name];
    if (m.name !== '3월') rawSeongju4to9Total += rawMonthlySeongju[m.name];
});
console.log(`Seongju Grand Total (3-9월): ${rawSeongjuTotal}`);
console.log(`Seongju Plan Total (4-9월): ${rawSeongju4to9Total}`);


console.log('\n=== PROCESSED (ENGINE) SUMS BY MONTH ===');
const result = processMpsFile('MPS2604-1.xlsx');
const engMonthlyTotal = {};
const engMonthlySeongju = {};

for (let m of monthCols) {
    engMonthlyTotal[m.name] = 0;
    engMonthlySeongju[m.name] = 0;
}

result.finalResults.forEach(r => {
    engMonthlyTotal[r.Month] = (engMonthlyTotal[r.Month] || 0) + r.Qty;
    if (r.Site === '성주') {
        engMonthlySeongju[r.Month] = (engMonthlySeongju[r.Month] || 0) + r.Qty;
    }
});

console.log('Engine Processed monthly sum (All):');
console.log(engMonthlyTotal);
let engGrandTotal = 0;
let eng4to9Total = 0;
monthCols.forEach(m => {
    engGrandTotal += engMonthlyTotal[m.name] || 0;
    if (m.name !== '3월') eng4to9Total += engMonthlyTotal[m.name] || 0;
});
console.log(`Engine Grand Total (3-9월): ${engGrandTotal}`);
console.log(`Engine Plan Total (4-9월): ${eng4to9Total}`);

console.log('\nEngine Processed monthly sum (Seongju):');
console.log(engMonthlySeongju);
let engSeongjuTotal = 0;
let engSeongju4to9Total = 0;
monthCols.forEach(m => {
    engSeongjuTotal += engMonthlySeongju[m.name] || 0;
    if (m.name !== '3월') engSeongju4to9Total += engMonthlySeongju[m.name] || 0;
});
console.log(`Engine Seongju Grand Total (3-9월): ${engSeongjuTotal}`);
console.log(`Engine Seongju Plan Total (4-9월): ${engSeongju4to9Total}`);
