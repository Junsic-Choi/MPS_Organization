const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const row2 = raw[2] || [];
const row3 = raw[3] || [];

const prodCols = [];
row3.forEach((cell, idx) => {
    if (String(cell).trim() === '생산') {
        let label = String(row2[idx] || '').trim();
        if (!label) {
            // look backwards for a label
            for (let c = idx; c >= 0; c--) {
                if (String(row2[c] || '').trim()) {
                    label = String(row2[c]).trim();
                    break;
                }
            }
        }
        prodCols.push({ idx, label });
    }
});

console.log('Production Columns found in "생산배포용":');
prodCols.forEach(c => {
    console.log(`Col ${c.idx}: "${c.label}"`);
});

let currentSite = '';
let currentGroup = '';

const monthlySums = {};
const monthlySeongju = {};

prodCols.forEach(c => {
    monthlySums[c.label] = 0;
    monthlySeongju[c.label] = 0;
});

raw.forEach((row, idx) => {
    if (idx <= 5) return; // skip headers
    
    const siteVal = String(row[0] || '').trim();
    const groupVal = String(row[1] || '').trim();
    const modelVal = String(row[2] || '').trim();

    if (siteVal) currentSite = siteVal;
    if (groupVal) currentGroup = groupVal;

    // We only sum rows that have a Model (or are not empty summary rows)
    if (!modelVal) return;

    const isSeongju = currentSite.includes('성주') || currentSite.includes('1842');

    prodCols.forEach(c => {
        const val = parseInt(row[c.idx]) || 0;
        monthlySums[c.label] += val;
        if (isSeongju) {
            monthlySeongju[c.label] += val;
        }
    });
});

console.log('\n=== SUMS FROM "생산배포용" SHEET ===');
console.log('Monthly Sums (All Sites):');
console.log(monthlySums);

let grandTotalAll = 0;
let grandTotalPlan = 0;
prodCols.forEach(c => {
    grandTotalAll += monthlySums[c.label];
    if (!c.label.includes('3월')) {
        grandTotalPlan += monthlySums[c.label];
    }
});
console.log(`Grand Total (including 3월 실적): ${grandTotalAll}`);
console.log(`Grand Total Plan (excluding 3월 실적): ${grandTotalPlan}`);

console.log('\nMonthly Sums (Seongju):');
console.log(monthlySeongju);

let seongjuTotalAll = 0;
let seongjuTotalPlan = 0;
prodCols.forEach(c => {
    seongjuTotalAll += monthlySeongju[c.label];
    if (!c.label.includes('3월')) {
        seongjuTotalPlan += monthlySeongju[c.label];
    }
});
console.log(`Seongju Total (including 3월 실적): ${seongjuTotalAll}`);
console.log(`Seongju Total Plan (excluding 3월 실적): ${seongjuTotalPlan}`);
