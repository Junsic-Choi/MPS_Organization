const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2604-1.xlsx');
console.log('Sheet Names:', wb.SheetNames);

const ws = wb.Sheets['생산배포용'];
if (!ws) {
    console.log('No "생산배포용" sheet found.');
    process.exit(0);
}

const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

console.log('\n--- First 15 rows of 생산배포용 ---');
for (let i = 0; i < Math.min(25, raw.length); i++) {
    console.log(`Row ${i}:`, (raw[i] || []).slice(0, 20).map(c => String(c || '').trim()));
}

// Let's find columns in 생산배포용
// We want to see what months/columns are there
const headerRow = raw[4] || [];
console.log('\nHeader Row (Row 4):', headerRow);

// Let's sum values in 생산배포용 by site and column
const siteSums = {};
let grandTotal = 0;

raw.forEach((row, idx) => {
    if (idx <= 4) return;
    const site = String(row[0] || '').trim();
    if (!site && !row[1] && !row[2]) return; // empty row

    // Let's find numeric columns
    let rowSum = 0;
    row.forEach((cell, cIdx) => {
        if (cIdx >= 3) {
            const val = parseInt(cell) || 0;
            rowSum += val;
        }
    });

    if (rowSum > 0) {
        // Find actual site name (handles merged cells by using last non-empty site)
    }
});

// Let's do a proper parse of 생산배포용
let currentSite = '';
let currentGroup = '';
const prodRows = [];

raw.forEach((row, idx) => {
    if (idx <= 4) return; // skip header
    const siteVal = String(row[0] || '').trim();
    const groupVal = String(row[1] || '').trim();
    const modelVal = String(row[2] || '').trim();

    if (siteVal) currentSite = siteVal;
    if (groupVal) currentGroup = groupVal;

    if (modelVal) {
        prodRows.push({
            site: currentSite.replace(/^\d+\.\s*/, '').trim(),
            group: currentGroup,
            model: modelVal,
            rowIdx: idx,
            cells: row
        });
    }
});

console.log(`\nProcessed ${prodRows.length} rows in 생산배포용.`);

// Let's see what columns are in Row 4 (or Row 3/2 if there are months)
const row2 = raw[2] || [];
const row3 = raw[3] || [];
const row4 = raw[4] || [];

console.log('\n--- Columns details for 생산배포용 ---');
for (let c = 0; c < Math.max(row2.length, row3.length, row4.length); c++) {
    if (row2[c] || row3[c] || row4[c]) {
        console.log(`Col ${c}: Row2="${row2[c] || ''}", Row3="${row3[c] || ''}", Row4="${row4[c] || ''}"`);
    }
}

// Let's sum columns that represent monthly plans/actuals in 생산배포용
// Typically columns are: 3월, 4월, 5월, 6월, 7월, 8월, 9월
// Let's find indices of columns that have month names in Row 2 or Row 3
const monthColsProd = [];
for (let c = 3; c < row4.length; c++) {
    const headerText = String(row2[c] || row3[c] || row4[c] || '');
    const mMatch = headerText.match(/(\d{1,2})\s*월/);
    if (mMatch) {
        monthColsProd.push({ name: mMatch[1] + '월', col: c });
    }
}

console.log('\nDetected Month Columns in 생산배포용:', monthColsProd);

const prodMonthlyTotal = {};
const prodMonthlySeongju = {};

monthColsProd.forEach(m => {
    prodMonthlyTotal[m.name] = 0;
    prodMonthlySeongju[m.name] = 0;
});

prodRows.forEach(r => {
    const isSeongju = r.site.includes('성주');
    monthColsProd.forEach(m => {
        const val = parseInt(r.cells[m.col]) || 0;
        prodMonthlyTotal[m.name] += val;
        if (isSeongju) {
            prodMonthlySeongju[m.name] += val;
        }
    });
});

console.log('\n=== 생산배포용 SUMS BY MONTH ===');
console.log('Total Monthly Sum in 생산배포용:', prodMonthlyTotal);
let prodGrandTotal = 0;
let prod4to9Total = 0;
monthColsProd.forEach(m => {
    prodGrandTotal += prodMonthlyTotal[m.name] || 0;
    if (m.name !== '3월') prod4to9Total += prodMonthlyTotal[m.name] || 0;
});
console.log(`Grand Total (All months): ${prodGrandTotal}`);
console.log(`Plan Total (4-9월): ${prod4to9Total}`);

console.log('\nSeongju Monthly Sum in 생산배포용:');
console.log(prodMonthlySeongju);
let prodSeongjuTotal = 0;
let prodSeongju4to9Total = 0;
monthColsProd.forEach(m => {
    prodSeongjuTotal += prodSeongjuMonthly = prodMonthlySeongju[m.name] || 0;
    if (m.name !== '3월') prodSeongju4to9Total += prodMonthlySeongju[m.name] || 0;
});
console.log(`Seongju Grand Total (All months): ${prodSeongjuTotal}`);
console.log(`Seongju Plan Total (4-9월): ${prodSeongju4to9Total}`);
