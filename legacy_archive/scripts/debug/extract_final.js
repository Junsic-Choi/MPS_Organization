const XLSX = require('xlsx');
const fs = require('fs');

console.log('Loading workbook...');
const workbook = XLSX.readFile('data_working.xlsx');

// 1. Meta Data (Sheet 2)
console.log('Processing Meta (Sheet 2)...');
const metaSheet = workbook.Sheets[workbook.SheetNames[1]]; // 2nd sheet (index 1)
const metaData = XLSX.utils.sheet_to_json(metaSheet, { header: 1 });
const metaMap = {};
let lastS = '', lastG = '';

for (let i = 6; i < metaData.length; i++) {
    const row = metaData[i];
    if (!row) continue;
    if (row[0]) lastS = String(row[0]).trim();
    if (row[1]) lastG = String(row[1]).trim();
    if (row[2]) {
        const m = String(row[2]).trim();
        const key = m.toUpperCase().replace('LYNX ', '');
        metaMap[key] = {
            site: lastS,
            group: lastG,
            model: m,
            rpm: row[3] ? String(row[3]).trim() : ''
        };
    }
}
console.log('Meta Map built:', Object.keys(metaMap).length);

// 2. Main Data (Sheet 4)
console.log('Processing MPS (Sheet 4)...');
const mpsSheet = workbook.Sheets[workbook.SheetNames[3]]; // 4th sheet (index 3)
const mpsData = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });

const tCols = [8, 12, 17, 22, 28, 34]; // 0-indexed J, N, S, X, AD, AJ -> Wait! I used 9,13...
// Excel Cols: I(9), M(13), R(18), W(23), AC(29), AI(35)
// 0-indexed: 8, 12, 17, 22, 28, 34
const months = [];
const headerRow = mpsData[2]; // Row 3
tCols.forEach(idx => {
    let h = String(headerRow[idx] || '');
    let mNum = h.replace(/[^0-9]/g, '');
    months.push(mNum ? mNum + '월' : h);
});
console.log('Months:', months.join(', '));

const results = [['Site', 'Group', 'Model', 'RPM', 'Month', 'Code', 'Product']];
let lastC = '', lastP = '', curM = null;

for (let r = 6; r < mpsData.length; r++) { // Start from Row 7 (index 6)
    const row = mpsData[r];
    if (!row) continue;

    if (row[3]) lastC = String(row[3]).trim(); // Col D
    if (row[4]) {
        lastP = String(row[4]).trim(); // Col E
        const kP = lastP.split('-')[0].toUpperCase();
        curM = metaMap[kP];
        if (!curM) {
            // Fuzzy
            for (const mk in metaMap) {
                if (mk.includes(kP) || kP.includes(mk)) {
                    curM = metaMap[mk];
                    break;
                }
            }
        }
    }

    // Expansion
    tCols.forEach((colIdx, mIdx) => {
        const val = row[colIdx];
        const qVal = parseFloat(val);
        if (qVal > 0) {
            const mS = curM ? curM.site : '';
            const mG = curM ? curM.group : '';
            const mM = curM ? curM.model : '';
            const mR = curM ? curM.rpm : '';
            for (let q = 0; q < qVal; q++) {
                results.push([mS, mG, mM, mR, months[mIdx], lastC, lastP]);
            }
        }
    });
}

console.log('Total extracted rows:', results.length - 1);

const csvContent = results.map(row => row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
fs.writeFileSync('FinalList_Direct.csv', '\uFEFF' + csvContent); // BOM for Excel UTF8
console.log('Saved to FinalList_Direct.csv');
