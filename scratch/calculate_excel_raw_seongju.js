const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets['MPS'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const row2 = raw[2] || [];
const row4 = raw[4] || [];

const numericCols = [];
for (let c = 8; c < row4.length; c++) {
    let hasNumeric = false;
    let sum = 0;
    for (let r = 5; r < raw.length; r++) {
        const val = parseInt((raw[r] || [])[c]) || 0;
        sum += val;
        if (val > 0) hasNumeric = true;
    }
    if (hasNumeric) {
        let label = '';
        for (let idx = c; idx >= 0; idx--) {
            if (row2[idx]) {
                label = row2[idx];
                break;
            }
        }
        numericCols.push({ idx: c, label, header: row4[c], sum });
    }
}

console.log('Numeric columns in MPS sheet:');
numericCols.forEach(nc => {
    console.log(`Col ${nc.idx}: label="${nc.label}", header="${nc.header}", sum=${nc.sum}`);
});

console.log('\nSeongju (Site 1842) sums by column:');
numericCols.forEach(nc => {
    let sjSum = 0;
    for (let r = 5; r < raw.length; r++) {
        const row = raw[r] || [];
        const site = String(row[6] || '').trim();
        const isSeongju = (site === '1842' || site.includes('성주'));
        if (isSeongju) {
            sjSum += parseInt(row[nc.idx]) || 0;
        }
    }
    console.log(`Col ${nc.idx} (${nc.label} - ${nc.header}): sjSum=${sjSum}`);
});
