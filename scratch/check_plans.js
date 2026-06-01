const XLSX = require('xlsx');

const wb = XLSX.readFile('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
const masterWs = wb.Sheets[wb.SheetNames[1]];
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

// Find Month columns
const typeRowIdx = 4; // Row 5 is header
const monthRowIdx = 2; // Row 3 has months
const monthRow = masterRaw[monthRowIdx];
const typeRow = masterRaw[typeRowIdx];

const monthCols = [];
typeRow.forEach((cell, idx) => {
    if (String(cell).trim() === '생산') {
        // Find preceding non-empty month
        for (let c = idx; c >= 0; c--) {
            if (monthRow[c]) {
                monthCols.push({ name: String(monthRow[c]).trim(), colIdx: idx });
                break;
            }
        }
    }
});

console.log('Detected Month Columns:', monthCols);

console.log('\n=== SEONGJU DC/DCM ROW PLAN QUANTITIES IN MASTER ===');
masterRaw.forEach((row, idx) => {
    if (idx < 5) return;
    const pl = row[1];
    const group = row[2];
    const model = row[3];
    const product = row[4];
    const site = row[6];
    if (site == '1842' && product && (product.startsWith('DC') || product.startsWith('DCM'))) {
        const plans = [];
        monthCols.forEach(m => {
            const qty = parseInt(row[m.colIdx]) || 0;
            if (qty > 0) {
                plans.push(`${m.name}: ${qty}`);
            }
        });
        if (plans.length > 0) {
            console.log(`Row ${idx+1}: Product="${product}", Plans=[${plans.join(', ')}]`);
        }
    }
});
