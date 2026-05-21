const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets['MPS'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

const row4 = raw[4] || [];
const monthCols = [
    { name: '4월', col: 12 },
    { name: '5월', col: 17 },
    { name: '6월', col: 22 },
    { name: '7월', col: 28 },
    { name: '8월', col: 34 },
    { name: '9월', col: 40 }
];

console.log('=== CHECKING SKIPPED OR MISSED ROWS OR QTY DIFFERENCES ===');

let totalQtyAllRows = 0;
let totalQtySeongjuAllRows = 0;

let skippedRows = 0;

raw.forEach((row, idx) => {
    if (idx <= 4) return; // skip headers
    
    // Calculate total quantity in 4-9월 for this row
    let rowPlanQty = 0;
    monthCols.forEach(m => {
        rowPlanQty += parseInt(row[m.col]) || 0;
    });

    if (rowPlanQty > 0) {
        const mModel = String(row[3] || '').trim();
        const pName = String(row[4] || '').trim();
        const site = String(row[6] || '').trim();
        
        totalQtyAllRows += rowPlanQty;
        if (site === '1842' || site.includes('성주')) {
            totalQtySeongjuAllRows += rowPlanQty;
        }

        if (!mModel && !pName) {
            skippedRows++;
            console.log(`Skipped Row ${idx}: Site="${site}", Plan Qty=${rowPlanQty}`);
        }
    }
});

console.log(`\nSum of all rows in Excel (4-9월): ${totalQtyAllRows}`);
console.log(`Sum of all Seongju rows in Excel (4-9월): ${totalQtySeongjuAllRows}`);
console.log(`Number of skipped rows (no Model and no Product): ${skippedRows}`);
