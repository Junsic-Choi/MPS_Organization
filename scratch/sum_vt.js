const fs = require('fs');
const XLSX = require('xlsx');

function checkVTL() {
    const buf = fs.readFileSync('MPS2604-1.xlsx');
    const wb = XLSX.read(buf, { type: 'buffer' });
    const mpsRaw = XLSX.utils.sheet_to_json(wb.Sheets['MPS'], { header: 1 });
    
    let totalQtyVT = 0;
    const monthCols = [8, 12, 17, 22, 28, 34];
    
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const prod = (row[4] || '').toString();
        if (prod.toUpperCase().includes('VT') || prod.toUpperCase().includes('VTL')) {
            let rowQty = 0;
            monthCols.forEach(c => {
                rowQty += parseInt(row[c]) || 0;
            });
            totalQtyVT += rowQty;
        }
    }
    console.log(`Total Quantity for VT/VTL models in MPS sheet: ${totalQtyVT}`);
}
checkVTL();
