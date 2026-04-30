const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.Sheets[workbook.SheetNames[1]];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    const mpsMonthIdxs = [8, 12, 17, 22, 28, 34];
    const monthNames = ["2월", "3월", "4월", "5월", "6월", "7월"];
    
    let log = "Monthly Production Span in Sheet 1 (MPS):\n";
    monthNames.forEach((month, mIdx) => {
        let firstRow = -1;
        let lastRow = -1;
        let totalQty = 0;
        const colIdx = mpsMonthIdxs[mIdx];
        
        for (let r = 5; r < mpsRaw.length; r++) {
            const qty = parseInt(mpsRaw[r][colIdx]) || 0;
            if (qty > 0) {
                if (firstRow === -1) firstRow = r;
                lastRow = r;
                totalQty += qty;
            }
        }
        log += `${month}: Range [${firstRow} to ${lastRow}], Total Qty=${totalQty}\n`;
    });

    fs.writeFileSync('mps_monthly_span.txt', log);
} catch (e) {
    fs.writeFileSync('mps_monthly_span.txt', 'ERROR: ' + e.message);
}
