const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.Sheets[workbook.SheetNames[1]];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    
    let log = "MPS Sheet Column Header Audit (Rows 0-4):\n";
    for (let r = 0; r < 5; r++) {
        log += `Row ${r}: ` + JSON.stringify(mpsRaw[r]) + "\n";
    }
    
    // Check sums for EACH individual "생산" column
    const headers = mpsRaw[4] || [];
    for (let c = 0; c < headers.length; c++) {
        if (headers[c] === "생산") {
            let sum = 0;
            for (let r = 5; r < mpsRaw.length; r++) {
                sum += (parseInt(mpsRaw[r][c]) || 0);
            }
            log += `Col ${c} ("생산"): Sum = ${sum}\n`;
        }
    }

    fs.writeFileSync('mps_col_audit.txt', log);
} catch (e) {
    fs.writeFileSync('mps_col_audit.txt', 'ERROR: ' + e.message);
}
