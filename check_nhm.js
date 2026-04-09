// check_nhm.js
const XLSX = require('xlsx');
const fs = require('fs');
try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[1]], { header: 1 });
    let results = "Searching for NHM:\n";
    for (let r = 0; r < mpsData.length; r++) {
        const row = mpsData[r] || [];
        const p = (row[4] || '').toString();
        if (p.includes('NHM')) results += `Row ${r}: Code=${row[3]}, Prod=${p}\n`;
    }
    fs.writeFileSync('nhm_check.txt', results);
} catch (e) {
    fs.writeFileSync('nhm_check.txt', 'ERROR: ' + e.message);
}
