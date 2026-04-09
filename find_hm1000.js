const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.Sheets[workbook.SheetNames[1]];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    
    let log = "Searching for HM1000 in Sheet 1 (MPS):\n";
    let found = false;
    for (let r = 0; r < mpsRaw.length; r++) {
        const rowStr = JSON.stringify(mpsRaw[r] || []);
        if (rowStr.toUpperCase().includes('HM1000')) {
            log += `FOUND at Row ${r}: ${rowStr}\n`;
            found = true;
        }
    }
    if (!found) log += "HM1000 NOT FOUND in Sheet 1.\n";

    fs.writeFileSync('hm1000_check.txt', log);
} catch (e) {
    fs.writeFileSync('hm1000_check.txt', 'ERROR: ' + e.message);
}
