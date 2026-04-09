// debug_codes.js
const XLSX = require('xlsx');
const fs = require('fs');
try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS')) || workbook.SheetNames[1];
    const mpsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[mpsSheet], { header: 1 });
    let log = "MPS Sheet Products:\n";
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        if (row[4]) log += `Code: ${row[3]}, Product: ${row[4]}\n`;
    }
    fs.writeFileSync('all_mps_codes.txt', log);
} catch (e) {
    fs.writeFileSync('all_mps_codes.txt', 'ERROR: ' + e.message);
}
