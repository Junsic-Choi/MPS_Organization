const XLSX = require('xlsx');
const fs = require('fs');
let log = "--- SCAN START ---\n";
try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[1]]; // MPS Sheet
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    data.forEach((row, r) => {
        const rowStr = JSON.stringify(row);
        if (rowStr.includes('DNC8060') || rowStr.includes('DNC 8060')) {
            log += `FOUND in MPS [Row ${r}]: ${rowStr}\n`;
        }
    });
} catch (e) {
    log += "ERROR: " + e.message + "\n";
}
fs.writeFileSync('mps_scan_dnc.txt', log);
