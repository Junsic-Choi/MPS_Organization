const XLSX = require('xlsx');
const fs = require('fs');
let log = "--- SCAN START ---\n";
try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    log += "SheetNames: " + workbook.SheetNames.join(', ') + "\n";
    workbook.SheetNames.forEach(sn => {
        const sheet = workbook.Sheets[sn];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        data.forEach((row, r) => {
            const rowStr = JSON.stringify(row);
            if (rowStr.includes('헤넥스') || rowStr.includes('Henex') || rowStr.includes('DNC8060')) {
                log += `FOUND in ${sn} [Row ${r}]: ${rowStr}\n`;
            }
        });
    });
} catch (e) {
    log += "ERROR: " + e.message + "\n";
}
fs.writeFileSync('deep_scan_result.txt', log);
