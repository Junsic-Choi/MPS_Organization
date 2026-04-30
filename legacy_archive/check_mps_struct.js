const XLSX = require('xlsx');
const fs = require('fs');
const workbook = XLSX.readFile('MPS2603-1.xlsx');
const mpsSheet = workbook.Sheets[workbook.SheetNames[1]];
const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
let log = "MPS Sheet Structure Check:\n";
log += "Row 4 (Header): " + JSON.stringify(mpsRaw[4]) + "\n";
for (let i = 5; i < 15; i++) {
    log += `Row ${i}: ` + JSON.stringify(mpsRaw[i]) + "\n";
}
fs.writeFileSync('mps_structure_deep.txt', log);
