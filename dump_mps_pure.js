const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[1]];
    const csv = XLSX.utils.sheet_to_csv(sheet);
    fs.writeFileSync('mps_pure_dump.csv', csv);
    console.log('MPS Pure Dump Complete.');
} catch (e) {
    console.error(e);
}
