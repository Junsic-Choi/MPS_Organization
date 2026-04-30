const XLSX = require('xlsx');
const fs = require('fs');
try {
    const wb = XLSX.readFile('data_working.xlsx');
    fs.writeFileSync('node_sheets.txt', wb.SheetNames.join(', '));
} catch (e) {
    fs.writeFileSync('node_sheets.txt', 'ERROR: ' + e.message);
}
