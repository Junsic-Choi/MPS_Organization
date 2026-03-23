const XLSX = require('xlsx');
const fs = require('fs');
try {
    const wb = XLSX.readFile('data_working.xlsx');
    const msg = 'SUCCESS: Read data_working.xlsx\nSheets: ' + wb.SheetNames.join(', ');
    fs.writeFileSync('test_read_output.txt', msg);
} catch (e) {
    fs.writeFileSync('test_read_output.txt', 'ERROR: ' + e.message);
}
