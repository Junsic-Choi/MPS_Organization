// check_new_source.js
const XLSX = require('xlsx');
try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    console.log('Sheets:', workbook.SheetNames);
    const sheet2 = workbook.Sheets[workbook.SheetNames[1]];
    const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });
    console.log('Sheet 2 Row 1:', data2[0]);
    console.log('Sheet 2 Row 7:', data2[6]);
} catch (e) {
    console.error('Error:', e.message);
}
