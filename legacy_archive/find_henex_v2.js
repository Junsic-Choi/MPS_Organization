const XLSX = require('xlsx');
const workbook = XLSX.readFile('MPS2603-1.xlsx');
workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    data.forEach((row, r) => {
        row.forEach((cell, c) => {
            if (cell && (cell.toString().includes('DNC8060') || cell.toString().includes('Henex') || cell.toString().includes('헤넥스'))) {
                console.log(`Sheet: ${sheetName}, Row: ${r + 1}, Col: ${c + 1}, Value: ${cell}`);
            }
        });
    });
});
