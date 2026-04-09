// inspect_mps2603.js
const XLSX = require('xlsx');
const fs = require('fs');
try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const sheetName = workbook.SheetNames[1]; // 생산요약
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    let out = `Sheet: ${sheetName}\n`;
    for (let i = 0; i < 20; i++) {
        out += `Row ${i}: ${JSON.stringify(data[i])}\n`;
    }
    fs.writeFileSync('mps2603_inspect.txt', out);
} catch (e) {
    fs.writeFileSync('mps2603_inspect.txt', 'ERROR: ' + e.message);
}
