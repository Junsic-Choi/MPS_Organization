const fs = require('fs');
const xlsx = require('xlsx');

try {
    const filename = '일반비_MPS2603-1(생산배포용).xlsx';
    const workbook = xlsx.readFile(filename);
    const sheetName = workbook.SheetNames.find(n => n.includes('생산배포용'));
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    let output = '';
    for (let r = 0; r < 15; r++) {
        output += `Row ${r + 1}: ${JSON.stringify(data[r])}\n`;
    }
    fs.writeFileSync('diag_dump.txt', output);
    console.log('Dump completed');
} catch (e) {
    fs.writeFileSync('diag_dump.txt', 'Error: ' + e.toString());
}
