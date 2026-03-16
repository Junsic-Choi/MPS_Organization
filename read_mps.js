const fs = require('fs');
const path = require('path');

try {
    const xlsx = require('xlsx');
    const dir = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE';
    const filePath = path.join(dir, '일반비_MPS2603-1(생산배포용).xlsx');

    console.log(`Reading file: ${filePath}`);
    const workbook = xlsx.readFile(filePath);

    const sheetName = 'MPS';
    if (workbook.SheetNames.includes(sheetName)) {
        const sheet = workbook.Sheets[sheetName];
        const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });

        console.log(`Sheet: ${sheetName}`);
        for (let i = 0; i < 20 && i < json.length; i++) {
            console.log(`Row ${i + 1}: ` + JSON.stringify(json[i]));
        }
    } else {
        console.log(`Sheet [${sheetName}] not found. Available sheets: ${workbook.SheetNames.join(', ')}`);
    }
} catch (e) {
    console.error('Error: ' + e.toString());
}
