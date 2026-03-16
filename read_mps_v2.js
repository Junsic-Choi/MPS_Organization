const fs = require('fs');
const path = require('path');

try {
    const xlsx = require('xlsx');
    const dir = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE';
    const filePath = path.join(dir, '일반비_MPS2603-1(생산배포용).xlsx');

    if (!fs.existsSync(filePath)) {
        console.error(`File not found: ${filePath}`);
        process.exit(1);
    }

    const workbook = xlsx.readFile(filePath);
    const sheetName = 'MPS';
    if (workbook.SheetNames.includes(sheetName)) {
        const sheet = workbook.Sheets[sheetName];
        const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });

        console.log(`SHEET_FOUND: ${sheetName}`);
        for (let i = 0; i < 30 && i < json.length; i++) {
            console.log(`ROW_${i}: ` + JSON.stringify(json[i]));
        }
    } else {
        console.log(`SHEET_NOT_FOUND: ${sheetName}. Available: ${workbook.SheetNames.join(', ')}`);
    }
} catch (e) {
    console.error('ERROR: ' + e.message);
}
