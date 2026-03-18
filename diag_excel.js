const fs = require('fs');
const xlsx = require('xlsx');

try {
    const filename = '일반비_MPS2603-1(생산배포용).xlsx';
    if (!fs.existsSync(filename)) {
        console.error('File not found: ' + filename);
        process.exit(1);
    }

    const workbook = xlsx.readFile(filename);
    console.log('Sheets:', workbook.SheetNames.join(', '));

    const sheetName = workbook.SheetNames.find(n => n.includes('생산배포용'));
    if (!sheetName) {
        console.error('Target sheet not found');
        process.exit(1);
    }

    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    console.log('--- Row 7 (Header) ---');
    const row7 = data[6]; // 0-indexed
    if (row7) {
        row7.forEach((cell, idx) => {
            console.log(`Col ${idx + 1}: [${cell}]`);
        });
    }

    console.log('\n--- Row 8 (Sample Data) ---');
    const row8 = data[7];
    if (row8) {
        row8.forEach((cell, idx) => {
            console.log(`Col ${idx + 1}: [${cell}]`);
        });
    }

    console.log('\n--- Checking columns for 3, 4, 5 ---');
    if (row7) {
        row7.forEach((cell, idx) => {
            if (cell && cell.toString().match(/^(\d+)/)) {
                console.log(`Found Month ${cell} at Col ${idx + 1}, Row 8 Value: ${row8 ? row8[idx] : 'N/A'}`);
            }
        });
    }

} catch (e) {
    console.error(e.toString());
}
