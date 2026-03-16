const fs = require('fs');

try {
    fs.writeFileSync('result_js.txt', 'Starting script...\n');
    const xlsx = require('xlsx');
    fs.appendFileSync('result_js.txt', 'xlsx required successfully.\n');

    const workbook = xlsx.readFile('MPS2603-1(생산배포용).xlsx');
    fs.appendFileSync('result_js.txt', 'Sheets: ' + workbook.SheetNames.join(', ') + '\n');

    if (workbook.SheetNames.includes('[생산배포용]')) {
        const sheet = workbook.Sheets['[생산배포용]'];
        const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });

        fs.appendFileSync('result_js.txt', 'First 10 rows:\n');
        for (let i = 0; i < 10 && i < json.length; i++) {
            fs.appendFileSync('result_js.txt', `Row ${i + 1}: ` + JSON.stringify(json[i]) + '\n');
        }
    } else {
        fs.appendFileSync('result_js.txt', 'Sheet [생산배포용] not found.\n');
    }
} catch (e) {
    fs.appendFileSync('result_js.txt', 'Error: ' + e.toString() + '\n' + e.stack);
}
