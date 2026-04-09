// inspect_all_sheets.js
const XLSX = require('xlsx');
const fs = require('fs');
try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    let out = `Sheet Count: ${workbook.SheetNames.length}\n`;
    workbook.SheetNames.forEach((n, idx) => {
        out += `\n--- Sheet ${idx}: ${n} ---\n`;
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[n], { header: 1 });
        for (let i = 0; i < 10; i++) {
            out += `Row ${i}: ${JSON.stringify(data[i])}\n`;
        }
    });
    fs.writeFileSync('all_sheets_inspect.txt', out);
} catch (e) {
    fs.writeFileSync('all_sheets_inspect.txt', 'ERROR: ' + e.message);
}
