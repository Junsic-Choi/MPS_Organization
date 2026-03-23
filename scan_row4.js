const XLSX = require('xlsx');
const fs = require('fs');
try {
    const buffer = fs.readFileSync('data_working.xlsx');
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const ws = workbook.Sheets[workbook.SheetNames[1]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const row4 = data[3] || [];
    let res = "Row 4 Categories:\n";
    row4.forEach((v, i) => {
        if (v) res += `Col ${i + 1}: [${v}]\n`;
    });
    fs.writeFileSync('row4_categories.txt', res);
} catch (e) {
    fs.writeFileSync('row4_categories.txt', 'ERROR: ' + e.message);
}
