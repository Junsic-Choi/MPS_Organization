const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('data_working.xlsx');
    const ws = workbook.Sheets[workbook.SheetNames[1]]; // 생산배포용
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    let dump = "SHEET 2 (생산배포용) STRUCTURE:\n";
    for (let r = 0; r < 20; r++) {
        const row = data[r] || [];
        dump += `Row ${r + 1}: ${row.slice(0, 50).map(v => v === undefined ? '' : v).join(' | ')}\n`;
    }

    fs.writeFileSync('node_sheet2_dump.txt', dump);
    console.log('Dump complete.');
} catch (e) {
    fs.writeFileSync('node_sheet2_dump.txt', 'ERROR: ' + e.stack);
}
