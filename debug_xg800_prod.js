const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH);
const prodName = workbook.SheetNames.find(n => n.includes('배포'));
const sheet = workbook.Sheets[prodName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

let out = "--- Production XG800 Data Deep Dive ---\n";
let lastSite = "";
const monthCols = [4, 7, 8, 9, 10, 12]; // 2월, 3월, 4월, 5월, 6월, 7월
for (let r = 0; r < raw.length; r++) {
    const row = raw[r] || [];
    if (row[0]) lastSite = row[0].toString().trim();
    const model = (row[2] || '').toString();
    if (lastSite.includes('휴텍') && model.includes('LYNX XG800')) {
        out += `Row ${r+1}: Site=${lastSite}, Model=${model}, RPM=${row[3]}, `;
        monthCols.forEach((col, i) => {
            out += `${i+2}월=${row[col] || 0}, `;
        });
        out += "\n";
    }
}

fs.writeFileSync('debug_xg800_prod.txt', out);
console.log('Done.');
