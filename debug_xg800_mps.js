const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH);
const mpsName = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const sheet = workbook.Sheets[mpsName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

let out = "--- MPS XG800 Data Deep Dive ---\n";
for (let r = 0; r < raw.length; r++) {
    const row = raw[r];
    const prod = (row[4] || '').toString();
    if (prod.includes('XG800')) {
        out += `Row ${r+1}: Product=${prod}, Site=${row[6]}, 2월=${row[8]}, 3월=${row[12]}, 4월=${row[17]}, 5월=${row[22]}, 6월=${row[28]}, 7월=${row[34]}\n`;
    }
}

fs.writeFileSync('debug_xg800_mps.txt', out);
console.log('Done.');
