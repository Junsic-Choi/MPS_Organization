const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH);

const mpsName = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const prodName = workbook.SheetNames.find(n => n.includes('배포'));

const mpsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[mpsName], { header: 1 });
const prodRaw = XLSX.utils.sheet_to_json(workbook.Sheets[prodName], { header: 1 });

let out = "--- Debug LYNX XG ---\n";

out += "\n[Production Sheet Scan for LYNX/XG]\n";
let lastSite = "";
prodRaw.slice(6).forEach((row, i) => {
    if (row[0]) lastSite = row[0].toString().trim();
    const model = (row[2] || '').toString();
    if (model.includes('LYNX') || model.includes('XG')) {
        out += `L${i+7}: Site=${lastSite}, Model=${model}, Qty2nd=${row[4]}, Qty3rd=${row[7]}\n`;
    }
});

out += "\n[MPS Sheet Scan for LYNX/XG]\n";
mpsRaw.forEach((row, i) => {
    const prod = (row[4] || '').toString();
    if (prod.includes('LYNX') || prod.includes('XG')) {
        out += `L${i+1}: Code=${row[3] || ''}, Product=${prod}\n`;
    }
});

fs.writeFileSync('debug_lynx_result.txt', out);
console.log('Done. Check debug_lynx_result.txt');
