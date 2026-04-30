const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const tempWB = XLSX.readFile(FILE_PATH, { bookSheets: true });
const mpsName = tempWB.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const workbook = XLSX.readFile(FILE_PATH, { sheets: [mpsName] });
const sheet = workbook.Sheets[mpsName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

let out = `Searching for "VCF" in MPS Sheet: [${mpsName}]\n\n`;
raw.forEach((row, i) => {
    const rowStr = JSON.stringify(row);
    if (rowStr.toUpperCase().includes('VCF')) {
        out += `Row ${i+1}: ${rowStr}\n`;
    }
});

fs.writeFileSync('debug_vcf_search.txt', out);
console.log('Done.');
