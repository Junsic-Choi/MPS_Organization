const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const target = process.argv[2] || 'MPS2603-1.xlsx';
const filePath = path.join(__dirname, target);

let out = '';
try {
    if (!fs.existsSync(filePath)) {
        out = 'FILE_NOT_FOUND: ' + filePath;
    } else {
        const stat = fs.statSync(filePath);
        out += 'FILE_SIZE: ' + stat.size + '\n';
        const wb = XLSX.readFile(filePath);
        out += 'SHEETS: ' + JSON.stringify(wb.SheetNames) + '\n';
    }
} catch(e) {
    out = 'ERROR: ' + e.message;
}

fs.writeFileSync(path.join(__dirname, 'sheet_audit_result.txt'), out, 'utf8');
console.log('Done. Result written to sheet_audit_result.txt');
process.exit(0);
