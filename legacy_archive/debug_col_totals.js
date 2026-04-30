const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const tempWB = XLSX.readFile(FILE_PATH, { bookSheets: true });
const prodName = tempWB.SheetNames.find(n => n.includes('배포'));
const workbook = XLSX.readFile(FILE_PATH, { sheets: [prodName] });
const sheet = workbook.Sheets[prodName];
const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });

// Headers are in row index 5 (Row 6)
const headerRow = raw[5] || [];
let out = `Sheet: [${prodName}]\n\n`;
out += `Header Row (Row 6): \n`;
headerRow.forEach((h, i) => {
    if (h) out += `  [${i}] = "${h}"\n`;
});

// Count non-zero values per column (col 4~15) across all data rows
out += '\n--- Data Column Totals (col 4-15 across all rows) ---\n';
const colTotals = {};
for (let c = 4; c <= 15; c++) colTotals[c] = 0;
for (let r = 6; r < raw.length; r++) {
    const row = raw[r] || [];
    for (let c = 4; c <= 15; c++) {
        const v = parseInt(row[c]) || 0;
        colTotals[c] += v;
    }
}
Object.keys(colTotals).forEach(c => {
    const h = headerRow[parseInt(c)] || '(no header)';
    out += `  Col[${c}] "${h}": Total = ${colTotals[c]}\n`;
});

fs.writeFileSync('debug_col_totals.txt', out);
console.log('Done.');
