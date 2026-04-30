const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const readOptions = { Props: false, cellFormula: false, cellHTML: false, cellStyles: false, cellText: false, bookSheets: true };
const tempWB = XLSX.readFile(FILE_PATH, readOptions);
const mpsName = tempWB.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const prodName = tempWB.SheetNames.find(n => n.includes('배포'));

const workbook = XLSX.readFile(FILE_PATH, { ...readOptions, sheets: [mpsName, prodName], bookSheets: false });
const mpsRaw = XLSX.utils.sheet_to_json(workbook.Sheets[mpsName], { header: 1 });
const prodRaw = XLSX.utils.sheet_to_json(workbook.Sheets[prodName], { header: 1 });

const mpsMonthIdxs = [8, 12, 17, 22, 28, 34]; // I, M, R, W, AC, AI
const prodMonthIdxs = [4, 7, 8, 9, 10, 12];   // Col 5, 8, 9... (생산배포 시트)

let mpsTotal = 0;
mpsRaw.slice(5).forEach(row => {
    const code = (row[3] || '').toString();
    const prod = (row[4] || '').toString();
    if (!code || !prod || code.toUpperCase().includes('TOTAL')) return;
    mpsMonthIdxs.forEach(idx => mpsTotal += (parseInt(row[idx]) || 0));
});

let prodTotal = 0;
prodRaw.slice(6).forEach(row => {
    prodMonthIdxs.forEach(idx => prodTotal += (parseInt(row[idx]) || 0));
});

console.log(`[Audit Result]`);
console.log(`MPS Total (Sum of I,M,R,W,AC,AI): ${mpsTotal}`);
console.log(`Production Total (Sum of Feb-Jul cols): ${prodTotal}`);

fs.writeFileSync('audit_counts.txt', `MPS: ${mpsTotal}\nProduction: ${prodTotal}`);
