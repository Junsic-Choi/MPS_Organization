const XLSX = require('xlsx');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH, { bookSheets: true });
console.log('Sheet Names:', workbook.SheetNames);

const prodName = workbook.SheetNames.find(n => n.includes('배포'));
const mpsName = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS'));

const wbFull = XLSX.readFile(FILE_PATH, { sheets: [prodName, mpsName] });
const prodRaw = XLSX.utils.sheet_to_json(wbFull.Sheets[prodName], { header: 1 });
const mpsRaw = XLSX.utils.sheet_to_json(wbFull.Sheets[mpsName], { header: 1 });

console.log('\n--- Production Sheet (First 50 Sites Found) ---');
const foundSites = new Set();
prodRaw.slice(6).forEach(row => {
    if (row[0]) foundSites.add(row[0].toString().trim());
});
console.log('Sites found:', Array.from(foundSites));

console.log('\n--- Searching LYNX / XG in Production ---');
let lastSite = "";
prodRaw.slice(6).forEach((row, i) => {
    if (row[0]) lastSite = row[0].toString().trim();
    const model = (row[2] || '').toString();
    if (model.includes('LYNX') || model.includes('XG')) {
        console.log(`L${i+7}: Site=${lastSite}, Model=${model}, Qty2nd=${row[4]}, Qty3rd=${row[7]}`);
    }
});

console.log('\n--- Searching LYNX / XG in MPS Tab ---');
mpsRaw.forEach((row, i) => {
    const prod = (row[4] || '').toString();
    if (prod.includes('LYNX') || prod.includes('XG')) {
        console.log(`L${i+1}: Code=${row[3]}, Product=${prod}`);
    }
});
