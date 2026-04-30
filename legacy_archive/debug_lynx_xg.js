const XLSX = require('xlsx');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH, { bookSheets: true });

const mpsName = workbook.SheetNames.find(n => n.toUpperCase().includes('MPS'));
const prodName = workbook.SheetNames.find(n => n.includes('배포'));

const wbFull = XLSX.readFile(FILE_PATH, { sheets: [mpsName, prodName] });
const mpsRaw = XLSX.utils.sheet_to_json(wbFull.Sheets[mpsName], { header: 1 });
const prodRaw = XLSX.utils.sheet_to_json(wbFull.Sheets[prodName], { header: 1 });

console.log('--- Search in Production (배포) ---');
let lastSite = "";
for (let r = 0; r < prodRaw.length; r++) {
    const row = prodRaw[r] || [];
    if (row[0]) lastSite = row[0].toString().trim();
    if (lastSite.includes('휴텍')) {
        const model = (row[2] || '').toString();
        if (model.includes('LYNX') || model.includes('XG')) {
            console.log(`Row ${r+1}: Site=${lastSite}, Model=${model}, RPM=${row[3]}, 2월=${row[4]}, 3월=${row[7]}`);
        }
    }
}

console.log('\n--- Search in MPS (Potential XG matches) ---');
for (let r = 0; r < mpsRaw.length; r++) {
    const row = mpsRaw[r] || [];
    const prod = (row[4] || '').toString();
    if (prod.includes('XG') || prod.includes('LYNX')) {
        console.log(`Row ${r+1}: Code=${row[3]}, Product=${prod}`);
    }
}
