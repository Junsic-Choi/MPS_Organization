const XLSX = require('xlsx');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH, { bookSheets: true });
console.log('All Sheet Names:', workbook.SheetNames);

const prodNames = workbook.SheetNames.filter(n => n.includes('배포'));
console.log('Target Sheet Names:', prodNames);

prodNames.forEach(name => {
    console.log(`\n--- Content of Sheet: ${name} (Top 20 Rows) ---`);
    const sheet = XLSX.readFile(FILE_PATH, { sheets: [name] }).Sheets[name];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(0, 20);
    data.forEach((row, i) => console.log(`${i}: ${JSON.stringify(row)}`));
});
