const XLSX = require('xlsx');
const workbook = XLSX.readFile('MPS2603-1.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
const sites = new Set();
data.forEach(row => {
    if (row[0]) sites.add(row[0].toString().trim());
});
console.log('Unique Sites in Sheet 0:');
console.log([...sites]);
