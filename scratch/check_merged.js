const XLSX = require('xlsx');
const path = require('path');

const filename = 'MPS2605-2.xlsx';
const filePath = path.join(__dirname, '..', filename);
const wb = XLSX.readFile(filePath);
const prodSheetName = wb.SheetNames.find(name => ['생산배포', '배포용', 'Production'].some(k => name.includes(k))) || wb.SheetNames[0];
const prodWs = wb.Sheets[prodSheetName];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });

console.log('--- Printing rows 238 to 248 ---');
for (let i = 238; i <= 248; i++) {
    console.log(`Row ${i}:`, JSON.stringify(prodRaw[i]));
}

console.log('\n--- Printing rows 58 to 68 ---');
for (let i = 58; i <= 68; i++) {
    console.log(`Row ${i}:`, JSON.stringify(prodRaw[i]));
}
