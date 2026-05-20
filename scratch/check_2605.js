const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2605-1.xlsx');
console.log('Sheet Names:', wb.SheetNames);

const prodSheet = wb.SheetNames.find(name => ['생산배포', '배포용', 'Production'].some(k => name.includes(k))) || wb.SheetNames[0];
console.log('Selected Sheet:', prodSheet);

const ws = wb.Sheets[prodSheet];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
console.log('Row count:', raw.length);
console.log('First 5 rows:');
for (let i = 0; i < Math.min(15, raw.length); i++) {
    console.log(`Row ${i}:`, raw[i] ? raw[i].slice(0, 15) : []);
}
