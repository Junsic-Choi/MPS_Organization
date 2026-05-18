const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
for(let i=210; i<230; i++) {
    console.log(`Row ${i+1}: ${JSON.stringify(data[i])}`);
}
