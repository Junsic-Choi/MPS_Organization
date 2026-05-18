const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['생산배포용'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
console.log('Sample Site Cell:', JSON.stringify(data[5][0]));
console.log('All Month Headers:', JSON.stringify(data[2]));
