const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['MPS'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
console.log('Row 4:', raw[4]);
