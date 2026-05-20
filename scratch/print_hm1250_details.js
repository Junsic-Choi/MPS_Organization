const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const ws = wb.Sheets['MPS'];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

console.log('Month Row 2:', raw[2].slice(0, 30));
console.log('Type Row 4:', raw[4].slice(0, 30));
console.log('Row 374 (HM1250):', raw[374]);
console.log('Row 377 (HM1250W):', raw[377]);
