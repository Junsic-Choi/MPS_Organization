const XLSX = require('xlsx');

const wb = XLSX.readFile('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
const prodWs = wb.Sheets[wb.SheetNames[0]];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });

console.log('=== UNIQUE GROUPS IN PRODUCTION SHEET ===');
const groups = new Set();
prodRaw.forEach((row, idx) => {
    if (idx < 2) return;
    const group = row[1];
    if (group) groups.add(String(group).trim());
});
console.log(Array.from(groups));
