const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join('c:', 'Users', 'i0215099', 'Desktop', 'MPS_UPDATE', 'MPS2603-1.xlsx');
const wb = XLSX.readFile(filePath);
const mpsWsName = wb.SheetNames.find(n => n.toUpperCase() === 'MPS');
const ws = wb.Sheets[mpsWsName];
const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

console.log('--- MPS SHEET TOP 10 ROWS ---');
data.slice(0, 10).forEach((row, i) => {
    console.log(`Row ${i}: `, row.map((cell, idx) => `[Col ${idx}:${String.fromCharCode(65+idx)}] ${cell}`).join(' | '));
});
