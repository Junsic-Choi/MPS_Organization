const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
console.log('Sheet Names:', wb.SheetNames);

const prodWs = wb.Sheets[wb.SheetNames[0]];
const masterWs = wb.Sheets[wb.SheetNames[1]];

const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log('\n=== FIRST 20 ROWS OF SHEET 0 (Production/배포용) ===');
for (let i = 0; i < 25; i++) {
    if (prodRaw[i]) {
        console.log(`Row ${i+1}:`, prodRaw[i].slice(0, 10).map(x => String(x).trim()));
    }
}

console.log('\n=== FIRST 20 ROWS OF SHEET 1 (Master/MPS) ===');
for (let i = 0; i < 25; i++) {
    if (masterRaw[i]) {
        console.log(`Row ${i+1}:`, masterRaw[i].slice(0, 10).map(x => String(x).trim()));
    }
}
