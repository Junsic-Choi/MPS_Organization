const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

try {
    const filePath = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(filePath);
    const mpsWsName = wb.SheetNames.find(n => n.toUpperCase() === 'MPS');
    const ws = wb.Sheets[mpsWsName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    const sample = data.slice(0, 15);
    fs.writeFileSync('mps_col_dump.json', JSON.stringify(sample, null, 2));
    console.log('Dumped to mps_col_dump.json');
} catch (e) {
    fs.writeFileSync('mps_error.txt', e.stack);
}
