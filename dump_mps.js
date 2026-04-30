const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const mpsWs = wb.Sheets['MPS'];
    const data = XLSX.utils.sheet_to_json(mpsWs, {header:1});

    console.log('--- MPS Row Dump (First 10) ---');
    for(let i=0; i<15; i++) {
        console.log(`Row ${i}:`, JSON.stringify(data[i]));
    }
}
solve();
