const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const data = XLSX.utils.sheet_to_json(wb.Sheets['MPS'], {header:1});

    console.log('--- Searching for DNM3 or 5AX in MPS --');
    for(let r=0; r<data.length; r++) {
        const row = data[r] || [];
        const prod = (row[4]||'').toString().toUpperCase();
        if (prod.includes('DNM3') || prod.includes('5AX')) {
            console.log(`Row ${r}: prod=${prod}, modelConfig=${row[3]}`);
        }
    }
}
solve();
