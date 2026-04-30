const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    
    const mstWs = wb.Sheets[wb.SheetNames[0]]; // Master ("생산배포용")
    const masterRaw = XLSX.utils.sheet_to_json(mstWs, {header:1});
    
    console.log('--- Master Sheet Rows for DNM750L II ---');
    for (let r=3; r<masterRaw.length; r++) {
        let m = (masterRaw[r][2] || '').toString().trim();
        if (m.includes('DNM750L')) {
            console.log(`Row ${r}: Group=${masterRaw[r][1]}, Model=${masterRaw[r][2]}, RPM=${masterRaw[r][3]}`);
        }
        if (m.includes('DNM7550')) {
            console.log(`Found DNM7550 in Master! Row ${r}: ${masterRaw[r]}`);
        }
    }
}
solve();
