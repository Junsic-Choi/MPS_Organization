const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const mstWs = wb.Sheets[wb.SheetNames[0]]; // Master ("생산배포용")
    const raw = XLSX.utils.sheet_to_json(mstWs, {header:1});

    let groups = new Set();
    for (let r=3; r<raw.length; r++) {
        let g = (raw[r][1] || '').toString().trim();
        if (g) groups.add(g);
    }
    console.log('--- Groups in Master Sheet ---');
    console.log(Array.from(groups).sort());
}

solve();
