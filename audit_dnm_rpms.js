const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const mst = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1});
    
    const rpms = new Set();
    for(let r=3; r<mst.length; r++) {
        const row = mst[r] || [];
        const group = (row[1] || '').toString();
        const model = (row[2] || '').toString();
        const rpm = (row[3] || '').toString();
        if (group.includes('DNM') || model.includes('DNM')) {
            if (rpm) rpms.add(rpm);
        }
    }
    console.log("DNM Series RPMs in Master:", Array.from(rpms).join(', '));
}
solve();
