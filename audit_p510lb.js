const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const mst = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1});
    
    console.log("Searching Master sheet for P510LB...");
    for(let r=0; r<mst.length; r++) {
        const rowStr = JSON.stringify(mst[r]);
        if (rowStr.includes('510L')) {
            console.log(`Row ${r}:`, rowStr);
        }
    }
}
solve();
