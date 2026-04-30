const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const mst = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1});
    
    // Look for DNM75L2, DNM7550 or F/S/H flags in Master sheet columns
    console.log("Searching Master sheet for DNM75L2 or DNM7550...");
    for(let r=0; r<mst.length; r++) {
        const rowStr = JSON.stringify(mst[r]);
        if (rowStr.includes('75L2') || rowStr.includes('7550')) {
            console.log(`Row ${r}: ${rowStr}`);
        }
    }
}
solve();
