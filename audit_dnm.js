const fs = require('fs');
const XLSX = require('xlsx');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const masterWs = wb.Sheets[wb.SheetNames[0]];
    const masterData = XLSX.utils.sheet_to_json(masterWs, {header:1});

    console.log('--- Master Sheet Audit for DNM/DEM ---');
    masterData.forEach((row, i) => {
        const rowStr = JSON.stringify(row);
        if (rowStr.includes('DNM') || rowStr.includes('DEM')) {
            console.log(`Row ${i}:`, row[0], '|', row[1], '|', row[2]);
        }
    });

    console.log('\n--- MPS Sheet Audit for DNM/DEM (Site Codes) ---');
    const mpsWs = wb.Sheets['MPS'];
    const mpsData = XLSX.utils.sheet_to_json(mpsWs, {header:1});
    let runningSite = '';
    for(let i=5; i<mpsData.length; i++) {
        const row = mpsData[i];
        if(!row) continue;
        if(row[6]) runningSite = row[6];
        const prod = (row[4]||'').toString();
        if(prod.includes('DNM755L') || prod.includes('DNM355A')) {
            console.log(`MPS Row ${i}: Prod=${prod}, SiteCode=${row[6]}, InheritedSite=${runningSite}`);
        }
    }
}
solve();
