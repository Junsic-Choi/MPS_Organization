const XLSX = require('xlsx');
const fs = require('fs');

function normRPM(r) {
    if (!r) return "";
    let s = r.toString().toUpperCase().replace(/\s+/g, '');
    if (s.endsWith('K')) return (parseFloat(s) * 1000).toString();
    return s.replace(/[^0-9]/g, '');
}

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.Sheets[workbook.SheetNames.find(n => n.toUpperCase().includes('MPS')) || workbook.SheetNames[1]];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    
    // Month indices in Sheet 1: 2월(Index 8), 3월(12), 4월(17), 5월(22), 6월(28), 7월(34)
    const mpsMonthIdxs = [8, 12, 17, 22, 28, 34];
    
    const mpsPool = [];
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const pl = (row[1] || '').toString().trim();
        const code = (row[3] || '').toString().trim();
        const prod = (row[4] || '').toString().trim();
        const rpm = (row[6] || '').toString().trim();
        
        if (code && prod) {
            const qtys = mpsMonthIdxs.map(idx => parseInt(row[idx]) || 0);
            mpsPool.push({ pl, code, prod, rpm: normRPM(rpm), qtys });
        }
    }

    const prodSheet = workbook.Sheets[workbook.SheetNames.find(n => n.includes('배포')) || workbook.SheetNames[0]];
    const prodRaw = XLSX.utils.sheet_to_json(prodSheet, { header: 1 });
    
    const siteMap = {}; // Site -> Set of PLs found in corresponding rows
    let lastSite = "";
    for (let r = 6; r < prodRaw.length; r++) {
        const row = prodRaw[r] || [];
        if (row[0]) lastSite = row[0].toString().trim();
        if (!lastSite || lastSite.includes('총합계')) continue;
        
        const model = row[2] ? row[2].toString().trim() : "";
        // Just checking Henex for now
        if (lastSite.includes('헤넥스')) {
             console.log(`Henex Row ${r}: Model=${model}, RPM=${row[3]}`);
        }
    }

    fs.writeFileSync('mps_pool_sample.json', JSON.stringify(mpsPool.slice(0, 50), null, 2));
    console.log(`MPS Pool size: ${mpsPool.length}`);

} catch (e) {
    console.error(e);
}
