const XLSX = require('xlsx');

function checkProduction(filename, site, mIdx) {
    console.log(`\nChecking Production for ${site} at Col ${mIdx}...`);
    const wb = XLSX.readFile(filename);
    const ws = wb.Sheets['생산배포용'];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    let runningSite = '';
    data.forEach((row, idx) => {
        if (idx < 5) return;
        let s = (row[0] || '').toString().replace(/^\d+\.\s*/, '').trim();
        if (s) runningSite = s;
        
        if (runningSite === site) {
            const q = parseInt(row[mIdx]) || 0;
            if (q > 0) {
                console.log(`Row ${idx+1}: ${row[2]} (Qty: ${q})`);
            }
        }
    });
}

checkProduction('MPS2605-1.xlsx', '성주', 12); // 12 is 9월
