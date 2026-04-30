const fs = require('fs');
const XLSX = require('xlsx');

function audit() {
    const filename = 'MPS2604-1.xlsx';
    if (!fs.existsSync(filename)) {
        console.log(`${filename} not found.`);
        return;
    }
    const buf = fs.readFileSync(filename);
    const wb = XLSX.read(buf, { type: 'buffer' });
    
    // 1. Master Map
    const masterRaw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
    const masterLookup = {};
    masterRaw.forEach((row, idx) => {
        if (idx < 5) return;
        const m = (row[2] || '').toString().toUpperCase().trim();
        if (m) {
            let k = m.split('-')[0].replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').replace(/[^A-Z0-9]/g, '');
            masterLookup[k] = m;
        }
    });

    // 2. MPS Sheet
    const mpsWs = wb.Sheets['MPS'] || wb.Sheets['mps'];
    if (!mpsWs) {
        console.log('MPS sheet not found');
        return;
    }
    const mpsRaw = XLSX.utils.sheet_to_json(mpsWs, { header: 1 });
    const unmapped = {};
    let totalUnmapped = 0;

    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const prod = (row[4] || '').toString().toUpperCase().trim();
        if (!prod || prod.includes('합계')) continue;
        
        let k = prod.split('-')[0].replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').replace(/[^A-Z0-9]/g, '');
        
        if (!masterLookup[k]) {
            // Check if there is any qty in the month columns
            let hasQty = false;
            for (let c = 8; c <= 40; c++) {
                if (parseInt(row[c]) > 0) { hasQty = true; break; }
            }
            if (hasQty) {
                unmapped[prod] = (unmapped[prod] || 0) + 1;
                totalUnmapped++;
            }
        }
    }

    console.log(`Total Unmapped Units: ${totalUnmapped}`);
    console.log(`Distinct Unmapped Products: ${Object.keys(unmapped).length}`);
    
    const sorted = Object.entries(unmapped).sort((a, b) => b[1] - a[1]);
    console.log('--- Top 20 Unmapped ---');
    sorted.slice(0, 20).forEach(([name, count]) => {
        console.log(`${name}: ${count}`);
    });
}
audit();
