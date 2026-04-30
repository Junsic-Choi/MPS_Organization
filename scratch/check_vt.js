const fs = require('fs');
const XLSX = require('xlsx');

function checkVTL() {
    const buf = fs.readFileSync('MPS2604-1.xlsx');
    const wb = XLSX.read(buf, { type: 'buffer' });
    const masterRaw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
    
    console.log('--- Master Sheet Models containing VT ---');
    masterRaw.forEach((row, idx) => {
        const m = (row[2] || '').toString();
        if (m.toUpperCase().includes('VT')) {
            console.log(`Row ${idx}: ${m}`);
        }
    });

    const mpsRaw = XLSX.utils.sheet_to_json(wb.Sheets['MPS'], { header: 1 });
    console.log('\n--- MPS Sheet Models containing VT (Unmapped) ---');
    const unmappedVT = {};
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const prod = (row[4] || '').toString();
        if (prod.toUpperCase().includes('VT')) {
             // Check if it's unmapped... (simplified check here)
             unmappedVT[prod] = (unmappedVT[prod] || 0) + 1;
        }
    }
    console.log(JSON.stringify(unmappedVT, null, 2));
}
checkVTL();
