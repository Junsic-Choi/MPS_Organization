const fs = require('fs');
const XLSX = require('xlsx');

function checkMaster() {
    const buf = fs.readFileSync('MPS2604-1.xlsx');
    const wb = XLSX.read(buf, { type: 'buffer' });
    const masterRaw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
    
    const targets = ['DBC130', 'VCF850', 'MYNX', 'NHP6300'];
    console.log('--- Master Plan Search ---');
    masterRaw.forEach((row, idx) => {
        const m = (row[2] || '').toString();
        targets.forEach(t => {
            if (m.toUpperCase().includes(t)) {
                console.log(`Row ${idx}: ${m} (Group: ${row[1]})`);
            }
        });
    });
}
checkMaster();
