const fs = require('fs');
const XLSX = require('xlsx');
const extractor = require('./extractor.js');

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    
    // Simulate extractor logic internally to see stats
    const wb = XLSX.read(buf, { type: 'buffer' });
    const masterRaw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
    const masterLookup = {};
    masterRaw.forEach((row, idx) => {
        if (idx <= 0) return; // ignoring header logic complexity
        const m = (row[2] || '').toString().trim();
        if (m) {
            let k = m.toString().toUpperCase().trim();
            if (k.startsWith('VF8')) k = 'VCF850' + k.substring(3);
            if (k.startsWith('M') && !k.startsWith('MYNX')) k = 'MYNX' + k.substring(1);
            k = k.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
            k = k.replace(/III/g, '3').replace(/II/g, '2');
            k = k.split('-')[0];
            k = k.replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '');
            k = k.replace(/[^A-Z0-9]/g, '');
            masterLookup[k] = true;
        }
    });

    const mpsRaw = XLSX.utils.sheet_to_json(wb.Sheets['MPS'], { header: 1 });
    let mapped = 0;
    let unmapped = 0;
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const prod = (row[4] || '').toString().toUpperCase().trim();
        if (!prod || prod.includes('합계')) continue;
        
        let k = prod.split('-')[0].replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').replace(/[^A-Z0-9]/g, '');
        if (masterLookup[k]) mapped++;
        else unmapped++;
    }
    console.log(`Mapped strictly to Master: ${mapped}`);
    console.log(`Unmapped (DNM7550 etc): ${unmapped}`);
}
solve();
