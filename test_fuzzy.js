const fs = require('fs');
const XLSX = require('xlsx');

function getPatternGroup(p) {
    const up = (p || '').toString().toUpperCase();
    if (up.includes('DBC') || up.includes('DBD') || up.includes('DBM')) return '34. BORING';
    if (up.includes('VCF') || up.includes('VF8')) return '17. VCF 850 Series';
    if (up.includes('DEM')) return '13-3. DEM4000';
    if (up.includes('DNM') && up.includes('5AX')) return '10. DNM 5AX Series';
    if (up.includes('DNM')) return '13. DNM Series';
    if (up.includes('MYNX') || up.includes('M65') || up.includes('M75')) return '12. MYNX Series';
    if (up.includes('SMX')) return '28. SMX Series';
    if (up.includes('P4100')) return '25. P4100 Series';
    if (up.includes('LYNX')) return '20. LYNX Series';
    return '';
}

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type: 'buffer'});
    const mst = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1});
    const mps = XLSX.utils.sheet_to_json(wb.Sheets['MPS'], {header:1});

    // 1. Collect all Master models per group
    const masterGroups = {};
    for(let r=3; r<mst.length; r++) {
        const row = mst[r] || [];
        const g = (row[1] || '').toString().trim();
        const m = (row[2] || '').toString().trim();
        if (g && m) {
            if(!masterGroups[g]) masterGroups[g] = new Set();
            masterGroups[g].add(m);
        }
    }

    // 2. Sample Unmapped from MPS
    const masterLookupString = {};
    for(let k in masterGroups) {
        masterGroups[k].forEach(m => {
            const key = m.toUpperCase().replace(/VIII/g, '8').replace(/VII/g, '7').replace(/II/g, '2').split('-')[0].replace(/[^A-Z0-9]/g, '');
            masterLookupString[key] = { group: k, model: m };
        });
    }

    const unmappedProds = new Set();
    for(let r=5; r<mps.length; r++) {
        const prod = (mps[r][4]||'').toString().trim().toUpperCase();
        if(!prod || prod.includes('합계')) continue;
        const key = prod.split('-')[0].replace(/PUMA|LYNX/g, '').replace(/^P|^L/, '').replace(/[^A-Z0-9]/g, '');
        if(!masterLookupString[key]) {
            unmappedProds.add(prod);
        }
    }

    // 3. Try to fuzzy match some
    console.log(`Total Unique Unmapped Prods (by string): ${unmappedProds.size}`);
    let i=0;
    unmappedProds.forEach(prod => {
        if(i++ < 20) {
            const group = getPatternGroup(prod);
            if(group && masterGroups[group]) {
                const candidates = Array.from(masterGroups[group]);
                // Very crude fuzzy: longest common prefix matching or character inclusion
                let bestMatch = '';
                let bestScore = -1;
                
                const pClean = prod.split('-')[0].replace(/[^A-Z0-9]/g, '');
                candidates.forEach(c => {
                    const cClean = c.toUpperCase().replace(/[^A-Z0-9]/g, '');
                    let score = 0;
                    for(let x=0; x<Math.min(pClean.length, cClean.length); x++) {
                        if(pClean[x] === cClean[x]) score++;
                        else break; // Prefix score
                    }
                    if(score > bestScore) { bestScore = score; bestMatch = c; }
                });
                console.log(`PROD: ${prod.padStart(30)} => Best in ${group}: ${bestMatch}`);
            } else {
                console.log(`PROD: ${prod.padStart(30)} => No Fallback Group / Candidates`);
            }
        }
    });

}
solve();
