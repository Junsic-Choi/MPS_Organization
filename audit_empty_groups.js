const fs = require('fs');
const XLSX = require('xlsx');

function getPatternGroup(p) {
    const up = (p || '').toString().toUpperCase();
    if (up.includes('DBC') || up.includes('DBD') || up.includes('DBM')) return '34. BORING';
    if (up.includes('VCF') || up.includes('VF8')) return '17. VCF 850 Series';
    if (up.includes('VTR')) return '33. VTR Series';
    if (up.includes('VT') || up.includes('VTL')) return '30. PV/VT Series';
    if (up.includes('MYNX') || up.includes('M65') || up.includes('M75')) return '12. MYNX Series';
    if (up.includes('DNM') || up.includes('DEM')) return '13. DNM/DEM Series';
    if (up.includes('DVF')) return '11. DVF Series';
    if (up.includes('SMX')) return '21. SMX Series';
    if (up.includes('DC325') || up.startsWith('DC')) return '35. DC Series';
    if (up.includes('PUMA')) return 'PUMA Series';
    if (up.includes('LYNX')) return 'LYNX Series';
    return '';
}

function solve() {
    const buf = fs.readFileSync('MPS2603-1.xlsx');
    const wb = XLSX.read(buf, {type:'buffer'});
    const mps = XLSX.utils.sheet_to_json(wb.Sheets['MPS'], {header:1});
    
    let emptyGroups = new Set();
    let runningGroup = '';
    
    for(let r=5; r<mps.length; r++) {
        const row = mps[r] || [];
        const groupCode = (row[2] || '').toString().trim();
        if (groupCode) runningGroup = groupCode;
        
        const mCode = (row[3] || '').toString().trim();
        const prod = (row[4] || '').toString().trim();
        if (!mCode || !prod || mCode.includes('계') || prod.includes('합계')) continue;
        
        let g = runningGroup;
        // In extractor.js:
        // const finalGroup = info ? info.group : (getPatternGroup(prod) || runningGroup);
        // Wait, for unusedData it's group: getPatternGroup(prod) || runningGroup.
        // Wait, if runningGroup is empty AND getPatternGroup is empty!
        if (!g && !getPatternGroup(prod)) {
            emptyGroups.add(prod.split('-')[0]);
        }
    }
    
    console.log("Empty Group Products:", Array.from(emptyGroups));
}
solve();
