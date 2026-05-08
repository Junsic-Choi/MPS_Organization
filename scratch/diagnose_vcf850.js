const XLSX = require('xlsx');
const fs = require('fs');

function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim().split('-')[0];
    
    // [SAFE ENHANCEMENT]: 접두사 표준화 (M/MYNX, V/VCF/VF, D/DNM 통합)
    if (n.startsWith('MYNX')) n = 'M' + n.substring(4);
    if (n.startsWith('VCF')) n = 'V' + n.substring(3);
    if (n.startsWith('VF') && !n.startsWith('VFC')) n = 'V' + n.substring(2);
    if (n.startsWith('DNM')) n = 'D' + n.substring(3);
    if (n.startsWith('NHP')) n = 'NH' + n.substring(3);
    if (n.startsWith('NHC')) n = 'NH' + n.substring(3);

    n = n.replace(/PUMA|LYNX/g, '').trim();
    n = n.replace(/^P|^L/, '');
    
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');
    
    return n.replace(/[^A-Z1-9]/g, '');
}

async function diagnose() {
    const filename = 'MPS2603-1.xlsx';
    const wb = XLSX.readFile(filename);
    const mpsWs = wb.Sheets[wb.SheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS'];
    const masterData = XLSX.utils.sheet_to_json(mpsWs, { header: 1 });

    console.log('--- Analyzing VCF850 ---');
    console.log('Field Model: VCF850LSR -> Key: ' + getMatchKey('VCF850LSR'));
    console.log('Field Model: VCF850SR -> Key: ' + getMatchKey('VCF850SR'));
    
    const targets = ['XLB', 'LSB', 'VT1100M'];
    console.log(`\nSearching Master Plan for EXACT matches (${targets.join(', ')}):`);
    for (let r = 0; r < masterData.length; r++) {
        const row = masterData[r] || [];
        const pName = (row[4] || '').toString();
        const mCode = (row[3] || '').toString();
        if (targets.some(t => pName.includes(t) || mCode.includes(t))) {
            console.log(`Row ${r+1}: [${row[1]}] [${mCode}] "${pName}" -> Key: ${getMatchKey(pName || mCode)}`);
        }
    }
}

diagnose();
