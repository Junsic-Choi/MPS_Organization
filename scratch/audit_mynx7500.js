const { processMpsFile } = require('c:/Users/i0215099/Desktop/MPS_UPDATE/extractor');
const XLSX = require('xlsx');
const fs = require('fs');

function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().trim().split('-')[0];
    n = n.replace(/PUMA|LYNX/g, '').trim();
    n = n.replace(/^P|^L/, '');
    n = n.replace(/VIII/g, '8').replace(/VII/g, '7').replace(/VI/g, '6').replace(/IV/g, '4').replace(/IX/g, '9');
    n = n.replace(/III/g, '3').replace(/II/g, '2');
    return n.replace(/[^A-Z1-9]/g, '');
}

async function audit() {
    const filename = 'MPS2604-1.xlsx';
    const wb = XLSX.readFile(filename);
    const sheetNames = wb.SheetNames;
    const masterWs = wb.Sheets[sheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS'];
    const raw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
    
    console.log('Searching for "VTR" with "M" in Master Plan:');
    let count = 0;
    for (let r = 5; r < raw.length; r++) {
        const row = raw[r] || [];
        const pName = (row[4] || '').toString();
        const mCode = (row[3] || '').toString();
        if ((pName.includes('VTR') && pName.includes('M')) || (mCode.includes('VTR') && mCode.includes('M'))) {
            console.log(`Row ${r+1}: [${mCode}] "${pName}" -> Key: ${getMatchKey(pName || mCode)}`);
            if (++count > 20) break;
        }
    }
}

audit();
