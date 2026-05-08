const XLSX = require('xlsx');
const fs = require('fs');

function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
    n = n.replace(/PUMA|LYNX|MYNX/g, '');
    n = n.replace(/^[MPL]/, ''); 
    return n.replace(/0/g, ''); 
}

const filename = 'MPS2603-1.xlsx';
if (!fs.existsSync(filename)) {
    console.error('File not found');
    process.exit(1);
}

const wb = XLSX.readFile(filename);
console.log('Sheet Names:', wb.SheetNames);

const masterWsName = wb.SheetNames.find(n => n.toUpperCase() === 'MPS');
console.log('Master WS Name:', masterWsName);

if (masterWsName) {
    const masterWs = wb.Sheets[masterWsName];
    const masterData = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
    console.log('Master Data Sample (Row 5-10):');
    masterData.slice(0, 10).forEach((row, i) => {
        console.log(`Row ${i}:`, row.slice(0, 6));
    });
    
    const codeMap = {};
    masterData.slice(5).forEach((row, i) => {
        const mCode = (row[3] || '').toString().trim();
        const pName = (row[4] || '').toString().trim();
        if (mCode || pName) {
            const key = getMatchKey(pName || mCode);
            if (!codeMap[key]) {
                codeMap[key] = { code: mCode, product: pName };
            }
        }
    });
    console.log('CodeMap Sample Keys:', Object.keys(codeMap).slice(0, 10));
}

const mpsWsName = wb.SheetNames.find(n => n.includes('배포')) || '생산배포용';
console.log('MPS WS Name:', mpsWsName);
if (mpsWsName) {
    const mpsWs = wb.Sheets[mpsWsName];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsWs, { header: 1 });
    console.log('MPS Raw Sample (Row 0-10):');
    mpsRaw.slice(0, 10).forEach((row, i) => {
        console.log(`Row ${i}:`, row.slice(0, 5));
    });
}
