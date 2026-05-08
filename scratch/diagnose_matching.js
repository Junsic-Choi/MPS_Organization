const XLSX = require('xlsx');

function getMatchKey(s) {
    if (!s) return '';
    let n = s.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
    n = n.replace(/PUMA|LYNX|MYNX/g, '');
    n = n.replace(/^[MPL]/, ''); 
    return n.replace(/0/g, ''); 
}

const wb = XLSX.readFile('MPS2603-1.xlsx');
const masterWs = wb.Sheets['MPS'];
const masterData = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

const codeMap = {};
masterData.slice(5).forEach(row => {
    const mCode = (row[3] || '').toString().trim();
    const pName = (row[4] || '').toString().trim();
    if (mCode || pName) {
        // [PROBLEM]: pName is "HM1000-F31P-0-K30". getMatchKey returns a very long string.
        // But runningModel is just "HM1000".
        const key = getMatchKey(pName.split('-')[0] || mCode);
        codeMap[key] = { code: mCode, product: pName };
    }
});

const mpsWs = wb.Sheets['생산배포용'];
const mpsRaw = XLSX.utils.sheet_to_json(mpsWs, { header: 1 });

let matchCount = 0;
let failCount = 0;
const samples = [];

mpsRaw.slice(5).forEach(row => {
    const model = row[2];
    if (!model) return;
    const key = getMatchKey(model);
    const mapped = codeMap[key];
    if (mapped && mapped.code) {
        matchCount++;
    } else {
        failCount++;
        if (samples.length < 5) samples.push({ model, key, mapped });
    }
});

console.log('--- Diagnosis Result ---');
console.log('Match Count:', matchCount);
console.log('Fail Count:', failCount);
console.log('Samples of failure:');
samples.forEach(s => console.log(`Model: [${s.model}] -> Key: [${s.key}] (Mapped: ${s.mapped ? 'YES(NoCode)' : 'NO'})`));
