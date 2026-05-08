const XLSX = require('xlsx');
const fs = require('fs');
const wb = XLSX.readFile('MPS2604-1.xlsx');
const sheetNames = wb.SheetNames;
let output = 'Sheet Names: ' + sheetNames.join(', ') + '\n';

// Try to find the production sheet
const mpsWsName = sheetNames.find(n => n.includes('배포') || n.includes('분석')) || '생산배포용';
const mpsWs = wb.Sheets[mpsWsName];
if (mpsWs) {
    const raw = XLSX.utils.sheet_to_json(mpsWs, { header: 1 });
    output += `\n--- [${mpsWsName}] First 10 rows ---\n`;
    for (let i = 0; i < Math.min(10, raw.length); i++) {
        const row = raw[i] || [];
        output += `Row ${i}: ` + row.map(c => (c === undefined || c === null ? '' : c).toString().trim()).join(' | ') + '\n';
    }
}

// Try to find the MPS sheet
const masterWsName = sheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS';
const masterWs = wb.Sheets[masterWsName];
if (masterWs) {
    const raw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
    output += `\n--- [${masterWsName}] First 10 rows ---\n`;
    for (let i = 0; i < Math.min(10, raw.length); i++) {
        const row = raw[i] || [];
        output += `Row ${i}: ` + row.map(c => (c === undefined || c === null ? '' : c).toString().trim()).join(' | ') + '\n';
    }
}

fs.writeFileSync('scratch/inspect_results.txt', output);
console.log('Results written to scratch/inspect_results.txt');
