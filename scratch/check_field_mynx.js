const XLSX = require('xlsx');
const fs = require('fs');

async function checkField() {
    const filename = 'MPS2603-1.xlsx';
    const wb = XLSX.readFile(filename);
    const sheetNames = wb.SheetNames;
    const wsName = sheetNames.find(n => n.includes('배포') || n.includes('분석')) || '생산배포용';
    const ws = wb.Sheets[wsName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    console.log(`Checking ${wsName} for MYNX:`);
    let count = 0;
    for (let r = 0; r < raw.length; r++) {
        const row = raw[r] || [];
        const content = JSON.stringify(row);
        if (content.includes('MYNX')) {
            console.log(`Row ${r+1}: ${content}`);
            if (++count > 20) break;
        }
    }
}

checkField();
