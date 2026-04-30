const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
try {
    const wb = XLSX.readFile(FILE_PATH);
    const mpsName = wb.SheetNames.find(n => n.toUpperCase().includes('MPS'));
    const ws = wb.Sheets[mpsName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
    
    // 처음 10개 행 출력
    console.log('--- MPS Sheet Header Structure ---');
    for (let r = 0; r < 10; r++) {
        console.log(`Row ${r}:`, (raw[r] || []).slice(0, 40).map((v, i) => `[${i}]${v || ''}`));
    }
} catch (e) {
    console.error(e.message);
}
