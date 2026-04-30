const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
try {
    const wb = XLSX.readFile(FILE_PATH, { bookSheets: true });
    console.log('SheetNames:', wb.SheetNames);
    
    // MPS 시트 찾기
    const mpsName = wb.SheetNames.find(n => n.toUpperCase().includes('MPS'));
    if (!mpsName) throw new Error('MPS sheet not found');
    
    const workbook = XLSX.readFile(FILE_PATH, { sheets: [mpsName] });
    const sheet = workbook.Sheets[mpsName];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    let stats = [];
    raw.forEach((row, i) => {
        const code = (row[3] || '').toString();
        const prod = (row[4] || '').toString();
        if (code && prod) {
            stats.push({ row: i+1, code, prod });
        }
    });
    
    fs.writeFileSync('mps_diag.json', JSON.stringify(stats, null, 2));
    console.log('Found ' + stats.length + ' rows with code and prod in MPS sheet');
} catch (e) {
    console.error(e.message);
}
