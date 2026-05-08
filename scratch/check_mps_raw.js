const XLSX = require('xlsx');

const mpsFile = 'MPS2604-1.xlsx';
const wb = XLSX.readFile(mpsFile);
const ws = wb.Sheets[wb.SheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

console.log('--- Searching for VCF/VF8/SMX in MPS tab ---');
data.forEach((row, idx) => {
    const rowStr = JSON.stringify(row);
    if (rowStr.includes('VCF') || rowStr.includes('VF8') || rowStr.includes('SMX')) {
        console.log(`Row ${idx + 1}: ${rowStr}`);
    }
});
