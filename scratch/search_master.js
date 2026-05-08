const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets[wb.SheetNames.find(n => n.toUpperCase() === 'MPS')];
const raw = XLSX.utils.sheet_to_json(ws, {header:1});

const patterns = ['DVF8000/50', 'VCF850', 'SMX2100'];

const searchPatterns = ['DNM750', 'DNM7550'];
console.log('--- DNM750 Series Search in MPS Sheet ---');
raw.forEach((row, i) => {
    const s = (row[4] || '').toString().toUpperCase();
    if (searchPatterns.some(p => s.includes(p))) {
        console.log(`Row ${i} Site[${row[1]}] Prod[${row[4]}] Qty(4월)[${row[11]}] Qty(6월)[${row[21]}]`);
    }
});
