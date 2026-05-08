const XLSX = require('xlsx');
const wb = XLSX.readFile('MPS2604-1.xlsx');
const ws = wb.Sheets[wb.SheetNames.find(n => n.toUpperCase() === 'MPS')];
const raw = XLSX.utils.sheet_to_json(ws, {header:1});

const targets = ['DVF8000', 'VCF850', 'SMX2100', 'MYNX6500'];

targets.forEach(t => {
    console.log(`\n--- SEARCHING FOR [${t}] ---`);
    raw.forEach((row, i) => {
        const pName = (row[4] || '').toString().toUpperCase();
        if (pName.includes(t)) {
            console.log(`Row ${i} Site[${row[1]}] Prod[${row[4]}] Qty(4월)[${row[11]}] Qty(6월)[${row[21]}]`);
        }
    });
});
