const XLSX = require('xlsx');
const fs = require('fs');

const file = 'MPS2605-1.xlsx';
const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
const masterWs = wb.Sheets['MPS'] || wb.Sheets[wb.SheetNames[1]];
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log("Checking all models for PL I0215001 in Master Plan...");
const models = new Set();
masterRaw.forEach((row, idx) => {
    if (idx < 5) return;
    const pl = (row[1] || '').toString().trim();
    const product = (row[4] || '').toString().trim();
    
    if (pl === 'I0215001') {
        models.add(product.split('-')[0]);
    }
});

console.log("PL I0215001 products:", Array.from(models));
