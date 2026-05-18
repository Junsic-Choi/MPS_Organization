const XLSX = require('xlsx');
const fs = require('fs');

const file = 'MPS2605-1.xlsx';
const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
const masterWs = wb.Sheets['MPS'] || wb.Sheets[wb.SheetNames[1]];
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log("Checking all models for Site 1840 and PL I0215001 in Master Plan...");
const sites = {};
masterRaw.forEach((row, idx) => {
    if (idx < 5) return;
    const pl = (row[1] || '').toString().trim();
    const site = (row[6] || '').toString().trim();
    const model = (row[3] || row[4] || '').toString().trim();
    
    if (site === '1840') {
        if (!sites['1840']) sites['1840'] = new Set();
        sites['1840'].add(model.split('-')[0]);
    }
    if (pl === 'I0215001') {
        if (!sites['I0215001']) sites['I0215001'] = new Set();
        sites['I0215001'].add(model.split('-')[0]);
    }
});

console.log("Site 1840 models:", Array.from(sites['1840'] || []));
console.log("PL I0215001 models:", Array.from(sites['I0215001'] || []));
