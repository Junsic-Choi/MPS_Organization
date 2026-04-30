const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    const sheet0 = wb.Sheets[wb.SheetNames[0]]; 
    const data = XLSX.utils.sheet_to_json(sheet0, { header: 1 });

    let out = [];
    const targets = ['M544', 'M545', 'DNM45', 'DNM57', 'DNM40'];
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row && row[2]) {
            const m = row[2].toString();
            if (targets.some(t => m.includes(t))) {
                out.push(`Row ${i+1}: "${m}"`);
            }
        }
    }
    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\master_full_audit.txt', out.join('\n'), 'utf8');
    console.log('Search complete. Found ' + out.length + ' rows.');
} catch (err) {
    console.error(err);
}
