const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    const sn = wb.SheetNames;
    const mpsWsName = sn.find(n => n.toUpperCase() === 'MPS') || sn[1];
    const mpsWs = wb.Sheets[mpsWsName]; 
    const data = XLSX.utils.sheet_to_json(mpsWs, { header: 1 });

    let results = [];
    results.push(`--- NHP Search in ${mpsWsName} ---`);
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row && row[4]) { // Prod Name column
            const p = row[4].toString();
            if (p.includes('NHP')) {
                results.push(`Row ${i+1}: "${p}"`);
            }
        }
    }
    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\roman_audit_mps.txt', results.join('\n'), 'utf8');
} catch (err) {
    console.error(err);
}
