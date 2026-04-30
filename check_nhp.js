const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    const sheet0 = wb.Sheets[wb.SheetNames[0]]; 
    const data = XLSX.utils.sheet_to_json(sheet0, { header: 1 });

    let results = [];
    results.push('--- NHP Series Search (Sheet 0) ---');
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row && row[2]) {
            const m = row[2].toString();
            if (m.includes('NHP')) {
                results.push(`Row ${i+1}: "${m}"`);
            }
        }
    }
    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\roman_audit.txt', results.join('\n'), 'utf8');
    console.log('Audit complete. See roman_audit.txt');
} catch (err) {
    console.error(err);
}
