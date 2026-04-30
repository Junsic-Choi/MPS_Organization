const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    const sheet0 = wb.Sheets[wb.SheetNames[0]]; 
    const data = XLSX.utils.sheet_to_json(sheet0, { header: 1 });

    const targets = ['DC428', 'DBC13', 'VT110', 'VTR121', 'DBC11', 'DC371', 'VT11M'];
    let out = [];
    out.push('--- Master Data Audit for Failing Items ---');
    data.forEach((r, i) => {
        const rowStr = JSON.stringify(r);
        if (targets.some(t => rowStr.includes(t))) {
            out.push(`Row ${i+1}: G:"${r[1]}" | M:"${r[2]}" | RPM:"${r[3]}"`);
        }
    });

    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\audit_failing_items.txt', out.join('\n'), 'utf8');
    console.log('Done');
} catch (err) {
    console.error(err);
}
