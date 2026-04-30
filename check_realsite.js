const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\Real site.xlsx';
    if (!fs.existsSync(file)) {
        console.log('Real site.xlsx does not exist.');
        process.exit(0);
    }
    const wb = XLSX.readFile(file);
    const ws = wb.Sheets[wb.SheetNames[0]]; 
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

    let out = [];
    out.push('--- Real site.xlsx Content ---');
    data.forEach((r, i) => {
        if (r && r.length > 0) {
            out.push(`Row ${i+1}: ${r.join('|')}`);
        }
    });
    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\realsite_audit.txt', out.join('\n'), 'utf8');
    console.log('Audit complete.');
} catch (err) {
    console.error(err);
}
