const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    const sheet0 = wb.Sheets[wb.SheetNames[0]]; 
    const data = XLSX.utils.sheet_to_json(sheet0, { header: 1 });

    let out = [];
    out.push('--- MYNX Search in Sheet 0 ---');
    data.forEach((r, i) => {
        const s = JSON.stringify(r);
        if (s.includes('6500') || s.includes('7500') || s.includes('MYNX')) {
            out.push(`Row ${i+1}: ${s}`);
        }
    });

    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\mynx_full_search.txt', out.join('\n'), 'utf8');
    console.log('Done');
} catch (err) {
    console.error(err);
}
