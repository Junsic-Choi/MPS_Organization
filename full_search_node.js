const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    let out = [];

    wb.SheetNames.forEach(sn => {
        const ws = wb.Sheets[sn];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        data.forEach((r, i) => {
            if (r && r.join('|').includes('M544')) {
                out.push(`Sheet: [${sn}] Row: ${i+1} | Data: ${r.join('|')}`);
            }
        });
    });

    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\full_search_utf8.txt', out.join('\n'), 'utf8');
    console.log('Search complete. Found ' + out.length + ' occurrences.');
} catch (err) {
    console.error(err);
}
