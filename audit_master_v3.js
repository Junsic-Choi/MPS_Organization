const XLSX = require('xlsx');
const fs = require('fs');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    const sheet0 = wb.Sheets[wb.SheetNames[0]]; 
    const data = XLSX.utils.sheet_to_json(sheet0, { header: 1 });

    let out = [];
    for (let i = 6; i < 150; i++) {
        const row = data[i];
        if (row && row[2]) {
            out.push(`Row ${i+1}: "${row[2]}" (Col B: "${row[1]}", Col A: "${row[0]}")`);
        }
    }
    fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\master_raw_utf8.txt', out.join('\n'), 'utf8');
    console.log('Done');
} catch (err) {
    console.error(err);
}
