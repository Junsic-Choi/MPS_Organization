const XLSX = require('xlsx');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    const wb = XLSX.readFile(file);
    const sheet0 = wb.Sheets[wb.SheetNames[0]]; 
    const data = XLSX.utils.sheet_to_json(sheet0, { header: 1 });

    const targets = ['M544', 'M545', 'DNM45', 'DNM57', 'DNM40'];
    console.log('--- Master Data Audit (Sheet 0) ---');
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row && row[2]) {
            const m = row[2].toString();
            if (targets.some(t => m.includes(t))) {
                console.log(`Row ${i+1}: S:"${row[0]}" | G:"${row[1]}" | M:"${m}" | RPM:"${row[3]}"`);
            }
        }
    }
} catch (err) {
    console.error(err);
}
