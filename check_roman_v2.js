const XLSX = require('xlsx');
const path = require('path');

try {
    const file = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2603-1.xlsx';
    console.log('Reading file:', file);
    const wb = XLSX.readFile(file);
    const sheet0 = wb.Sheets[wb.SheetNames[0]]; 
    const data = XLSX.utils.sheet_to_json(sheet0, { header: 1 });

    console.log('--- [Sheet 0] Row 7-100 Model Check ---');
    for (let i = 6; i < 100; i++) {
        const row = data[i];
        if (row && row[2]) {
            console.log(`Row ${i+1}: "${row[2]}"`);
        }
    }
} catch (err) {
    console.error('Error:', err.message);
}
