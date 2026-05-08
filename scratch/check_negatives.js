const XLSX = require('xlsx');
const fs = require('fs');

const files = ['MPS2603-1.xlsx', 'MPS2604-1.xlsx'];
files.forEach(filename => {
    if (!fs.existsSync(filename)) {
        console.log(`${filename} not found`);
        return;
    }

    const wb = XLSX.readFile(filename);
    const sheetNames = wb.SheetNames;
    const masterWs = wb.Sheets[sheetNames.find(n => n.toUpperCase() === 'MPS') || 'MPS'];

    if (masterWs) {
        const raw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });
        console.log(`\nScanning ${filename} MPS sheet (${raw.length} rows)...`);
        
        for (let r = 0; r < raw.length; r++) {
            const row = raw[r] || [];
            row.forEach((cell, c) => {
                const val = (cell || '').toString().trim();
                if (val.startsWith('-') && !isNaN(parseFloat(val))) {
                    console.log(`Row ${r+1}, Col ${c+1}: Value ${cell}, Code: ${row[3]}, Product: ${row[4]}`);
                }
            });
        }
    } else {
        console.log(`${filename} MPS sheet not found`);
    }
});
