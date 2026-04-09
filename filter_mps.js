const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.Sheets[workbook.SheetNames[1]];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    const mpsMonthIdxs = [8, 12, 17, 22, 28, 34];
    
    let totalSum = 0;
    let rowCount = 0;
    
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const code = (row[3] || '').toString().trim();
        const prod = (row[4] || '').toString().trim();
        
        // Filter out sub-total rows: assume they lack code or product name, 
        // or have "Total" in them.
        if (!code || !prod || code.toUpperCase().includes('TOTAL')) continue;
        
        mpsMonthIdxs.forEach(c => {
            totalSum += (parseInt(row[c]) || 0);
        });
        rowCount++;
    }

    let report = `Filtering Result:\nTotal Production Units: ${totalSum}\nValid Product Rows: ${rowCount}\n`;
    fs.writeFileSync('mps_filter_audit.txt', report);
} catch (e) {
    fs.writeFileSync('mps_filter_audit.txt', 'ERROR: ' + e.message);
}
