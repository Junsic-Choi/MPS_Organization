const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('MPS2603-1.xlsx');
    const mpsSheet = workbook.Sheets[workbook.SheetNames[1]];
    const mpsRaw = XLSX.utils.sheet_to_json(mpsSheet, { header: 1 });
    const mpsMonthIdxs = [8, 12, 17, 22, 28, 34];
    const monthNames = ["2월", "3월", "4월", "5월", "6월", "7월"];
    
    let log = "Searching for DNC 8060 production in Sheet 1 (MPS):\n";
    for (let r = 5; r < mpsRaw.length; r++) {
        const row = mpsRaw[r] || [];
        const prod = (row[4] || '').toString().toUpperCase();
        if (prod.includes('DNC8060') || prod.includes('DNC 8060')) {
            const qtys = mpsMonthIdxs.map((idx, i) => `${monthNames[i]}: ${row[idx]||0}`).join(', ');
            log += `Row ${r}: Product=${row[4]}, Code=${row[3]}, Qtys=[${qtys}]\n`;
        }
    }

    fs.writeFileSync('dnc_mps_audit.txt', log);
} catch (e) {
    fs.writeFileSync('dnc_mps_audit.txt', 'ERROR: ' + e.message);
}
