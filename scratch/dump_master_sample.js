const XLSX = require('xlsx');
const fs = require('fs');

const file = 'MPS2605-1.xlsx';
const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
const masterWs = wb.Sheets['MPS'] || wb.Sheets[wb.SheetNames[1]];
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log("Master Plan Row Sample:");
for (let r = 0; r < 100; r++) {
    const row = masterRaw[r];
    if (!row) continue;
    const rowStr = row.join(' | ');
    if (rowStr.includes('LEO') || rowStr.includes('ST38') || rowStr.includes('DST20') || rowStr.includes('지티')) {
        console.log(`Row ${r}: ${rowStr}`);
    }
}
