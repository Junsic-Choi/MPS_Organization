const XLSX = require('xlsx');
const fs = require('fs');

const file = 'MPS2605-1.xlsx';
const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
const masterWs = wb.Sheets['MPS'] || wb.Sheets[wb.SheetNames[1]];
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log("Searching for LEO and 지티 in Master Plan...");
masterRaw.forEach((row, idx) => {
    const rowStr = (row || []).join('|');
    if (rowStr.includes('LEO') || rowStr.includes('지티')) {
        console.log(`Row ${idx}: ${rowStr}`);
    }
});

const prodWs = wb.Sheets['생산배포용'] || wb.Sheets[wb.SheetNames[0]];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
console.log("\nSearching for LEO and 지티 in Production Data...");
prodRaw.forEach((row, idx) => {
    const rowStr = (row || []).join('|');
    if (rowStr.includes('LEO') || rowStr.includes('지티')) {
        console.log(`Row ${idx}: ${rowStr}`);
        if (idx < 50) { // Just show a few
             // console.log(row);
        }
    }
});
