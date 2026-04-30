const XLSX = require('xlsx');
const file = 'c:/Users/i0215099/Desktop/MPS_UPDATE/MPS2603-1.xlsx';
const wb = XLSX.readFile(file);
const ws = wb.Sheets['MPS'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

console.log('Searching for "NHM" in MPS sheet...');
let found = false;
data.forEach((row, i) => {
    const rowStr = JSON.stringify(row);
    if (rowStr.includes('NHM')) {
        console.log(`Found in Row ${i+1}:`, row);
        found = true;
    }
});
if (!found) console.log('Not found in MPS sheet.');
