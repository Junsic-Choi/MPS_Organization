const XLSX = require('xlsx');

const wb = XLSX.readFile('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
const masterWs = wb.Sheets[wb.SheetNames[1]];
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log('=== SEONGJU (1842) ROWS IN MASTER SHEET ===');
masterRaw.forEach((row, idx) => {
    if (idx < 5) return;
    const pl = row[1];
    const group = row[2];
    const model = row[3];
    const product = row[4];
    const site = row[6];
    const ver = row[7];
    if (site == '1842') {
        // Let's see if there is any plan quantity
        let hasPlan = false;
        for (let col = 8; col < row.length; col++) {
            if (parseInt(row[col]) > 0) {
                hasPlan = true;
                break;
            }
        }
        if (hasPlan) {
            console.log(`Row ${idx+1}: PL="${pl}", Group="${group}", Model="${model}", Product="${product}", Site="${site}", Ver="${ver}"`);
        }
    }
});
