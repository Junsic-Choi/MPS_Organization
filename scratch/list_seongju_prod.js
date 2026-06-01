const XLSX = require('xlsx');

const wb = XLSX.readFile('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
const prodWs = wb.Sheets[wb.SheetNames[0]];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });

console.log('=== SEONGJU ROWS IN PRODUCTION SHEET ===');
let currentSite = '';
let currentGroup = '';
prodRaw.forEach((row, idx) => {
    if (idx < 2) return;
    const site = String(row[0] || '').trim();
    const group = String(row[1] || '').trim();
    const model = String(row[2] || '').trim();
    const rpm = String(row[3] || '').trim();
    
    if (site) currentSite = site;
    if (group) currentGroup = group;
    
    if (currentSite.includes('성주') && model) {
        // Check if there is plan qty
        let qValues = row.slice(4).map(v => parseInt(v) || 0);
        let sumQ = qValues.reduce((a, b) => a + b, 0);
        if (sumQ > 0) {
            console.log(`Row ${idx+1}: Site="${currentSite}", Group="${currentGroup}", Model="${model}", RPM="${rpm}", QtySum=${sumQ}`);
        }
    }
});
