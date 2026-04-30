const XLSX = require('xlsx');
const path = require('path');

const FILE_PATH = path.join(__dirname, 'MPS2603-1.xlsx');
const workbook = XLSX.readFile(FILE_PATH, { bookSheets: true });
const prodName = workbook.SheetNames.find(n => n.includes('배포'));
const wbFull = XLSX.readFile(FILE_PATH, { sheets: [prodName] });
const prodRaw = XLSX.utils.sheet_to_json(wbFull.Sheets[prodName], { header: 1 });

console.log('--- Scanning Production Sheet for Sites ---');
const sites = new Set();
for (let r = 0; r < prodRaw.length; r++) {
    const site = (prodRaw[r][0] || '').toString().trim();
    if (site && !site.includes('총합계')) sites.add(site);
}
console.log('Available Sites:', Array.from(sites));

console.log('\n--- Searching for Hutec (휴텍) specifically ---');
let lastSite = "";
for (let r = 0; r < prodRaw.length; r++) {
    const row = prodRaw[r] || [];
    if (row[0]) lastSite = row[0].toString().trim();
    if (lastSite.includes('휴텍')) {
        const model = (row[2] || '').toString();
        // Check for 2nd / 3rd month quantities to find active rows
        const val = parseInt(row[4]) || parseInt(row[7]) || 0;
        if (val > 0) {
            console.log(`Row ${r+1}: Site=${lastSite}, Model=${model}, Qty=${val}`);
        }
    }
}
