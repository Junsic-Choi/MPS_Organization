const XLSX = require('xlsx');
const fs = require('fs');
const { processMpsFile } = require('../extractor');

const log = [];
function print(msg) {
    console.log(msg);
    log.push(msg);
}

const wb = XLSX.readFile('MPS2604-1.xlsx');
print('Sheets in MPS2604-1.xlsx: ' + JSON.stringify(wb.SheetNames));

const result = processMpsFile('MPS2604-1.xlsx');
print('\n--- processMpsFile Result ---');
print('Total Results (rows): ' + result.finalResults.length);

const siteSums = {};
let grandTotal = 0;
result.finalResults.forEach(r => {
    siteSums[r.Site] = (siteSums[r.Site] || 0) + r.Qty;
    grandTotal += r.Qty;
});

print('Site sums from extractor:');
print(JSON.stringify(siteSums, null, 2));
print('Grand Total from extractor: ' + grandTotal);

// Detail of each Site
const detailBySite = {};
result.finalResults.forEach(r => {
    if (!detailBySite[r.Site]) detailBySite[r.Site] = [];
    detailBySite[r.Site].push(r);
});

// Write monthly summary
const monthlySiteSums = {};
result.finalResults.forEach(r => {
    if (!monthlySiteSums[r.Site]) monthlySiteSums[r.Site] = {};
    monthlySiteSums[r.Site][r.Month] = (monthlySiteSums[r.Site][r.Month] || 0) + r.Qty;
});
print('\n--- Monthly site sums ---');
print(JSON.stringify(monthlySiteSums, null, 2));

// Let's audit what sheets and what rows are parsed.
const masterSheetName = wb.SheetNames.find(n => n.includes('MPS') || n.includes('Master') || n.includes('마스터')) || wb.SheetNames[1];
const ws = wb.Sheets[masterSheetName];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

print('\n--- Master Sheet Header Row Finder ---');
for (let i = 0; i < Math.min(25, raw.length); i++) {
    print(`Row ${i}: ` + JSON.stringify((raw[i] || []).slice(0, 35).map(c => String(c || '').trim())));
}

fs.writeFileSync('scratch/analyze_2604_out.txt', log.join('\n'));
