const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const filename = 'MPS2605-2.xlsx';
const filePath = path.join(__dirname, '..', filename);
if (!fs.existsSync(filePath)) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
}

const wb = XLSX.readFile(filePath);
console.log('Sheet Names:', wb.SheetNames);

const prodSheetName = wb.SheetNames.find(name => ['생산배포', '배포용', 'Production'].some(k => name.includes(k))) || wb.SheetNames[0];
const prodWs = wb.Sheets[prodSheetName];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });

console.log(`\nProd Sheet: "${prodSheetName}" - Row Count: ${prodRaw.length}`);

// Find header row
let headerRowIdx = -1;
for (let r = 0; r < Math.min(20, prodRaw.length); r++) {
    const row = prodRaw[r] || [];
    if (row.some(c => String(c || '').includes('RPM') || String(c || '').includes('기종'))) {
        headerRowIdx = r;
        break;
    }
}

if (headerRowIdx !== -1) {
    console.log('Header Row:', headerRowIdx);
    console.log('Headers:', JSON.stringify(prodRaw[headerRowIdx]));
} else {
    console.log('Could not find header row');
}

// Find VCF850 rows
console.log('\n--- VCF850 / VF8 rows ---');
prodRaw.forEach((row, idx) => {
    const rowStr = JSON.stringify(row);
    if (rowStr.toUpperCase().includes('VCF850') || rowStr.toUpperCase().includes('VF8')) {
        console.log(`Row ${idx}:`, rowStr);
    }
});
