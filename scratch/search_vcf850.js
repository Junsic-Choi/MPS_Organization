const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const prodWs = wb.Sheets['생산배포용'];
const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });

prodRaw.forEach((row, idx) => {
    const rowStr = row.join(' | ');
    if (rowStr.includes('VCF850') || rowStr.includes('VF8')) {
        console.log(`Row ${idx}:`, row.slice(0, 10));
    }
});
