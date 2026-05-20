const XLSX = require('xlsx');

const wb = XLSX.readFile('MPS2605-1.xlsx');
const prodWs = wb.Sheets['생산배포용'];
const masterWs = wb.Sheets['MPS'];

const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log('--- HM1250 in 생산배포용 ---');
prodRaw.forEach((row, idx) => {
    const rowStr = row.join(' | ');
    if (rowStr.includes('HM1250')) {
        console.log(`Row ${idx}:`, row.slice(0, 15));
    }
});

console.log('\n--- HM1250 in MPS ---');
masterRaw.forEach((row, idx) => {
    const rowStr = row.join(' | ');
    if (rowStr.includes('HM1250') || rowStr.includes('HM125')) {
        console.log(`Row ${idx}:`, row.slice(0, 10));
    }
});
