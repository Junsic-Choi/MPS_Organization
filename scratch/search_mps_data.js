const XLSX = require('xlsx');

const wb = XLSX.readFile('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
const prodWs = wb.Sheets[wb.SheetNames[0]];
const masterWs = wb.Sheets[wb.SheetNames[1]];

const prodRaw = XLSX.utils.sheet_to_json(prodWs, { header: 1 });
const masterRaw = XLSX.utils.sheet_to_json(masterWs, { header: 1 });

console.log('=== SEARCHING FOR 10206873 or DC3710F or DCM315 in Production sheet ===');
prodRaw.forEach((row, idx) => {
    const rowStr = JSON.stringify(row);
    if (rowStr.includes('10206873') || rowStr.includes('DC3710') || rowStr.includes('DCM315')) {
        console.log(`Prod Row ${idx+1}:`, row);
    }
});

console.log('\n=== SEARCHING FOR 10206873 or DC3710F or DCM315 in Master/MPS sheet ===');
masterRaw.forEach((row, idx) => {
    const rowStr = JSON.stringify(row);
    if (rowStr.includes('10206873') || rowStr.includes('DC3710') || rowStr.includes('DCM315') || rowStr.includes('10215538')) {
        console.log(`Master Row ${idx+1}:`, row);
    }
});
