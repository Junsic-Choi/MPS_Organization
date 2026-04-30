const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const targetFile = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\일반비_MPS2603-1(생산배포용).xlsx';
const outputFile = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\mps_sheet_dump_verified.txt';

try {
    const workbook = XLSX.readFile(targetFile);
    console.log('Sheet Names:', workbook.SheetNames);
    const sheetName = workbook.SheetNames[3]; // Should be MPS
    console.log('Dumping sheet:', sheetName);
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    let output = "";
    data.forEach((row, idx) => {
        if (idx < 2000) {
            output += row.join('|') + "\n";
        }
    });
    
    fs.writeFileSync(outputFile, output);
    console.log('Successfully written to:', outputFile);
} catch (e) {
    console.error('CRITICAL FAILURE:', e);
}
