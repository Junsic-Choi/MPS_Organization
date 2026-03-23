const fs = require('fs');

console.log('Checking files...');
console.log('Auto_Extract_Final.ps1 exists:', fs.existsSync('Auto_Extract_Final.ps1'));
console.log('data_working.xlsx exists:', fs.existsSync('data_working.xlsx'));
console.log('_FinalList_4650_Latest.csv exists:', fs.existsSync('_FinalList_4650_Latest.csv'));

try {
    const log1 = fs.readFileSync('extraction_full_log.txt', 'utf8');
    console.log('--- extraction_full_log.txt ---');
    console.log(log1.substring(0, 500));
} catch (e) { console.log('No extraction_full_log.txt'); }

try {
    const log2 = fs.readFileSync('final_delivery_log.txt', 'utf16le');
    console.log('--- final_delivery_log.txt ---');
    console.log(log2.substring(0, 500));
} catch (e) { console.log('No final_delivery_log.txt'); }
