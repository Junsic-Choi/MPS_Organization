const fs = require('fs');
const extractor = require('../extractor.js');

const buf = fs.readFileSync('MPS2605-1.xlsx');
// We need to override console.log or see what it prints
const result = extractor.processMpsFile(buf);
console.log('Final Results count:', result.finalResults.length);
if (result.finalResults.length > 0) {
    console.log('First 5 finalResults:', result.finalResults.slice(0, 5));
}
