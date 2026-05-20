const fs = require('fs');
const extractor = require('../extractor.js');

const buf = fs.readFileSync('MPS2605-1.xlsx');
const result = extractor.processMpsFile(buf);

const siteSums = {};
let grandTotal = 0;

result.finalResults.forEach(r => {
    const monthVal = parseInt(r.Month.replace('월', ''));
    if (monthVal >= 10) return; // Exclude 10월 to match our 4-9월 counts
    
    siteSums[r.Site] = (siteSums[r.Site] || 0) + r.Qty;
    grandTotal += r.Qty;
});

console.log('Final Site Sums (4~9월) from updated extractor.js:');
console.log(siteSums);
console.log('Grand Total:', grandTotal);
