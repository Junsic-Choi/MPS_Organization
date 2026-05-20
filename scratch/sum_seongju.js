const fs = require('fs');
const extractor = require('../extractor.js');

const buf = fs.readFileSync('MPS2605-1.xlsx');
const result = extractor.processMpsFile(buf);

let totalQty = 0;
let seongjuQty = 0;
const siteCounts = {};

result.finalResults.forEach(r => {
    totalQty += r.Qty;
    siteCounts[r.Site] = (siteCounts[r.Site] || 0) + r.Qty;
    if (r.Site === '성주') {
        seongjuQty += r.Qty;
    }
});

console.log('Total Results (entries count):', result.finalResults.length);
console.log('Total Qty overall:', totalQty);
console.log('Seongju Qty:', seongjuQty);
console.log('Site Qty Counts:', siteCounts);
