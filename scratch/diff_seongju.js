const fs = require('fs');
const extractor = require('../extractor.js');

const buf = fs.readFileSync('MPS2605-1.xlsx');
const result = extractor.processMpsFile(buf);

console.log('--- ALL SEONGJU ITEMS IN EXTRACTOR ---');
let count = 0;
let extractorSeongjuSum = 0;
result.finalResults.forEach(r => {
    if (r.Site === '성주') {
        extractorSeongjuSum += r.Qty;
        console.log(`Site=성주, Model=${r.Model}, Group=${r.Group}, Product=${r.Product}, Month=${r.Month}, Qty=${r.Qty}`);
    }
});
console.log('Extractor Seongju sum:', extractorSeongjuSum);
