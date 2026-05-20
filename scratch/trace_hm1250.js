const fs = require('fs');
const extractor = require('../extractor.js');

const buf = fs.readFileSync('MPS2605-1.xlsx');
const result = extractor.processMpsFile(buf);

result.finalResults.forEach(r => {
    if (r.Model.includes('HM1250') || r.Product.includes('HM1250')) {
        console.log(`Product=${r.Product}, Model=${r.Model}, Site=${r.Site}, Qty=${r.Qty}, Month=${r.Month}`);
    }
});
