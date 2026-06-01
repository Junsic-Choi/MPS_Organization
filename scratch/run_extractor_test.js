const { processMpsFile } = require('../extractor');
const fs = require('fs');

const fileBuffer = fs.readFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
const result = processMpsFile(fileBuffer, {});

console.log('=== EXTRACTED SEONGJU DC/DCM RESULTS ===');
result.finalResults.forEach(r => {
    if (r.Site === '성주' && (r.Product.startsWith('DC') || r.Product.startsWith('DCM'))) {
        console.log(r);
    }
});
