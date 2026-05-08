const { processMpsFile } = require('../extractor');
const fs = require('fs');

async function test() {
    const filename = 'MPS2603-1.xlsx';
    if (!fs.existsSync(filename)) {
        console.log('File not found');
        return;
    }
    const buffer = fs.readFileSync(filename);
    const result = await processMpsFile(buffer, {});
    
    console.log('Master Plan Pool Sample (ML0486):');
    const sample = result.masterPlanPool.filter(r => r.Code === 'ML0486');
    sample.forEach(r => console.log(JSON.stringify(r)));
}

test();
