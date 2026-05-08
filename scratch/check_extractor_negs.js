const { processMpsFile } = require('../extractor');
const fs = require('fs');

async function test() {
    const files = ['MPS2603-1.xlsx', 'MPS2604-1.xlsx'];
    for (const filename of files) {
        if (!fs.existsSync('../' + filename)) continue;
        console.log(`Checking ${filename}...`);
        const buffer = fs.readFileSync('../' + filename);
        const result = await processMpsFile(buffer, {});
        
        const negs = result.masterPlanPool.filter(r => r.Qty < 0);
        if (negs.length > 0) {
            console.log(`Found ${negs.length} negative quantities in ${filename}:`);
            negs.slice(0, 5).forEach(r => console.log(JSON.stringify(r)));
        } else {
            console.log(`No negative quantities found in ${filename}`);
        }
    }
}

test();
