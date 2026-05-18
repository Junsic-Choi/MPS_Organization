const { processMpsFile } = require('../extractor');
const fs = require('fs');

async function test() {
    const filename = 'c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-1.xlsx';
    if (!fs.existsSync(filename)) {
        console.log('File not found: ' + filename);
        return;
    }
    const buffer = fs.readFileSync(filename);
    try {
        const result = await processMpsFile(buffer, {});
        console.log('--- Extraction Result ---');
        console.log('Final Results Count:', result.finalResults.length);
        console.log('Unused Data Count:', result.unusedData.length);
        console.log('Master Plan Pool Count:', result.masterPlanPool.length);
        
        if (result.finalResults.length > 0) {
            console.log('First 3 Final Results:', JSON.stringify(result.finalResults.slice(0, 3), null, 2));
        } else {
            console.log('WARNING: No final results extracted!');
        }
    } catch (err) {
        console.error('Error during extraction:', err);
    }
}

test();
