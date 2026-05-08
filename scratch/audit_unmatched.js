const { processMpsFile } = require('c:/Users/i0215099/Desktop/MPS_UPDATE/extractor');
const fs = require('fs');

async function audit() {
    const filename = 'MPS2604-1.xlsx';
    if (!fs.existsSync(filename)) {
        console.log('File not found at ' + filename);
        return;
    }
    const buffer = fs.readFileSync(filename);
    const result = await processMpsFile(buffer, {});
    
    console.log(`Total Unmatched Items: ${result.unusedData.length}`);
    
    // Group by model to see common failures
    const failures = {};
    result.unusedData.forEach(item => {
        const key = item.Model;
        if (!failures[key]) failures[key] = 0;
        failures[key]++;
    });
    
    console.log('\nTop 10 Unmatched Models:');
    Object.entries(failures)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .forEach(([model, count]) => {
            console.log(`${model}: ${count} rows`);
        });
}

audit();
