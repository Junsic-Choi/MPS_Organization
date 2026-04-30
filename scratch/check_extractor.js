const fs = require('fs');
const { processMpsFile } = require('../extractor.js');

async function check() {
    const filename = process.argv[2] || 'MPS2604-1.xlsx';
    if (!fs.existsSync(filename)) return;
    const buf = fs.readFileSync(filename);
    const result = await processMpsFile(buf);
    console.log(`Final Results: ${result.finalResults.length}`);
    console.log(`Unmapped (unusedData): ${result.unusedData.length}`);
    
    // Group by model in unusedData
    const counts = {};
    result.unusedData.forEach(d => {
        const k = d.ProductName;
        counts[k] = (counts[k] || 0) + 1;
    });
    
    const sorted = Object.entries(counts).sort((a,b) => b[1] - a[1]);
    console.log('--- Unmapped Samples ---');
    sorted.slice(0, 50).forEach(([k, v]) => {
        console.log(`${k}: ${v}`);
    });
}
check();
