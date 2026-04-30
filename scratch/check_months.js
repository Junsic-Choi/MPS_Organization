const fs = require('fs');
const { processMpsFile } = require('../extractor.js');

async function check() {
    const filename = process.argv[2] || 'MPS2604-1.xlsx';
    if (!fs.existsSync(filename)) return;
    const buf = fs.readFileSync(filename);
    const result = await processMpsFile(buf);
    
    const monthStats = {};
    result.unusedData.forEach(d => {
        monthStats[d.Month] = (monthStats[d.Month] || 0) + 1;
    });
    
    console.log('Unmapped count by Month:');
    console.log(JSON.stringify(monthStats, null, 2));
}
check();
