const fs = require('fs');
const path = require('path');

const csvPath = path.join(__dirname, '_MPS_Final_Data_v3.csv');
const content = fs.readFileSync(csvPath, 'utf8').replace(/^\ufeff/, '');
const lines = content.split('\n');

console.log('--- Search Results in _MPS_Final_Data_v3.csv ---');
lines.forEach((line, i) => {
    // Search for the site 21 or the name Hutec/휴텍
    if (line.includes('21.') || line.includes('휴텍') || line.includes('LYNX XG')) {
        console.log(`Line ${i+1}: ${line}`);
    }
});
