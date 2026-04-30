const fs = require('fs');
const path = require('path');

const csvPath = path.join(__dirname, '_MPS_Final_Data_v3.csv');
const content = fs.readFileSync(csvPath, 'utf8').replace(/^\ufeff/, '');
const lines = content.split('\n');

let result = "--- Search Results in _MPS_Final_Data_v3.csv ---\n";
lines.forEach((line, i) => {
    if (line.includes('21.') || line.includes('휴텍') || line.includes('LYNX XG')) {
        result += `Line ${i+1}: ${line}\n`;
    }
});

fs.writeFileSync('debug_hutec_mapping.txt', result);
console.log('Done.');
