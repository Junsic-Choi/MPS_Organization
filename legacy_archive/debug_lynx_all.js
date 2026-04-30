const fs = require('fs');
const path = require('path');

const csvPath = path.join(__dirname, '_MPS_Final_Data_v3.csv');
const content = fs.readFileSync(csvPath, 'utf8').replace(/^\ufeff/, '');
const lines = content.split('\n');

let result = "--- All LYNX Mappings ---\n";
lines.forEach((line, i) => {
    if (line.includes('LYNX')) {
        result += `Line ${i+1}: ${line}\n`;
    }
});

fs.writeFileSync('debug_lynx_all_mappings.txt', result);
console.log('Done.');
