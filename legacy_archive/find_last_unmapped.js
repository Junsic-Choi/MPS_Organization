const fs = require('fs');
const readline = require('readline');

const stream = fs.createReadStream('_MPS_Final_Data_v3.csv');
const rl = readline.createInterface({ input: stream, crlfDelay: Infinity });

let count = 0;
rl.on('line', (line) => {
    count++;
    if (count === 1) return; // skip header
    const cols = line.split(',').map(c => c.trim().replace(/^"|"$/g, ''));
    const prod = cols[6]; // Product
    if (!prod || prod === 'UNMAPPED') {
        console.log(`[FOUND UNMAPPED] Row ${count}: Model=${cols[2]}, Code=${cols[5]}, Month=${cols[4]}`);
    }
});

rl.on('close', () => {
    console.log('Scan complete.');
});
