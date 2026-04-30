const fs = require('fs');
const content = fs.readFileSync('_FinalList_utf8.csv', 'utf8');
const lines = content.trim().split('\n');
const dataLines = lines.slice(1);

let mappedCount = 0;
let unmappedCount = 0;

dataLines.forEach(line => {
    const parts = line.split(',');
    if (parts.length >= 8) {
        const prodName = parts[7].replace(/"/g, '').trim();
        if (prodName !== '') {
            mappedCount++;
        } else {
            unmappedCount++;
        }
    } else {
        unmappedCount++;
    }
});

const out = `Mapped: ${mappedCount}\nUnmapped: ${unmappedCount}\nTotal: ${mappedCount + unmappedCount}`;
fs.writeFileSync('mapped_stat.txt', out);
