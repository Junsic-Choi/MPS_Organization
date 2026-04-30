const fs = require('fs');
const content = fs.readFileSync('_FinalList_utf8.csv', 'utf8');
const lines = content.trim().split('\n');
const header = lines[0];
const dataLines = lines.slice(1);

const monthCounts = {};
dataLines.forEach(line => {
    const parts = line.split(',');
    if (parts.length > 4) {
        const month = parts[4].replace(/"/g, '');
        monthCounts[month] = (monthCounts[month] || 0) + 1;
    }
});

const out = `Total Rows: ${dataLines.length}\nMonth Distribution: ${JSON.stringify(monthCounts, null, 2)}`;
fs.writeFileSync('stat_out.txt', out);
