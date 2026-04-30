const fs = require('fs');
const pool = JSON.parse(fs.readFileSync('mps_pool_dump.json', 'utf8'));

let out = "--- XG800 Pool Detailed Scan ---\n";
Object.keys(pool).forEach(m => {
    const keys = Object.keys(pool[m]);
    const xgKey = keys.find(k => k === 'XG800');
    if (xgKey) {
        out += `Month: ${m}, Key Found: ${xgKey}, Count: ${pool[m][xgKey].length}\n`;
        pool[m][xgKey].forEach((item, i) => {
            out += `  [${i}] Code=${item.code}, Product=${item.product}\n`;
        });
    } else {
        const fuzzy = keys.filter(k => k.includes('XG'));
        out += `Month: ${m}, Key XG800 NOT FOUND. Similar: ${JSON.stringify(fuzzy)}\n`;
    }
});

fs.writeFileSync('debug_pool_detailed.txt', out);
console.log('Done.');
