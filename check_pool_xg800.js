const fs = require('fs');
const pool = JSON.parse(fs.readFileSync('mps_pool_dump.json', 'utf8'));

console.log('--- XG800 Pool Content Check ---');
Object.keys(pool).forEach(m => {
    const keys = Object.keys(pool[m]);
    const xgKey = keys.find(k => k === 'XG800');
    if (xgKey) {
        console.log(`Month: ${m}, Key Found: ${xgKey}, Count: ${pool[m][xgKey].length}`);
    } else {
        const fuzzy = keys.filter(k => k.includes('XG'));
        console.log(`Month: ${m}, Key XG800 NOT FOUND. Similar: ${JSON.stringify(fuzzy)}`);
    }
});
