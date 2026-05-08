const { processMpsFile } = require('../extractor');
const res = processMpsFile('MPS2604-1.xlsx');

const targets = ['DVF8000', 'VCF850', 'SMX2100', 'MYNX6500'];
console.log('--- Checking quotaPool for targets ---');

targets.forEach(t => {
    // getMatchKey normalization logic
    let key = t.toUpperCase().replace(/[^A-Z1-9]/g, '');
    console.log(`Target[${t}] Key[${key}] in quotaPool?`, !!res.quotaPool[key]);
    if (res.quotaPool[key]) {
        for (let m in res.quotaPool[key]) {
            console.log(`  Month[${m}] items: ${res.quotaPool[key][m].length}`);
        }
    }
});
