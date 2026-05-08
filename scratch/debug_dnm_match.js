const { processMpsFile, getMatchKey } = require('../extractor');
const res = processMpsFile('MPS2604-1.xlsx');

const targets = ['DNM750/50 II'];
targets.forEach(t => {
    console.log(`\n--- Tracking [${t}] (4월) ---`);
    const prod = res.unusedData.find(u => u.Model === t && u.Month === '4월');
    if (prod) {
        console.log('Production exists in unusedData. Qty:', prod.Qty);
        // Let's check why it didn't match.
        const key = require('./extractor').getMatchKey(t);
        console.log('Match Key:', key);
        // Check quotaPool
        // Wait, I can't access quotaPool easily.
    } else {
        console.log('Not in unusedData (Matched!)');
    }
});
