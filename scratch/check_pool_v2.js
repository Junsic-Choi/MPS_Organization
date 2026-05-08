const { processMpsFile } = require('../extractor');
const res = processMpsFile('MPS2604-1.xlsx');

const targets = ['DVF8000', 'VCF850', 'SMX2100', 'MYNX6500'];
console.log('--- Checking masterPlanPool for targets ---');

targets.forEach(t => {
    const found = res.masterPlanPool.filter(m => m.Product.includes(t) || m.Model.includes(t));
    console.log(`Target[${t}] instances in pool: ${found.length}`);
    if (found.length > 0) {
        found.forEach(f => {
            console.log(`  Month[${f.Month}] Qty[${f.Qty}] Product[${f.Product}]`);
        });
    }
});
