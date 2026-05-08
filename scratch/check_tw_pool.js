const { processMpsFile } = require('../extractor');
const res = processMpsFile('MPS2604-1.xlsx');

console.log('--- Unmatched Master Plan Pool Items (TW Series) ---');
res.masterPlanPool.forEach(m => {
    if (m.Qty > 0 && m.Product.includes('TW')) {
        console.log(`Month[${m.Month}] Qty[${m.Qty}] Product[${m.Product}]`);
    }
});
