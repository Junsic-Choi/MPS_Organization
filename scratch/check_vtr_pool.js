const { processMpsFile } = require('../extractor');
const res = processMpsFile('MPS2604-1.xlsx');

const twPool = {};
console.log('--- Remaining Quotas for VTR162 and VTR121 ---');
res.masterPlanPool.forEach(m => {
    if (m.Product.includes('VTR162') || m.Product.includes('VTR121')) {
        console.log(`Month[${m.Month}] Qty[${m.Qty}] Product[${m.Product}]`);
    }
});
