const { getMatchKey } = require('./extractor');

const testCases = [
    { name: 'DVF8000', prod: 'DVF8000/50', master: 'DVF8000' },
    { name: 'DNM750', prod: 'DNM750/50 II', master: 'DNM7550' },
    { name: 'SMX2100', prod: 'SMX2100STB', master: 'SMX2100' },
    { name: 'VCF850', prod: 'VCF850LSR', master: 'VCF850' },
    { name: 'MYNX6500', prod: 'MYNX6500/40', master: 'MYNX6500/40' },
    { name: 'VTR1620', prod: 'VTR1620M', master: 'VTR162' }
];

console.log('--- Key Matching Audit ---');
testCases.forEach(tc => {
    const kProd = getMatchKey(tc.prod);
    const kMaster = getMatchKey(tc.master);
    const match = kProd === kMaster;
    console.log(`[${tc.name}] Prod[${tc.prod}]->${kProd} | Master[${tc.master}]->${kMaster} | MATCH: ${match}`);
});
