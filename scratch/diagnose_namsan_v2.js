const { processMpsFile, getMatchKey } = require('../extractor');
const fs = require('fs');

const mpsFile = 'MPS2604-1.xlsx';

console.log('--- Testing getMatchKey Logic ---');
const testModels = ['VCF850LSR', 'VF8LSR', 'SMX2100STB', 'SMX2100SB', 'SMX2100S'];
testModels.forEach(m => {
    console.log(`"${m}" -> "${getMatchKey(m)}"`);
});

console.log('\n--- Processing MPS File ---');
const results = processMpsFile(mpsFile);

console.log('\n--- Analyzing Unused Data for VCF850 / SMX ---');
const namsanUnused = results.unusedData.filter(d => d.Site.includes('남산') && (d.Model.includes('VCF850') || d.Model.includes('SMX')));

namsanUnused.forEach(d => {
    console.log(`Unused: [${d.Month}] ${d.Site} | ${d.Model} | Key: ${getMatchKey(d.Model)} | Qty: ${d.Qty}`);
});

console.log('\n--- Checking Master Pool for Keys ---');
const targetKeys = new Set(namsanUnused.map(d => getMatchKey(d.Model)));
console.log('Target Keys in Unused:', Array.from(targetKeys));

const matchingMaster = results.masterPlanPool.filter(m => {
    const k = getMatchKey(m.Model);
    return targetKeys.has(k);
});

if (matchingMaster.length > 0) {
    console.log(`Found ${matchingMaster.length} master entries with matching keys:`);
    matchingMaster.forEach(m => {
        console.log(`Master: [${m.Month}] ${m.Group} | ${m.Code} | ${m.Product} | Key: ${getMatchKey(m.Model)} | Qty: ${m.Qty}`);
    });
} else {
    console.log('No matching keys found in Master Plan Pool.');
}
