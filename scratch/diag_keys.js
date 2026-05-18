const { getMatchKey, processMpsFile } = require('../extractor');
const XLSX = require('xlsx');

async function debugMatching() {
    const filename = 'MPS2605-1.xlsx';
    const wb = XLSX.readFile(filename);
    const result = processMpsFile(filename);
    
    console.log('--- MPS Master Plan Keys ---');
    const keys = new Set();
    Object.keys(result.masterModelsByGroup).forEach(group => {
        result.masterModelsByGroup[group].forEach(model => {
            const key = getMatchKey(model);
            console.log(`Model: ${model} -> Key: ${key}`);
            keys.add(key);
        });
    });

    console.log('\n--- Production Model Test ---');
    const testModels = ['VCF850LSR', 'VCF850SR', 'DNM750L/50 II', 'PUMA VTR1216M'];
    testModels.forEach(m => {
        console.log(`Prod: ${m} -> Key: ${getMatchKey(m)}`);
    });
}

debugMatching();
