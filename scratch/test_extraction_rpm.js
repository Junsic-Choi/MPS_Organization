const { processMpsFile } = require('../extractor.js');
const path = require('path');

const filename = 'MPS2605-2.xlsx';
const filePath = path.join(__dirname, '..', filename);

console.log('Running test extraction on:', filePath);
const { finalResults } = processMpsFile(filePath);

const targetMonth = '7월';
const targetModel = 'VF8LSR2';

console.log(`\n--- Verification for ${targetModel} in ${targetMonth} ---`);
const matches = finalResults.filter(r => r.Model === targetModel && r.Month === targetMonth);
if (matches.length === 0) {
    console.log('No matches found.');
} else {
    matches.forEach((r, idx) => {
        console.log(`Match ${idx + 1}: Site="${r.Site}", Model="${r.Model}", Product="${r.Product}", RPM="${r.RPM}", Qty=${r.Qty}`);
    });
}
