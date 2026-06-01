const path = require('path');
const { processMpsFile } = require('./extractor.js');

const filePath = path.join(__dirname, 'MPS2605-1.xlsx');
console.log('Running integration test on:', filePath);

const { finalResults } = processMpsFile(filePath);

// Let's perform assertions
let passed = true;
let totalChecked = 0;

function assert(condition, message) {
    totalChecked++;
    if (condition) {
        console.log(`[PASS] ${message}`);
    } else {
        console.error(`[FAIL] ${message}`);
        passed = false;
    }
}

// 1. VF8LSR2 should be classified under '17. VCF 850 Series'
const vf8Lsr2Entries = finalResults.filter(r => r.Model === 'VF8LSR2');
assert(vf8Lsr2Entries.length > 0, `Found ${vf8Lsr2Entries.length} entries for VF8LSR2`);
vf8Lsr2Entries.forEach((entry, idx) => {
    assert(entry.Group === '17. VCF 850 Series', `Entry ${idx + 1} for VF8LSR2 at Site ${entry.Site} is grouped as "${entry.Group}"`);
});

// 2. VF85SR2 should be classified under '17. VCF 850 Series'
const vf85Sr2Entries = finalResults.filter(r => r.Model === 'VF85SR2');
assert(vf85Sr2Entries.length > 0, `Found ${vf85Sr2Entries.length} entries for VF85SR2`);
vf85Sr2Entries.forEach((entry, idx) => {
    assert(entry.Group === '17. VCF 850 Series', `Entry ${idx + 1} for VF85SR2 at Site ${entry.Site} is grouped as "${entry.Group}"`);
});

// 3. DVF8000 should be classified under '11. DVF Series'
const dvf8000Entries = finalResults.filter(r => r.Model === 'DVF8000');
assert(dvf8000Entries.length > 0, `Found ${dvf8000Entries.length} entries for DVF8000`);
dvf8000Entries.forEach((entry, idx) => {
    assert(entry.Group === '11. DVF Series', `Entry ${idx + 1} for DVF8000 at Site ${entry.Site} is grouped as "${entry.Group}"`);
});

// 4. VCF550L should be classified under '17. VCF 5500 Series'
const vcf550lEntries = finalResults.filter(r => r.Model === 'VCF550L');
assert(vcf550lEntries.length > 0, `Found ${vcf550lEntries.length} entries for VCF550L`);
vcf550lEntries.forEach((entry, idx) => {
    assert(entry.Group === '17. VCF 5500 Series', `Entry ${idx + 1} for VCF550L at Site ${entry.Site} is grouped as "${entry.Group}"`);
});

console.log('\n----------------------------------------');
if (passed) {
    console.log(`ALL ${totalChecked} INTEGRATION TESTS PASSED SUCCESSFULLY!`);
} else {
    console.error(`SOME TESTS FAILED. PLEASE AUDIT EXTRACTION RULES.`);
    process.exit(1);
}
