const { getMatchKey } = require('../extractor');

const testCases = [
    { name: 'Field: MYNX 7500/50 II', val: 'MYNX7500/50 II' },
    { name: 'Master: M75502', val: 'M75502' },
    { name: 'Field: MYNX 5400/40 II', val: 'MYNX5400/40 II' },
    { name: 'Master: M54402', val: 'M54402' },
    { name: 'Field: MYNX 6500/50 II', val: 'MYNX6500/50 II' },
    { name: 'Master: M65502', val: 'M65502' }
];

console.log('--- Testing getMatchKey for MYNX ---');
testCases.forEach(tc => {
    const key = getMatchKey(tc.val);
    console.log(`${tc.name.padEnd(25)} -> Key: ${key}`);
});
