const { getMatchKey } = require('./extractor.js');

const tests = [
    "PUMA VTR1216M",
    "PUMA VTR1216FC",
    "VTR12FC",
    "VTR1620M",
    "VTR2025M"
];

tests.forEach(t => {
    console.log(`"${t}" -> "${getMatchKey(t)}"`);
});
