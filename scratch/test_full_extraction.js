const fs = require('fs');

const originalCode = fs.readFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\extractor.js', 'utf8');

// Replace the condition in prodRaw.forEach
const modifiedCode = originalCode.replace('if (lastMetaModel && rpm) {', 'if (lastMetaModel) {');

fs.writeFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\scratch\\extractor_relaxed.js', modifiedCode, 'utf8');

const { processMpsFile } = require('./extractor_relaxed');
const fileBuffer = fs.readFileSync('c:\\Users\\i0215099\\Desktop\\MPS_UPDATE\\MPS2605-2.xlsx');
const result = processMpsFile(fileBuffer, {});

console.log('Total Results extracted:', result.finalResults.length);

const seongjuDcm = result.finalResults.filter(r => r.Site === '성주' && (r.Product.startsWith('DC') || r.Product.startsWith('DCM')));
console.log('\n=== EXTRACTED SEONGJU DC/DCM RESULTS ===');
seongjuDcm.forEach(r => {
    console.log(`Month: ${r.Month}, Model: ${r.Model}, Group: ${r.Group}, RPM: "${r.RPM}", Product: "${r.Product}"`);
});

// Let's check if there are other sites or groups that got affected, e.g. check how many have Group: "기타" or Group: "10"
const groupCounts = {};
result.finalResults.forEach(r => {
    groupCounts[r.Group] = (groupCounts[r.Group] || 0) + 1;
});
console.log('\n=== Group Counts ===');
console.log(groupCounts);
